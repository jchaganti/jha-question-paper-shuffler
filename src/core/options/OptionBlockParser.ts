import {
  OPTION_LETTERS,
  type OptionLayoutNote,
  type OptionLetter,
  type SkipReason,
} from '../../shared/types';
import type { AnswerScheme } from '../../shared/answerStyle';
import type { Element } from '../docx/dom';
import { NS, isElement, visibleText } from '../docx/xml';
import type { NumberingIndex } from '../parse/NumberingIndex';
import { paragraphNumId } from '../parse/PaperParser';
import type { QuestionBlock } from '../parse/PaperModel';
import {
  buildParagraphAtoms,
  splitAtomsAtOffsets,
  splitLeadCorePad,
  splitTrailingBlock,
  type Atom,
  type ParagraphAtoms,
} from './Atoms';

export interface OptionSlot {
  readonly letter: OptionLetter;
  /** The "(A)" label itself - never moved. Empty when Word letters this option. */
  readonly labelAtoms: Atom[];
  /** Tabs between the label and the answer text - never moved, so columns stay aligned. */
  readonly leadAtoms: Atom[];
  /** The answer text/equation - this is what gets permuted. */
  readonly coreAtoms: Atom[];
  /** Trailing tabs and blank paragraphs used for column alignment - never moved. */
  readonly padAtoms: Atom[];
  /** Paragraph (inside the option region) that hosts this slot. */
  readonly paragraphIndex: number;
}

export interface OptionBlock {
  readonly paragraphs: ParagraphAtoms[];
  /** Atoms before the first option, e.g. a "1 2 3 4" header row - never moved. */
  readonly prefixAtoms: Atom[];
  readonly slots: readonly OptionSlot[];
}

export type OptionParseResult =
  | { readonly ok: true; readonly block: OptionBlock; readonly notes: readonly OptionLayoutNote[] }
  | {
      readonly ok: false;
      readonly reason: SkipReason;
      /** What is wrong, naming the option it is wrong with. */
      readonly detail: string;
      /** What to change in Word so that this question can be shuffled. */
      readonly fix: string;
    };

/**
 * The ways a paper labels its four options. A paper picks one and uses it throughout, so
 * the reader deduces which one from the question in front of it rather than assuming a
 * house style: letters `(A)`, digits `(1)`, or roman numerals `(i)`.
 *
 * Every scheme carries a fifth token even though only four options are supported, so that
 * a five-option question is *refused* instead of having its fifth option silently folded
 * into the fourth and moved around with it.
 */
type LabelScheme = 'upperLetter' | 'lowerLetter' | 'anyLetter' | 'digit' | 'lowerRoman' | 'upperRoman';

/** Maps a label token as typed to the option slot it names, 0-based. */
const SCHEME_TOKENS: Record<LabelScheme, readonly string[]> = {
  upperLetter: ['A', 'B', 'C', 'D', 'E'],
  lowerLetter: ['a', 'b', 'c', 'd', 'e'],
  anyLetter: ['A', 'B', 'C', 'D', 'E'],
  digit: ['1', '2', '3', '4', '5'],
  lowerRoman: ['i', 'ii', 'iii', 'iv', 'v'],
  upperRoman: ['I', 'II', 'III', 'IV', 'V'],
};

/** `(X)` or `X)`. Roman alternatives are longest-first so "iv)" never reads as "i". */
const SCHEME_PATTERNS: Record<LabelScheme, RegExp> = {
  upperLetter: /\(([A-E])\)|([A-E])\)/g,
  lowerLetter: /\(([a-e])\)|([a-e])\)/g,
  anyLetter: /\(([A-Ea-e])\)|([A-Ea-e])\)/g,
  digit: /\(([1-5])\)|([1-5])\)/g,
  lowerRoman: /\((iii|iv|ii|i|v)\)|(iii|iv|ii|i|v)\)/g,
  upperRoman: /\((III|IV|II|I|V)\)|(III|IV|II|I|V)\)/g,
};

/** The slot a label token names, or undefined when the token belongs to another scheme. */
function slotOf(scheme: LabelScheme, raw: string): number | undefined {
  const token = scheme === 'anyLetter' ? raw.toUpperCase() : raw;
  const index = SCHEME_TOKENS[scheme].indexOf(token);
  return index >= 0 ? index : undefined;
}

const LETTER_SCHEMES: readonly LabelScheme[] = ['upperLetter', 'lowerLetter', 'anyLetter'];

/**
 * The one slot missing from three labels that are otherwise A, B, C, D in order - 2 for
 * `A,B,D`, 0 for `B,C,D`. Undefined unless exactly one is missing and the rest are in order,
 * so a repeated or out-of-order run is never read as a gap.
 */
function missingSlot(found: readonly LabelHit[]): number | undefined {
  if (found.length !== OPTION_LETTERS.length - 1) return undefined;
  const letters = found.map((hit) => hit.letter);
  for (let gap = 0; gap < OPTION_LETTERS.length; gap++) {
    const expected = OPTION_LETTERS.filter((_letter, index) => index !== gap);
    if (expected.every((letter, index) => letters[index] === letter)) return gap;
  }
  return undefined;
}

/**
 * Word number formats that write an option name the way each label scheme does, narrowest
 * first: a list in the *same case* as the typed labels before one that merely belongs to the
 * same family.
 *
 * That order is what tells an option apart from a lettered statement list beside it. One
 * paper writes its two situations as an `(a) (b)` list and its option (A) as an `(A)` list;
 * the typed labels are `(B) (C) (D)`, so the upper-case list is the options and the lower
 * case one is not.
 */
const FORMATS_OF_SCHEME: Record<LabelScheme, readonly (readonly string[])[]> = {
  upperLetter: [['upperLetter'], ['upperLetter', 'lowerLetter']],
  lowerLetter: [['lowerLetter'], ['upperLetter', 'lowerLetter']],
  anyLetter: [['upperLetter', 'lowerLetter']],
  digit: [['decimal']],
  lowerRoman: [['lowerRoman'], ['lowerRoman', 'upperRoman']],
  upperRoman: [['upperRoman'], ['lowerRoman', 'upperRoman']],
};

/** A numbering level that prints its number in brackets - `(%1)` - as an option list does. */
const BRACKETED_LEVEL = /[([{]\s*%1\s*[)\]}]/;

/**
 * Lays typed labels and Word-lettered paragraphs out in document order and checks that they
 * make exactly four options, each typed label at the position its letter names.
 *
 * A lettered paragraph sorts *before* any typed label inside it, because Word draws its
 * number at the start of the line: `(A) 1 amp <tab> (B) 1.5 amp` is one paragraph holding
 * option A, lettered by Word, followed by a typed option B.
 *
 * Returns the lettered paragraphs with the slot each one fills, or undefined when the two
 * do not add up - which is the whole safeguard: anything that is not exactly four, in
 * order, is left for the caller to report rather than read.
 */
function merged(
  found: readonly LabelHit[],
  candidates: readonly { paragraphIndex: number; bracketed: boolean }[],
): LetteredOption[] | undefined {
  if (found.length + candidates.length !== OPTION_LETTERS.length) return undefined;

  const entries = [
    ...found.map((hit) => ({ paragraphIndex: hit.paragraphIndex, start: hit.start, hit })),
    ...candidates.map((candidate) => ({ paragraphIndex: candidate.paragraphIndex, start: -1, hit: undefined })),
  ].sort((a, b) => a.paragraphIndex - b.paragraphIndex || a.start - b.start);

  const lettered: LetteredOption[] = [];
  for (const [slot, entry] of entries.entries()) {
    if (entry.hit === undefined) lettered.push({ paragraphIndex: entry.paragraphIndex, slot });
    else if (entry.hit.letter !== OPTION_LETTERS[slot]) return undefined;
  }
  return lettered;
}

/** True when the question holds a table that carries what look like option labels. */
function optionsInsideTable(block: QuestionBlock): boolean {
  return block.nodes
    .filter((node) => node.namespaceURI === NS.w && node.localName === 'tbl')
    .some((table) => /\(\s*([ABCDabcd]|[1-4]|i{1,3}|iv)\s*\)/.test(visibleText(table)));
}

function tableRefusal(): Extract<OptionParseResult, { ok: false }> {
  return {
    ok: false,
    reason: 'options-inside-table',
    detail:
      'The options of this question are in a table. To swap two options the tool moves their ' +
      'text along the line they sit on, and text in a table is not on a line it can move ' +
      'along.',
    fix:
      'Click anywhere in the table. On the Layout tab (the one that appears at the right-hand ' +
      'end when the cursor is in a table), click Convert to Text and choose Tabs. The options ' +
      'will look the same on the page - still lined up in columns - and can then be shuffled.',
  };
}

/** A label written the way the page writes it - `(1)`, `B)`, `(iii)` - for messages. */
function labelAsTyped(hit: { raw: string; bracketed: boolean }): string {
  return hit.bracketed ? `(${hit.raw})` : `${hit.raw})`;
}

interface LabelHit {
  readonly paragraphIndex: number;
  readonly letter: OptionLetter;
  readonly start: number;
  readonly end: number;
  /** The scheme this label was read under - which family of token it is written in. */
  readonly scheme: LabelScheme;
  /** The token exactly as the page writes it - "A", "1", "iii" - for error messages. */
  readonly raw: string;
  /** The label as typed: `(A)` upper case, `(a)` lower case. Letter schemes only. */
  readonly upperCase: boolean;
  /** True when the label starts a paragraph or follows a tab - the unambiguous case. */
  readonly afterTab: boolean;
  /** True when the opening bracket is in the text, i.e. the label read as `(A)` not `A)`. */
  readonly bracketed: boolean;
}

/** The one option Word letters by itself, and the slot its letter belongs to. */
interface LetteredOption {
  readonly paragraphIndex: number;
  readonly slot: number;
}

/** How hard the parser is allowed to look for labels, from safest to most permissive. */
interface SearchPass {
  readonly scheme: LabelScheme;
  readonly relaxed: boolean;
}

/**
 * Tried in order, first clean A,B,C,D wins.
 *
 * Letters come before digits and digits before roman numerals, because that is the order
 * of decreasing certainty that a run of labels really is the *options*. A question often
 * lists items as `(i)...(iv)` or `(a)...(d)` above answers written another way, so the
 * scheme most likely to be the answers is asked first and the most easily confused with an
 * item list is asked last.
 */
const SEARCH_PASSES: readonly SearchPass[] = [
  { scheme: 'upperLetter', relaxed: false },
  { scheme: 'upperLetter', relaxed: true },
  { scheme: 'lowerLetter', relaxed: false },
  { scheme: 'lowerLetter', relaxed: true },
  { scheme: 'anyLetter', relaxed: false },
  { scheme: 'anyLetter', relaxed: true },
  { scheme: 'digit', relaxed: false },
  { scheme: 'digit', relaxed: true },
  { scheme: 'lowerRoman', relaxed: false },
  { scheme: 'lowerRoman', relaxed: true },
  { scheme: 'upperRoman', relaxed: false },
  { scheme: 'upperRoman', relaxed: true },
];

/** Which label schemes write an option name in the same family as the answer key does. */
const SCHEMES_OF_ANSWER: Record<AnswerScheme, readonly LabelScheme[]> = {
  letter: ['upperLetter', 'lowerLetter', 'anyLetter'],
  digit: ['digit'],
  roman: ['lowerRoman', 'upperRoman'],
};

/**
 * The passes to run, with the ones matching the answer key's own scheme first.
 *
 * The key is the paper's own statement of how it names an option, and it settles the case
 * the pass order alone gets wrong: a match-the-columns question in a paper keyed `1 2 3 4`
 * often carries an `(a)...(d)` list of column entries *above* its real `(1)...(4)` options,
 * and reading letters first would take that list for the options. Nothing is guessed - the
 * paper says "my answers are digits", so the digit labels are its options.
 */
function passesFor(answerScheme: AnswerScheme | undefined): readonly SearchPass[] {
  if (!answerScheme) return SEARCH_PASSES;
  const preferred = SCHEMES_OF_ANSWER[answerScheme];
  return [
    ...SEARCH_PASSES.filter((pass) => preferred.includes(pass.scheme)),
    ...SEARCH_PASSES.filter((pass) => !preferred.includes(pass.scheme)),
  ];
}

/**
 * Locates the four option slots of a question that carries typed "(A)" labels.
 *
 * Real papers are inconsistent, so the labels are looked for in several passes, from the
 * safest reading to the most permissive, and the first pass that yields a clean A,B,C,D is
 * used:
 *
 *  1. **Upper case before lower case.** A match-the-columns question often lists its items
 *     as "(a) ... (d)" and then its answers as "(A) ... (D)"; reading upper case first picks
 *     the answers rather than a mixture of the two.
 *  2. **After a tab, then after punctuation.** A label normally starts a paragraph or follows
 *     a tab. Some papers lose that tab ("...(i), (iv) and (v) (D) ..."), so a second pass
 *     accepts a label that follows a space *and* a non-alphanumeric character - which still
 *     refuses "Both (A) and (B)" and "Assertion (A):", where a letter comes first.
 *  3. **A first option lettered by Word.** Some questions have option (A) as an automatically
 *     lettered list paragraph and type only (B), (C) and (D). When exactly one single-item
 *     lettered list sits before label (B), that paragraph is option (A).
 *  4. **The answer key's scheme first.** Whichever family the key names its answers in -
 *     letters, digits or roman numerals - is tried before the others, because the key is the
 *     paper's own statement of how it names an option.
 *
 * Whatever cannot be read this way is reported and left untouched - the parser never guesses.
 */
export class OptionBlockParser {
  constructor(
    private readonly numbering?: NumberingIndex,
    /** How the paper's answer key names an option; the matching labels are tried first. */
    private readonly answerScheme?: AnswerScheme,
  ) {}

  parse(block: QuestionBlock): OptionParseResult {
    const paragraphs = block.nodes.filter(
      (node) => node !== block.questionParagraph && node.namespaceURI === NS.w && node.localName === 'p',
    );

    if (paragraphs.length === 0) {
      // A question whose options are *entirely* inside a table has no option paragraphs at
      // all, so the table has to be looked for here too - saying "nothing follows the stem"
      // about a question with four visible options would send the author hunting for
      // something that is plainly there.
      if (optionsInsideTable(block)) return tableRefusal();
      return {
        ok: false,
        reason: 'options-not-found',
        detail: 'There is nothing at all below this question - no options were found under it.',
        fix:
          'Open the paper in Word and look at this question. Its options have probably ended up ' +
          'under the question above or below it, or have been deleted. Put the four options ' +
          'under this question, each one on its own line.',
      };
    }

    const baseAtoms = buildParagraphAtoms(paragraphs);

    for (const pass of passesFor(this.answerScheme)) {
      const found = findLabels(baseAtoms, pass);
      const sequence = found.map((hit) => hit.letter).join('');

      if (sequence === 'ABCD') return this.build(paragraphs, found, []);

      // Some labels typed, the rest produced by Word's own numbering - which prints "(C)"
      // without any "(C)" in the text. The lettered paragraphs fill the slots the typed
      // labels do not name.
      const lettered = this.letteredOptionsFor(paragraphs, found, pass.scheme);
      if (lettered !== undefined) return this.build(paragraphs, found, lettered);
    }

    return this.explainFailure(block, paragraphs, baseAtoms);
  }

  /**
   * Finds the options Word letters by itself, when the rest are typed - a paper that types
   * `(A)`, `(B)`, `(D)` and lets Word produce the `(C)`, or types only `(D)` and lets Word
   * produce the first three.
   *
   * The reading is a merge, not a search. Every option paragraph is either a typed label or
   * an item of a lettered list; laying both out in document order has to produce exactly
   * four options, with each typed label sitting at the position its own letter names. That
   * is what tells the options apart from a lettered *statement* list inside the question: a
   * question with three lettered statements and four typed labels merges to seven, and is
   * refused rather than read.
   *
   * Only lists whose format matches the scheme being tried count - `(A)` labels against an
   * upper or lower case letter list, `(4)` against a decimal one - so a roman item list is
   * never folded into lettered options. When the merge does not come out at four, it is
   * tried again with only the *bracketed* lists, since `(%1)` is how a paper writes an
   * option and `%1.` is how it writes a statement.
   */
  private letteredOptionsFor(
    paragraphs: readonly Element[],
    found: readonly LabelHit[],
    scheme: LabelScheme,
  ): LetteredOption[] | undefined {
    // With nothing typed there is no anchor: a lettered list on its own is the auto-lettered
    // reader's business, and it checks things this merge cannot.
    if (!this.numbering || found.length === 0 || found.length >= OPTION_LETTERS.length) return undefined;
    // Out of order or repeated labels are a defect in the paper, not a gap to fill.
    if (found.some((hit, i) => i > 0 && hit.letter <= found[i - 1]!.letter)) return undefined;

    const candidates: { paragraphIndex: number; bracketed: boolean; format: string }[] = [];
    paragraphs.forEach((paragraph, index) => {
      const numId = paragraphNumId(paragraph);
      const definition = numId ? this.numbering!.get(numId) : undefined;
      if (!definition) return;
      candidates.push({
        paragraphIndex: index,
        bracketed: BRACKETED_LEVEL.test(definition.levelText),
        format: definition.format,
      });
    });

    // Narrowest reading first, and the first one that accounts for exactly four options
    // wins. Anything wider is only tried because the narrower one found nothing at all, so
    // a list that fits is never passed over for one that merely could.
    for (const formats of FORMATS_OF_SCHEME[scheme]) {
      const family = candidates.filter((candidate) => formats.includes(candidate.format));
      const bracketed = family.filter((candidate) => candidate.bracketed);
      const answer = merged(found, bracketed) ?? merged(found, family);
      if (answer !== undefined) return answer;
    }
    return undefined;
  }

  /**
   * Describes what was odd about a layout the parser nevertheless managed to read.
   *
   * Derived from the labels themselves rather than from which pass matched, so the note
   * says what is actually wrong with the document: a paper written entirely in lower case
   * `(a)...(d)` is consistent and gets no note, even though it needs a later pass.
   */
  private notesFor(
    found: readonly LabelHit[],
    lettered: readonly LetteredOption[],
    anchoredPictures: number,
  ): OptionLayoutNote[] {
    const notes: OptionLayoutNote[] = [];
    const shown = (hit: LabelHit): string => labelAsTyped(hit);

    if (lettered.length > 0) {
      const positions = lettered
        .map((option) => ['first', 'second', 'third', 'fourth'][option.slot] ?? 'first')
        .join(', ');
      const many = lettered.length > 1;
      notes.push({
        issue: 'mixed-auto-and-typed-labels',
        detail:
          `Word is putting the label on the ${positions} option${many ? 's' : ''}, but ` +
          `${found.map(shown).join(', ')} ${found.length > 1 ? 'were' : 'was'} typed in by hand.`,
        fix:
          'Label all four options the same way: either let Word label all four, or type all four ' +
          'labels yourself.',
      });
    }

    const noTab = found.filter((hit) => !hit.afterTab);
    if (noTab.length > 0) {
      notes.push({
        issue: 'label-not-after-tab',
        detail:
          `There is no tab in front of ${noTab.length === 1 ? 'label' : 'labels'} ` +
          `${noTab.map(shown).join(', ')}. The tool had to work out where ` +
          `${noTab.length === 1 ? 'that option starts' : 'those options start'} from the ` +
          'punctuation instead, which is a guess.',
        fix: 'Press the Tab key before each option label, so every label follows a tab.',
      });
    }

    const letterLabels = found.filter((hit) => LETTER_SCHEMES.includes(hit.scheme));
    if (letterLabels.length === found.length && found.length > 1 && found.some((hit) => hit.upperCase !== found[0]!.upperCase)) {
      notes.push({
        issue: 'mixed-label-case',
        detail: `These labels mix capitals and small letters: ${found.map(shown).join(' ')}.`,
        fix: `Use the same case for all four labels - ${SCHEME_TOKENS[found[0]!.scheme].slice(0, 4).map((token) => `(${token})`).join(' ')}.`,
      });
    }

    if (anchoredPictures > 0) {
      notes.push({
        issue: 'floating-picture-in-option-area',
        detail:
          `${anchoredPictures} picture(s) float among the options - they are positioned against ` +
          'the page rather than sitting in a line. The options were shuffled and every picture ' +
          'was left exactly where it was, because a floating picture does not travel with the ' +
          'text beside it.',
        fix:
          'Check that none of these pictures was meant to belong to one particular option. If ' +
          'one was, right-click it, choose Wrap Text, then In Line with Text, so it moves with ' +
          'its option.',
      });
    }

    return notes;
  }

  /** Turns a set of labels (plus an optional Word-lettered option A) into option slots. */
  private build(
    paragraphs: readonly Element[],
    found: readonly LabelHit[],
    lettered: readonly LetteredOption[],
  ): OptionParseResult {
    // Cut atoms so that every label starts and ends on an atom boundary.
    const paragraphAtoms = buildParagraphAtoms(paragraphs).map((entry, paragraphIndex) => {
      const offsets = found
        .filter((hit) => hit.paragraphIndex === paragraphIndex)
        .flatMap((hit) => [hit.start, hit.end]);
      return offsets.length === 0 ? entry : { ...entry, atoms: splitAtomsAtOffsets(entry.atoms, offsets) };
    });

    const flat: Atom[] = paragraphAtoms.flatMap((entry) => entry.atoms);
    const located = found.map((hit) => locateLabel(paragraphAtoms, hit));
    if (located.some((range) => range === undefined)) {
      return {
        ok: false,
        reason: 'unexpected-label-sequence',
        detail:
          'The four option labels are there, but the tool cannot work out where each option ' +
          'begins and ends.',
        fix:
          'Retype the labels of this question: delete each one and type it again, pressing Tab ' +
          'after each label before the answer text.',
      };
    }
    const labelRanges = (located as { from: number; to: number }[]).map((range, i) => ({
      ...range,
      bracketed: found[i]!.bracketed,
    }));

    // How this paper writes the label of slot i - "(1)", "B)", "(iii)" - so that every
    // message below names the option the way the reader will find it on the page.
    const shownLabel = (i: number): string => {
      const hit = found.find((candidate) => candidate.letter === OPTION_LETTERS[i]);
      if (hit) return labelAsTyped(hit);
      const scheme = found[0]?.scheme ?? 'upperLetter';
      return `(${SCHEME_TOKENS[scheme][i] ?? OPTION_LETTERS[i]})`;
    };

    // A Word-lettered option has no label atoms of its own; its content starts at its
    // paragraph. Each is spliced in at the slot whose letter was missing from the text -
    // in ascending slot order, so every splice lands at an index the earlier ones settled.
    for (const option of [...lettered].sort((a, b) => a.slot - b.slot)) {
      const start = firstAtomIndexOf(paragraphAtoms, option.paragraphIndex);
      if (start === undefined) {
        return {
          ok: false,
          reason: 'options-not-found',
          detail: `Word has put a label ${shownLabel(option.slot)} on a line, but the line is empty.`,
          fix: `Type the answer for option ${shownLabel(option.slot)} on that line, or delete the empty line.`,
        };
      }
      labelRanges.splice(option.slot, 0, { from: start, to: start, bracketed: true });
    }

    const symbolLabel = symbolBeforeLabel(flat, labelRanges);
    if (symbolLabel !== undefined) {
      return {
        ok: false,
        reason: 'label-bracket-is-a-symbol',
        detail:
          `The bracket in front of option ${shownLabel(OPTION_LETTERS.indexOf(symbolLabel))} was put in ` +
          'with Insert, Symbol, from a font of little drawings such as Wingdings. In the file it is ' +
          'a drawing, not a bracket, so the tool cannot see where the option before it ends.',
        fix:
          `Delete the label ${shownLabel(OPTION_LETTERS.indexOf(symbolLabel))} and simply type it ` +
          'again on the keyboard.',
      };
    }

    // A paragraph holding nothing but whitespace separates the options from whatever
    // trails them. An empty paragraph has no atoms at all, so "no atoms" counts as blank.
    const blankParagraphs = new Set(
      paragraphAtoms
        .filter((entry) => entry.atoms.every((atom) => atom.blank))
        .map((entry) => entry.paragraphIndex),
    );

    const slots: OptionSlot[] = [];
    let anchoredPictures = 0;
    for (let i = 0; i < OPTION_LETTERS.length; i++) {
      const range = labelRanges[i]!;
      const isLast = i + 1 === OPTION_LETTERS.length;
      const contentEnd = isLast ? flat.length : labelRanges[i + 1]!.from;
      // The last option has no following label to stop it, so it would otherwise swallow
      // everything trailing the option list. A blank paragraph ends the list.
      const { kept: content, trailing } = isLast
        ? splitTrailingBlock(flat.slice(range.to, contentEnd), blankParagraphs)
        : { kept: flat.slice(range.to, contentEnd), trailing: [] as Atom[] };

      // Tabs, spacer paragraphs, the spaces that separate one option from the next label,
      // and any floating picture anchored at either end are layout: they belong to this
      // slot and stay put.
      const split = splitLeadCorePad(content);
      const { lead: leadAtoms, core: coreAtoms } = split;
      // Trailing material goes back to its own paragraphs untouched (padding always does).
      const padAtoms = [...split.pad, ...trailing];
      const labelAtoms = flat.slice(range.from, range.to);
      const paragraphIndex =
        labelAtoms[0]?.paragraphIndex ?? leadAtoms[0]?.paragraphIndex ?? coreAtoms[0]?.paragraphIndex ?? 0;
      const letter = OPTION_LETTERS[i]!;
      anchoredPictures += [...leadAtoms, ...padAtoms].filter((atom) => atom.floatingGraphic).length;

      if (coreAtoms.length === 0) {
        // An option can perfectly well *be* a picture - but only an inline one travels with
        // its option. A floating picture is placed from the page, and which option each one
        // belongs to cannot be read from the file: the pictures and the labels sit in
        // different paragraphs, in an order that differs from question to question. Rather
        // than guess and risk pairing an answer with the wrong picture, say what to change.
        // The pictures are looked for across the whole question, because the one that gives
        // an option its answer is often anchored in a neighbouring paragraph.
        const floating = flat.some((atom) => atom.floatingGraphic);
        return {
          ok: false,
          reason: floating ? 'option-contains-floating-graphic' : 'options-not-found',
          detail: floating
            ? `The answers to this question are pictures, and option ${shownLabel(i)} has no text ` +
              'of its own. Each picture floats: it is positioned against the page rather than ' +
              'sitting in the line, so the tool cannot tell which picture belongs to which option.'
            : `Option ${shownLabel(i)} has a label but nothing written after it.`,
          fix: floating
            ? 'Right-click each of the four pictures in turn, choose Wrap Text, then In Line with ' +
              'Text. Each picture then sits in the line of its own option and travels with it.'
            : `Type the answer for option ${shownLabel(i)} after its label.`,
        };
      }
      if (coreAtoms.some((atom) => atom.paragraphIndex !== paragraphIndex)) {
        return {
          ok: false,
          reason: 'option-spans-paragraphs',
          detail:
            `Option ${shownLabel(i)} carries on to the line below, because Enter was pressed in ` +
            'the middle of it. Word treats what follows Enter as a separate line of its own, and ' +
            'the tool cannot move half an option.',
          fix:
            'Click at the very start of the line that carries on, and press Backspace to join it ' +
            'back up. If you want the option printed on two lines, press Shift+Enter at that ' +
            'point instead of Enter. It looks exactly the same on the page, but Word keeps the ' +
            'option together as one line.',
        };
      }
      // Only a picture wedged *between* words of the answer is fatal: the words would move
      // to another slot and the picture, placed from the page, would stay behind.
      if (coreAtoms.some((atom) => atom.floatingGraphic)) {
        return {
          ok: false,
          reason: 'option-contains-floating-graphic',
          detail:
            `Option ${shownLabel(i)} has a picture in the middle of its words, and that picture ` +
            'floats: it is positioned against the page instead of sitting in the line. If the ' +
            'option moved, the words would go and the picture would stay behind.',
          fix:
            'Right-click that picture, choose Wrap Text, then In Line with Text. The picture then ' +
            'sits in the line like a letter does, and moves with the option.',
        };
      }

      slots.push({ letter, labelAtoms, leadAtoms, coreAtoms, padAtoms, paragraphIndex });
    }

    const prefixAtoms = flat.slice(0, labelRanges[0]!.from);
    anchoredPictures += prefixAtoms.filter((atom) => atom.floatingGraphic).length;

    return {
      ok: true,
      block: { paragraphs: paragraphAtoms, prefixAtoms, slots },
      notes: this.notesFor(found, lettered, anchoredPictures),
    };
  }

  /**
   * Explains, in the user's terms, why none of the passes worked.
   *
   * The report quotes the labels in the scheme that came closest, and in the tokens the
   * page actually writes, so the author is not told about "(A)...(D)" when their paper is
   * written "(1)...(4)".
   */
  private explainFailure(
    block: QuestionBlock,
    paragraphs: readonly Element[],
    baseAtoms: readonly ParagraphAtoms[],
  ): OptionParseResult {
    const best = passesFor(this.answerScheme)
      .filter((pass) => !pass.relaxed)
      .map((pass) => findLabels(baseAtoms, pass))
      .reduce((a, b) => (b.length > a.length ? b : a), [] as LabelHit[]);

    if (best.length === 0) {
      if (optionsInsideTable(block)) return tableRefusal();
      return {
        ok: false,
        reason: 'options-not-found',
        detail:
          'No option labels were found under this question. There is text below it, but nothing ' +
          'that reads as four labelled options.',
        fix:
          'Check in Word that this question really has four options and that each one starts ' +
          'with its own label. The surest way is to select the four options and click the ' +
          'numbering button on the Home tab, choosing the "(A) (B) (C) (D)" style - Word then ' +
          'keeps the labels right for you.',
      };
    }

    // Two options on one line, pushed apart with the space bar instead of a tab. The
    // labels are all there and in order - only the separator in front of them is wrong -
    // so name them and say what to press, rather than reporting the letters that survived.
    const afterSpaces = labelsAfterSpaceRun(baseAtoms, best[0]!.scheme);
    if (afterSpaces.length > 0) {
      const combined = [...best, ...afterSpaces].sort(byDocumentOrder);
      const letters = combined.map((hit) => hit.letter).join('');
      // Accepting the space runs has to account for all four options, or the separator is
      // not the whole story: one letter may be left to a paragraph Word letters itself, but
      // an unexplained gap means something else is wrong and gets its own message below.
      const complete =
        letters === OPTION_LETTERS.join('') ||
        this.letteredOptionsFor(paragraphs, combined, best[0]!.scheme) !== undefined;
      if (complete) {
        const many = afterSpaces.length > 1;
        return {
          ok: false,
          reason: 'label-after-spaces-not-tab',
          detail:
            `Two options share a line here, and label${many ? 's' : ''} ` +
            `${afterSpaces.map(labelAsTyped).join(', ')} ${many ? 'were' : 'was'} pushed across ` +
            'the page by holding down the space bar. To the tool a row of spaces looks like ' +
            'ordinary text, so it cannot see where the option before ends.',
          fix:
            `Delete the spaces in front of ${many ? 'each of these labels' : 'this label'} and ` +
            'press the Tab key once instead. The page will look the same.',
        };
      }
    }

    // A floating picture in the option area is the usual reason a label goes unrecognised:
    // the picture's text box sits in front of it. Say that rather than "unexpected labels".
    const hasFloatingGraphic = baseAtoms.some((entry) => entry.atoms.some((atom) => atom.floatingGraphic));
    const sequence = best.map((hit) => hit.raw).join('');
    if (hasFloatingGraphic) {
      return {
        ok: false,
        reason: 'option-contains-floating-graphic',
        detail:
          `The tool could only read ${sequence.length === 0 ? 'none of the labels' : `"${sequence}"`} ` +
          'here. A picture floats among the options - it is positioned against the page rather ' +
          'than sitting in a line - and any words inside that picture get mixed up with the real ' +
          'option labels.',
        fix:
          'Right-click the picture, choose Wrap Text, then In Line with Text. If the picture is a ' +
          'diagram for the question as a whole rather than for one option, move it above the ' +
          'options instead.',
      };
    }

    const tokens = SCHEME_TOKENS[best[0]!.scheme].slice(0, 4);
    const expected = tokens.map((token) => `(${token})`).join(' ');
    const missing = tokens.filter((token) => !sequence.includes(token));
    return {
      ok: false,
      reason: 'unexpected-label-sequence',
      detail:
        `This question should have the four labels ${expected}. The tool found "${sequence}"` +
        `${missing.length > 0 ? ` and could not find ${missing.map((token) => `(${token})`).join(' or ')}` : ''}` +
        '. Either a label is missing or mistyped, or something else in the question is being ' +
        'mistaken for one.',
      fix:
        `In Word, check that this question has exactly four options labelled ${expected}, each ` +
        'label used once, and each label either at the start of its line or straight after a Tab. ' +
        'If the question also gives the reader a list of statements to consider, label that list ' +
        'differently - "1." or "(i) (ii)" - so it cannot be mistaken for the options.',
    };
  }
}

/**
 * Finds a label whose opening bracket was typed as a character from a font this tool
 * cannot read.
 *
 * Such a bracket is a `w:sym` holding no text of its own, so reading the paragraph gives
 * "A)" and the bracket looks like the tail of the *previous* option's content. The parser
 * cannot tell the two apart, so the question is refused rather than shuffled with a guessed
 * boundary - moving the option would leave the bracket behind.
 *
 * A bracket typed in the Symbol font, which is how Insert > Symbol writes one by default,
 * is not this case: its code is decoded, the label reads as `(A)` like any other, and this
 * check never sees it. What is left is a picture font - Wingdings and its relatives - whose
 * codes name drawings rather than characters.
 *
 * The missing bracket is the whole signal, so only a label that read as `A)` is suspect. A
 * label that read as `(A)` has its bracket, and a symbol in front of it is the previous
 * option's content, which after a shuffle can land in front of any label.
 *
 * Returns the letter of the offending label, or undefined when every label is readable.
 */
function symbolBeforeLabel(
  flat: readonly Atom[],
  labelRanges: readonly { from: number; to: number; bracketed: boolean }[],
): OptionLetter | undefined {
  for (let i = 0; i < labelRanges.length; i++) {
    const range = labelRanges[i]!;
    // An option lettered by Word has no typed label, so there is no bracket to check.
    if (range.to <= range.from || range.bracketed) continue;
    let previous = range.from - 1;
    while (previous >= 0 && flat[previous]!.blank) previous--;
    if (previous >= 0 && flat[previous]!.unreadableSymbol) return OPTION_LETTERS[i];
  }
  return undefined;
}

function findLabels(paragraphAtoms: readonly ParagraphAtoms[], pass: SearchPass): LabelHit[] {
  const found: LabelHit[] = [];
  const pattern = SCHEME_PATTERNS[pass.scheme];
  paragraphAtoms.forEach((entry, paragraphIndex) => {
    const text = entry.atoms.map((atom) => atom.text).join('');
    pattern.lastIndex = 0;
    let match: RegExpExecArray | null;
    while ((match = pattern.exec(text)) !== null) {
      const raw = match[1] ?? match[2] ?? '';
      const slot = slotOf(pass.scheme, raw);
      if (slot === undefined) continue;
      if (!isLabelPosition(text, match.index, pass.relaxed)) continue;
      found.push({
        paragraphIndex,
        // Slots are named A-D throughout the tool whatever the paper writes on the page;
        // the label atoms themselves are never moved, so the page keeps its own tokens.
        letter: OPTION_LETTERS[slot] ?? ('E' as OptionLetter),
        start: match.index,
        end: match.index + match[0].length,
        scheme: pass.scheme,
        raw,
        upperCase: raw === raw.toUpperCase(),
        afterTab: isLabelPosition(text, match.index, false),
        bracketed: match[1] !== undefined,
      });
    }
  });
  return found;
}

/**
 * How wide a run of spaces has to be before it reads as column alignment rather than
 * ordinary prose. One space in front of a label is "Both (A) and (B)"; several is an author
 * pushing the second column across the page because they did not press Tab.
 */
const COLUMN_GAP_SPACES = 2;

/**
 * Labels that would have been read as labels but for what sits in front of them: a run of
 * spaces following a letter or digit. That is the one position both the strict and the
 * relaxed pass refuse, because it is also the shape of "Both (A) and (B)".
 *
 * These are never used to *read* a question - only to explain one that could not be read.
 * On their own they prove nothing; the caller accepts them only when adding them completes
 * the run of four, which is what tells a column gap apart from a sentence.
 */
function labelsAfterSpaceRun(paragraphAtoms: readonly ParagraphAtoms[], scheme: LabelScheme): LabelHit[] {
  const found: LabelHit[] = [];
  const pattern = SCHEME_PATTERNS[scheme];
  paragraphAtoms.forEach((entry, paragraphIndex) => {
    const text = entry.atoms.map((atom) => atom.text).join('');
    pattern.lastIndex = 0;
    let match: RegExpExecArray | null;
    while ((match = pattern.exec(text)) !== null) {
      const raw = match[1] ?? match[2] ?? '';
      const slot = slotOf(scheme, raw);
      if (slot === undefined) continue;
      if (isLabelPosition(text, match.index, true)) continue;
      const spaces = text.slice(0, match.index).length - text.slice(0, match.index).replace(/ +$/, '').length;
      if (spaces < COLUMN_GAP_SPACES) continue;
      found.push({
        paragraphIndex,
        letter: OPTION_LETTERS[slot] ?? ('E' as OptionLetter),
        start: match.index,
        end: match.index + match[0].length,
        scheme,
        raw,
        upperCase: raw === raw.toUpperCase(),
        afterTab: false,
        bracketed: match[1] !== undefined,
      });
    }
  });
  return found;
}

function byDocumentOrder(a: LabelHit, b: LabelHit): number {
  return a.paragraphIndex - b.paragraphIndex || a.start - b.start;
}

/**
 * A label counts at the start of a paragraph or straight after a tab. In a relaxed pass it
 * also counts after a space, provided the character before that space is not alphanumeric -
 * which keeps "Both (A) and (B)" and "Assertion (A):" out.
 */
function isLabelPosition(text: string, index: number, relaxed: boolean): boolean {
  let i = index - 1;
  let sawSpace = false;
  while (i >= 0 && text[i] === ' ') {
    sawSpace = true;
    i--;
  }
  if (i < 0) return true;
  const previous = text[i]!;
  if (previous === '\t' || previous === '\n') return true;
  if (!relaxed || !sawSpace) return false;
  return !/[A-Za-z0-9]/.test(previous);
}

function firstAtomIndexOf(paragraphs: readonly ParagraphAtoms[], paragraphIndex: number): number | undefined {
  let flatIndex = 0;
  for (const entry of paragraphs) {
    if (entry.paragraphIndex === paragraphIndex) return entry.atoms.length > 0 ? flatIndex : undefined;
    flatIndex += entry.atoms.length;
  }
  return undefined;
}

function locateLabel(
  paragraphs: readonly ParagraphAtoms[],
  label: { paragraphIndex: number; start: number; end: number },
): { from: number; to: number } | undefined {
  let flatIndex = 0;
  for (const entry of paragraphs) {
    if (entry.paragraphIndex !== label.paragraphIndex) {
      flatIndex += entry.atoms.length;
      continue;
    }
    let offset = 0;
    let from = -1;
    let to = -1;
    for (let i = 0; i < entry.atoms.length; i++) {
      const atom = entry.atoms[i]!;
      if (offset === label.start) from = flatIndex + i;
      offset += atom.text.length;
      if (offset === label.end) {
        to = flatIndex + i + 1;
        break;
      }
    }
    return from >= 0 && to > from ? { from, to } : undefined;
  }
  return undefined;
}

/**
 * Content fingerprint of one option, used to prove after generation that no answer text
 * was lost, duplicated or mis-mapped.
 *
 * Plain text alone is not enough: "-11" and "11" differ only by a Symbol-font glyph
 * (`w:sym`), and two equations differ only by the object they embed. The fingerprint
 * therefore also carries symbol characters, math text and relationship ids.
 */
export function slotSignature(atoms: readonly Atom[]): string {
  const text = atoms
    .map((atom) => atom.text)
    .join('')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  const extras = atoms.flatMap((atom) => collectDistinguishingTokens(atom.node));
  return `${text}|${extras.sort().join(',')}`;
}

function collectDistinguishingTokens(root: Element): string[] {
  const out: string[] = [];
  const walk = (el: Element): void => {
    for (let i = 0; i < el.attributes.length; i++) {
      const a = el.attributes.item(i);
      if (a && a.namespaceURI === NS.r) out.push(`rel:${a.localName}=${a.value}`);
    }
    if (el.namespaceURI === NS.w && el.localName === 'sym') {
      out.push(`sym:${el.getAttributeNS(NS.w, 'font') ?? ''}:${el.getAttributeNS(NS.w, 'char') ?? ''}`);
    }
    if (el.namespaceURI === NS.m && el.localName === 't') {
      out.push(`math:${(el.textContent ?? '').trim()}`);
    }
    for (let i = 0; i < el.childNodes.length; i++) {
      const node = el.childNodes.item(i);
      if (isElement(node)) walk(node);
    }
  };
  walk(root);
  return out;
}
