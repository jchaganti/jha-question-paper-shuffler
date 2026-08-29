import {
  OPTION_LETTERS,
  type OptionLayoutNote,
  type OptionLetter,
  type SkipReason,
} from '../../shared/types';
import type { Element } from '../docx/dom';
import { NS, isElement, visibleText } from '../docx/xml';
import type { NumberingIndex } from '../parse/NumberingIndex';
import { paragraphNumId } from '../parse/PaperParser';
import type { QuestionBlock } from '../parse/PaperModel';
import {
  buildParagraphAtoms,
  splitAtomsAtOffsets,
  splitLeadCorePad,
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
  | { readonly ok: false; readonly reason: SkipReason; readonly detail: string };

/**
 * Labels are detected up to (E) even though only four are supported, so that a
 * five-option question is *refused* instead of having its "(E) ..." text silently
 * folded into option (D) and moved around with it.
 */
const UPPER_LABEL_RE = /\(([A-E])\)|([A-E])\)/g;
const LOWER_LABEL_RE = /\(([a-e])\)|([a-e])\)/g;
const ANY_LABEL_RE = /\(([A-Ea-e])\)|([A-Ea-e])\)/g;

interface LabelHit {
  readonly paragraphIndex: number;
  readonly letter: OptionLetter;
  readonly start: number;
  readonly end: number;
  /** The label as typed: `(A)` upper case, `(a)` lower case. */
  readonly upperCase: boolean;
  /** True when the label starts a paragraph or follows a tab - the unambiguous case. */
  readonly afterTab: boolean;
  /** True when the opening bracket is in the text, i.e. the label read as `(A)` not `A)`. */
  readonly bracketed: boolean;
}

/** How hard the parser is allowed to look for labels, from safest to most permissive. */
interface SearchPass {
  readonly pattern: RegExp;
  readonly relaxed: boolean;
  readonly description: string;
}

const SEARCH_PASSES: readonly SearchPass[] = [
  { pattern: UPPER_LABEL_RE, relaxed: false, description: 'upper case, after a tab' },
  { pattern: UPPER_LABEL_RE, relaxed: true, description: 'upper case, after punctuation' },
  { pattern: LOWER_LABEL_RE, relaxed: false, description: 'lower case, after a tab' },
  { pattern: LOWER_LABEL_RE, relaxed: true, description: 'lower case, after punctuation' },
  { pattern: ANY_LABEL_RE, relaxed: false, description: 'either case, after a tab' },
  { pattern: ANY_LABEL_RE, relaxed: true, description: 'either case, after punctuation' },
];

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
 *
 * Whatever cannot be read this way is reported and left untouched - the parser never guesses.
 */
export class OptionBlockParser {
  constructor(private readonly numbering?: NumberingIndex) {}

  parse(block: QuestionBlock): OptionParseResult {
    const paragraphs = block.nodes.filter(
      (node) => node !== block.questionParagraph && node.namespaceURI === NS.w && node.localName === 'p',
    );

    if (paragraphs.length === 0) {
      return { ok: false, reason: 'options-not-found', detail: 'No paragraph follows the question stem.' };
    }

    const baseAtoms = buildParagraphAtoms(paragraphs);

    for (const pass of SEARCH_PASSES) {
      const found = findLabels(baseAtoms, pass);
      const sequence = found.map((hit) => hit.letter).join('');

      if (sequence === 'ABCD') return this.build(paragraphs, found, undefined);
      if (sequence === 'BCD') {
        const lettered = this.letteredFirstOption(paragraphs, found[0]!);
        if (lettered !== undefined) return this.build(paragraphs, found, lettered);
      }
    }

    return this.explainFailure(block, baseAtoms);
  }

  /**
   * Finds the paragraph that holds option (A) when Word letters it: a lettered list with a
   * single item, at or before the paragraph that carries the typed "(B)".
   */
  private letteredFirstOption(paragraphs: readonly Element[], labelB: LabelHit): number | undefined {
    if (!this.numbering) return undefined;

    const counts = new Map<string, number>();
    for (const paragraph of paragraphs) {
      const numId = paragraphNumId(paragraph);
      if (numId) counts.set(numId, (counts.get(numId) ?? 0) + 1);
    }

    const candidates: { index: number; bracketed: boolean }[] = [];
    paragraphs.forEach((paragraph, index) => {
      if (index > labelB.paragraphIndex) return;
      const numId = paragraphNumId(paragraph);
      if (!numId || counts.get(numId) !== 1) return;
      const definition = this.numbering!.get(numId);
      if (!definition || !['upperLetter', 'lowerLetter'].includes(definition.format)) return;
      candidates.push({ index, bracketed: /[([{]\s*%1\s*[)\]}]/.test(definition.levelText) });
    });

    if (candidates.length === 0) return undefined;
    const bracketed = candidates.filter((candidate) => candidate.bracketed);
    const narrowed = bracketed.length > 0 ? bracketed : candidates;
    return narrowed.length === 1 ? narrowed[0]!.index : undefined;
  }

  /**
   * Describes what was odd about a layout the parser nevertheless managed to read.
   *
   * Derived from the labels themselves rather than from which pass matched, so the note
   * says what is actually wrong with the document: a paper written entirely in lower case
   * `(a)...(d)` is consistent and gets no note, even though it needs a later pass.
   */
  private notesFor(found: readonly LabelHit[], letteredFirstParagraph: number | undefined): OptionLayoutNote[] {
    const notes: OptionLayoutNote[] = [];

    if (letteredFirstParagraph !== undefined) {
      notes.push({
        issue: 'mixed-auto-and-typed-labels',
        detail:
          'Option (A) is lettered by Word, but ' +
          found.map((hit) => `(${hit.letter})`).join(', ') +
          ' are typed into the text.',
        fix: 'Letter all four options the same way: either let Word letter all four, or type all four labels.',
      });
    }

    const noTab = found.filter((hit) => !hit.afterTab);
    if (noTab.length > 0) {
      notes.push({
        issue: 'label-not-after-tab',
        detail:
          `No tab in front of ${noTab.length === 1 ? 'label' : 'labels'} ` +
          `${noTab.map((hit) => `(${hit.letter})`).join(', ')}, so ${noTab.length === 1 ? 'it was' : 'they were'} ` +
          'read from the punctuation before it instead.',
        fix: 'Press Tab before each option label, so the label always follows a tab.',
      });
    }

    if (found.length > 1 && found.some((hit) => hit.upperCase !== found[0]!.upperCase)) {
      notes.push({
        issue: 'mixed-label-case',
        detail:
          'The labels mix upper and lower case: ' +
          found.map((hit) => `(${hit.upperCase ? hit.letter : hit.letter.toLowerCase()})`).join(' ') +
          '.',
        fix: 'Use the same case for all four labels - (A) (B) (C) (D).',
      });
    }

    return notes;
  }

  /** Turns a set of labels (plus an optional Word-lettered option A) into option slots. */
  private build(
    paragraphs: readonly Element[],
    found: readonly LabelHit[],
    letteredFirstParagraph: number | undefined,
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
      return { ok: false, reason: 'unexpected-label-sequence', detail: 'Could not align option labels with runs.' };
    }
    const labelRanges = (located as { from: number; to: number }[]).map((range, i) => ({
      ...range,
      bracketed: found[i]!.bracketed,
    }));

    // A Word-lettered option (A) has no label atoms; its content starts at its paragraph.
    if (letteredFirstParagraph !== undefined) {
      const start = firstAtomIndexOf(paragraphAtoms, letteredFirstParagraph);
      if (start === undefined) {
        return { ok: false, reason: 'options-not-found', detail: 'Option (A) has no content.' };
      }
      labelRanges.unshift({ from: start, to: start, bracketed: true });
    }

    const symbolLabel = symbolBeforeLabel(flat, labelRanges);
    if (symbolLabel !== undefined) {
      return {
        ok: false,
        reason: 'label-bracket-is-a-symbol',
        detail:
          `Option label (${symbolLabel}) follows a character inserted from a symbol font, which is ` +
          'how an opening bracket typed with Insert > Symbol appears. That character carries no ' +
          'readable text, so where the option before it ends cannot be established.',
      };
    }

    const slots: OptionSlot[] = [];
    for (let i = 0; i < OPTION_LETTERS.length; i++) {
      const range = labelRanges[i]!;
      const contentEnd = i + 1 < OPTION_LETTERS.length ? labelRanges[i + 1]!.from : flat.length;
      const content = flat.slice(range.to, contentEnd);

      // Tabs, spacer paragraphs and the spaces that separate one option from the next
      // label are layout: they belong to this slot and stay put.
      const { lead: leadAtoms, core: coreAtoms, pad: padAtoms } = splitLeadCorePad(content);
      const labelAtoms = flat.slice(range.from, range.to);
      const paragraphIndex =
        labelAtoms[0]?.paragraphIndex ?? leadAtoms[0]?.paragraphIndex ?? coreAtoms[0]?.paragraphIndex ?? 0;
      const letter = OPTION_LETTERS[i]!;

      if (coreAtoms.length === 0) {
        return { ok: false, reason: 'options-not-found', detail: `Option (${letter}) has no content.` };
      }
      if (coreAtoms.some((atom) => atom.paragraphIndex !== paragraphIndex)) {
        return {
          ok: false,
          reason: 'option-spans-paragraphs',
          detail: `Option (${letter}) continues on another paragraph.`,
        };
      }
      if (coreAtoms.some((atom) => atom.floatingGraphic)) {
        return {
          ok: false,
          reason: 'option-contains-floating-graphic',
          detail: `Option (${letter}) contains a floating picture, which cannot be moved.`,
        };
      }

      slots.push({ letter, labelAtoms, leadAtoms, coreAtoms, padAtoms, paragraphIndex });
    }

    return {
      ok: true,
      block: { paragraphs: paragraphAtoms, prefixAtoms: flat.slice(0, labelRanges[0]!.from), slots },
      notes: this.notesFor(found, letteredFirstParagraph),
    };
  }

  /** Explains, in the user's terms, why none of the passes worked. */
  private explainFailure(block: QuestionBlock, baseAtoms: readonly ParagraphAtoms[]): OptionParseResult {
    const strict = findLabels(baseAtoms, { pattern: ANY_LABEL_RE, relaxed: false, description: '' });

    if (strict.length === 0) {
      const tableHasLabels = block.nodes
        .filter((node) => node.namespaceURI === NS.w && node.localName === 'tbl')
        .some((table) => /\(\s*[ABCDabcd]\s*\)/.test(visibleText(table)));
      if (tableHasLabels) {
        return { ok: false, reason: 'options-inside-table', detail: 'Option labels were found inside a table.' };
      }
      return {
        ok: false,
        reason: 'options-not-found',
        detail: 'No "(A)...(D)" labels found; the options are probably auto-lettered by Word.',
      };
    }

    // A floating picture in the option area is the usual reason a label goes unrecognised:
    // the picture's text box sits in front of it. Say that rather than "unexpected labels".
    const hasFloatingGraphic = baseAtoms.some((entry) => entry.atoms.some((atom) => atom.floatingGraphic));
    const sequence = strict.map((hit) => hit.letter).join('');
    if (hasFloatingGraphic) {
      return {
        ok: false,
        reason: 'option-contains-floating-graphic',
        detail:
          `Labels read as "${sequence}" because a floating picture sits in the option area; ` +
          'its text cannot be told apart from the option labels.',
      };
    }

    return {
      ok: false,
      reason: 'unexpected-label-sequence',
      detail: `Expected labels A,B,C,D but found "${sequence}".`,
    };
  }
}

/**
 * Finds a label whose opening bracket was typed as a symbol-font character.
 *
 * Such a bracket is a `w:sym`: it prints as "(" but holds no text, so reading the
 * paragraph gives "A)" and the bracket looks like the tail of the *previous* option's
 * content. The parser cannot tell the two apart, so the question is refused rather than
 * shuffled with a guessed boundary - moving the option would leave the bracket behind.
 *
 * The missing bracket is the whole signal, so only a label that read as `A)` is suspect. A
 * label that read as `(A)` has its bracket, and a symbol in front of it is the previous
 * option's content - an answer of "∞" or "°C" sitting last in its column, which after a
 * shuffle can land in front of any label.
 *
 * Returns the letter of the offending label, or undefined when every label is plain text.
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
    if (previous >= 0 && flat[previous]!.symbol) return OPTION_LETTERS[i];
  }
  return undefined;
}

function findLabels(paragraphAtoms: readonly ParagraphAtoms[], pass: SearchPass): LabelHit[] {
  const found: LabelHit[] = [];
  paragraphAtoms.forEach((entry, paragraphIndex) => {
    const text = entry.atoms.map((atom) => atom.text).join('');
    pass.pattern.lastIndex = 0;
    let match: RegExpExecArray | null;
    while ((match = pass.pattern.exec(text)) !== null) {
      const raw = match[1] ?? match[2] ?? '';
      const letter = raw.toUpperCase() as OptionLetter;
      if (!isLabelPosition(text, match.index, pass.relaxed)) continue;
      found.push({
        paragraphIndex,
        letter,
        start: match.index,
        end: match.index + match[0].length,
        upperCase: raw === letter,
        afterTab: isLabelPosition(text, match.index, false),
        bracketed: match[1] !== undefined,
      });
    }
  });
  return found;
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
