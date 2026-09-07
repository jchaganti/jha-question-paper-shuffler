import type { OptionLayoutNote, SkipReason } from '../../shared/types';
import type { Element, Node } from '../docx/dom';
import { NS, firstChild, replaceChildren } from '../docx/xml';
import type { NumberingIndex } from '../parse/NumberingIndex';
import { paragraphNumId } from '../parse/PaperParser';
import type { QuestionBlock } from '../parse/PaperModel';
import { buildParagraphAtoms, splitLeadCorePad, type Atom } from './Atoms';
import { slotSignature } from './OptionBlockParser';
import { mergeAdjacentRuns } from './RunMerger';
import type { IOptionSetParser, OptionSet, OptionSetResult } from './OptionSet';

/** Word numbering formats that produce option letters. */
const LETTER_FORMATS = ['upperLetter', 'lowerLetter'];
/** `(%1)` or `[%1]` - how an *option* list is written, as opposed to `%1.` for statements. */
const BRACKETED = /[([{]\s*%1\s*[)\]}]/;
const OPTION_COUNT = 4;

/**
 * Handles options that Word letters automatically: four consecutive list paragraphs whose
 * "(A)" ... "(D)" are generated from the numbering definition rather than typed.
 *
 * Because the letter comes from the paragraph's position in the list, shuffling means
 * moving the *content* between the four paragraphs and leaving each paragraph's numbering
 * properties untouched - Word then re-letters them (A) to (D) in place.
 */
export class AutoLetteredOptionParser implements IOptionSetParser {
  constructor(private readonly numbering: NumberingIndex) {}

  parse(block: QuestionBlock): OptionSetResult {
    const lists = this.letterListsIn(block);
    if (lists.length === 0) {
      return {
        ok: false,
        reason: 'options-not-found',
        detail:
          'No option labels were found under this question - neither labels typed in by hand nor ' +
          'labels put there by Word.',
        fix:
          'Check in Word that this question really has four options and that each one starts ' +
          'with its own label. The surest way is to select the four options and click the ' +
          'numbering button on the Home tab, choosing the "(A) (B) (C) (D)" style.',
      };
    }

    const rightSize = lists.filter((list) => list.paragraphs.length === OPTION_COUNT);
    if (rightSize.length === 0) {
      const sizes = lists.map((list) => list.paragraphs.length).join(' and ');
      return {
        ok: false,
        reason: 'unexpected-option-count',
        detail:
          `Word is putting the labels on this question's options, but it is labelling ${sizes} ` +
          'lines, not four. Every question needs exactly four options, or the answer key stops ' +
          'meaning what it says.',
        fix:
          'Count the options in Word. If one is missing, type it. If an extra line has been ' +
          'caught up in the list by mistake, click that line and switch the numbering off for it.',
      };
    }

    // A question can hold both a lettered statement list ("A. ...") and the option list
    // ("(A) ..."). Options are the bracketed one; upper case wins over lower case.
    const candidates = narrow(rightSize);
    // Preferring one list over another is a judgement, so say so even when it worked.
    const notes: OptionLayoutNote[] =
      rightSize.length > 1 && candidates.length === 1
        ? [
            {
              issue: 'several-lettered-lists',
              detail:
                `This question has ${rightSize.length} lettered lists of four lines. The tool used ` +
                `the ${candidates[0]!.bracketed ? 'one in brackets' : 'one in capitals'} as the options.`,
              fix:
                'Use the "(A) (B) (C) (D)" style for the options only. Give any other list in the ' +
                'question a different style, such as "1." or "(i) (ii)".',
            },
          ]
        : [];
    if (candidates.length > 1) {
      return {
        ok: false,
        reason: 'ambiguous-option-list',
        detail:
          `This question has ${candidates.length} lettered lists of four lines, both written the ` +
          'same way, so the tool cannot tell which one holds the answers. The other is probably a ' +
          'list of statements the question asks about.',
        fix:
          'Use the "(A) (B) (C) (D)" style for the answer options only. Give the other list a ' +
          'different style - "1." or "(i) (ii) (iii) (iv)" - so the two cannot be confused.',
      };
    }

    const paragraphs = candidates[0]!.paragraphs;
    // Split each option paragraph the same way a typed-label option is split: a floating
    // picture anchored in it, and the tabs around the answer, belong to the *position* and
    // stay put; only the answer itself moves to another letter.
    const slots = paragraphs.map((paragraph) => splitLeadCorePad(buildParagraphAtoms([paragraph])[0]!.atoms));
    let anchoredPictures = 0;

    for (let i = 0; i < slots.length; i++) {
      const { lead, core, pad } = slots[i]!;
      anchoredPictures += [...lead, ...pad].filter((atom) => atom.floatingGraphic).length;

      if (core.length === 0) {
        const floating = [...lead, ...pad].some((atom) => atom.floatingGraphic);
        return {
          ok: false,
          reason: floating ? 'option-contains-floating-graphic' : 'options-not-found',
          detail: floating
            ? `The answers to this question are pictures, and option ${'ABCD'[i] ?? i + 1} has no ` +
              'text of its own. Each picture floats: it is positioned against the page rather than ' +
              'sitting in the line, so the tool cannot tell which picture belongs to which option.'
            : `Word has put a label on this line for option ${'ABCD'[i] ?? i + 1}, but the line is empty.`,
          fix: floating
            ? 'Right-click each of the four pictures in turn, choose Wrap Text, then In Line with ' +
              'Text. Each picture then sits in the line of its own option and travels with it.'
            : `Type the answer for option ${'ABCD'[i] ?? i + 1} on that line, or delete the empty line.`,
        };
      }
      // Only a picture wedged between the words of one answer is fatal: the words would
      // move to another letter and the picture, placed from the page, would stay behind.
      if (core.some((atom) => atom.floatingGraphic)) {
        return {
          ok: false,
          reason: 'option-contains-floating-graphic',
          detail:
            `Option ${'ABCD'[i] ?? i + 1} has a floating picture in the middle of its words. A ` +
            'floating picture is placed from the page, so it would stay behind while the words moved.',
          fix:
            'Click that picture, open Layout Options and choose "In line with text". It then sits ' +
            'in the line like a letter does, and travels with its option.',
        };
      }
    }

    const allNotes = [...notes];
    if (anchoredPictures > 0) {
      allNotes.push({
        issue: 'floating-picture-in-option-area',
        detail:
          `${anchoredPictures} floating picture(s) are anchored among the options. The options ` +
          'were shuffled and each picture was left exactly where it is, because a floating ' +
          'picture is placed from the page and does not travel with the text beside it.',
        fix:
          'Check that no picture was meant to belong to one particular option. If one was, ' +
          'select it and set Layout Options to "In line with text" so it moves with its option.',
      });
    }

    return { ok: true, options: new AutoLetteredOptionSet(paragraphs, slots), notes: allNotes };
  }

  private letterListsIn(block: QuestionBlock): LetterList[] {
    const grouped = new Map<string, Element[]>();
    for (const node of block.nodes) {
      if (node === block.questionParagraph) continue;
      const numId = paragraphNumId(node);
      if (!numId) continue;
      const definition = this.numbering.get(numId);
      if (!definition || !LETTER_FORMATS.includes(definition.format)) continue;
      const list = grouped.get(numId) ?? [];
      list.push(node);
      grouped.set(numId, list);
    }

    return [...grouped.entries()].map(([numId, paragraphs]) => {
      const definition = this.numbering.get(numId);
      return {
        numId,
        paragraphs,
        bracketed: BRACKETED.test(definition?.levelText ?? ''),
        upperCase: definition?.format === 'upperLetter',
      };
    });
  }
}

interface LetterList {
  readonly numId: string;
  readonly paragraphs: Element[];
  readonly bracketed: boolean;
  readonly upperCase: boolean;
}

/** Prefers "(A)" over "A.", and upper case over lower case, but only when that discriminates. */
function narrow(lists: readonly LetterList[]): readonly LetterList[] {
  for (const prefer of [(list: LetterList) => list.bracketed, (list: LetterList) => list.upperCase]) {
    if (lists.length <= 1) break;
    const preferred = lists.filter(prefer);
    if (preferred.length > 0 && preferred.length < lists.length) lists = preferred;
  }
  return lists;
}

class AutoLetteredOptionSet implements OptionSet {
  readonly layout = 'auto-lettered' as const;

  constructor(
    private readonly paragraphs: readonly Element[],
    private readonly slots: readonly { lead: Atom[]; core: Atom[]; pad: Atom[] }[],
  ) {}

  /**
   * The fingerprint is the *moving* part only. A floating picture anchored in one of these
   * paragraphs stays behind, so counting it would make every slot look changed after a
   * shuffle and the verifier would report a fault that is not there.
   */
  get signatures(): string[] {
    return this.slots.map((slot) => slotSignature(slot.core));
  }

  apply(permutation: readonly number[]): void {
    if (permutation.length !== this.paragraphs.length) {
      throw new Error(
        `Permutation of length ${permutation.length} cannot be applied to ${this.paragraphs.length} options`,
      );
    }

    this.paragraphs.forEach((paragraph, index) => {
      const source = this.slots[permutation[index]!];
      if (!source) throw new Error(`Permutation refers to option ${permutation[index]}, which does not exist`);
      const own = this.slots[index]!;
      const pPr = firstChild(paragraph, NS.w, 'pPr');
      const nodes: Node[] = [];
      // The numbering properties stay, so Word keeps lettering this paragraph in place.
      if (pPr) nodes.push(pPr);
      // This paragraph's own tabs and anchored pictures; the answer comes from the source.
      for (const atom of own.lead) nodes.push(atom.node);
      for (const atom of source.core) nodes.push(atom.node);
      for (const atom of own.pad) nodes.push(atom.node);
      replaceChildren(paragraph, mergeAdjacentRuns(nodes));
    });
  }
}

export const AUTO_LETTERED_SKIP_REASONS: readonly SkipReason[] = [
  'unexpected-option-count',
  'ambiguous-option-list',
];
