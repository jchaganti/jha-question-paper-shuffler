import type { OptionLayoutNote, SkipReason } from '../../shared/types';
import type { Element, Node } from '../docx/dom';
import { NS, childElements, firstChild, replaceChildren } from '../docx/xml';
import type { NumberingIndex } from '../parse/NumberingIndex';
import { paragraphNumId } from '../parse/PaperParser';
import type { QuestionBlock } from '../parse/PaperModel';
import { buildParagraphAtoms, type Atom } from './Atoms';
import { slotSignature } from './OptionBlockParser';
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
        detail: 'No "(A)...(D)" labels and no auto-lettered option list either.',
      };
    }

    const rightSize = lists.filter((list) => list.paragraphs.length === OPTION_COUNT);
    if (rightSize.length === 0) {
      const sizes = lists.map((list) => `numId ${list.numId} has ${list.paragraphs.length}`).join(', ');
      return {
        ok: false,
        reason: 'unexpected-option-count',
        detail: `An auto-lettered list was found but not with four items (${sizes}).`,
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
                `This question has ${rightSize.length} lettered lists of four items; the ` +
                `${candidates[0]!.bracketed ? 'bracketed' : 'upper case'} one was taken as the options.`,
              fix:
                'Letter only the options with a bracketed "(A) (B) (C) (D)" list, and give any other ' +
                'lettered list a different style, such as "A." or "(i) (ii)".',
            },
          ]
        : [];
    if (candidates.length > 1) {
      return {
        ok: false,
        reason: 'ambiguous-option-list',
        detail:
          `More than one four-item lettered list could be the options ` +
          `(numId ${candidates.map((list) => list.numId).join(', ')}).`,
      };
    }

    const paragraphs = candidates[0]!.paragraphs;
    const contents = paragraphs.map((paragraph) => buildParagraphAtoms([paragraph])[0]!.atoms);

    for (let i = 0; i < contents.length; i++) {
      const atoms = contents[i]!;
      if (atoms.every((atom) => atom.blank)) {
        return { ok: false, reason: 'options-not-found', detail: `Auto-lettered option ${i + 1} is empty.` };
      }
      if (atoms.some((atom) => atom.floatingGraphic)) {
        return {
          ok: false,
          reason: 'option-contains-floating-graphic',
          detail: `Auto-lettered option ${i + 1} contains a floating picture, which cannot be moved.`,
        };
      }
    }

    return { ok: true, options: new AutoLetteredOptionSet(paragraphs, contents), notes };
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
    private readonly contents: readonly Atom[][],
  ) {}

  get signatures(): string[] {
    return this.contents.map((atoms) => slotSignature(atoms));
  }

  apply(permutation: readonly number[]): void {
    if (permutation.length !== this.paragraphs.length) {
      throw new Error(
        `Permutation of length ${permutation.length} cannot be applied to ${this.paragraphs.length} options`,
      );
    }

    // Snapshot first: a paragraph's content may be needed after its own has been replaced.
    const snapshots = this.paragraphs.map((paragraph) =>
      childElements(paragraph)
        .filter((child) => !(child.namespaceURI === NS.w && child.localName === 'pPr'))
        .map((child) => child.cloneNode(true)),
    );

    this.paragraphs.forEach((paragraph, index) => {
      const pPr = firstChild(paragraph, NS.w, 'pPr');
      const nodes: Node[] = [];
      // The numbering properties stay, so Word keeps lettering this paragraph in place.
      if (pPr) nodes.push(pPr);
      nodes.push(...(snapshots[permutation[index]!] ?? []));
      replaceChildren(paragraph, nodes);
    });
  }
}

export const AUTO_LETTERED_SKIP_REASONS: readonly SkipReason[] = [
  'unexpected-option-count',
  'ambiguous-option-list',
];
