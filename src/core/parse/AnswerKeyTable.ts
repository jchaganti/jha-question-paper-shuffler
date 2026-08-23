import type { OptionLetter } from '../../shared/types';
import type { Element } from '../docx/dom';
import { NS, childElements, createElement, descendants, isElement, ownerDocumentOf, visibleText } from '../docx/xml';

interface LetterCellRef {
  readonly questionNumber: number;
  readonly cell: Element;
}

const LETTER_RE = /^[A-D]$/;

/**
 * The answer-key grid at the end of the paper.
 *
 * The key is *edited in place*: the tool only rewrites the letter inside each
 * existing cell. Because questions are only ever shuffled inside a subject, the
 * question numbers printed in the key never move, so the generated key keeps the
 * original table's geometry, borders, spans and fonts exactly.
 */
export class AnswerKeyTable {
  private readonly cells = new Map<number, Element>();

  /** Question numbers the key lists more than once - a sign of per-subject numbering. */
  readonly duplicateNumbers: number[] = [];

  private constructor(readonly tables: readonly Element[]) {
    for (const table of tables) {
      for (const ref of AnswerKeyTable.pairsOf(table)) {
        if (this.cells.has(ref.questionNumber)) {
          if (!this.duplicateNumbers.includes(ref.questionNumber)) {
            this.duplicateNumbers.push(ref.questionNumber);
          }
          continue;
        }
        this.cells.set(ref.questionNumber, ref.cell);
      }
    }
    this.duplicateNumbers.sort((a, b) => a - b);
  }

  /**
   * Finds the answer-key table(s) inside `candidates` (the body nodes at or after
   * the "ANSWER KEY" heading, or the whole body when no heading exists).
   */
  static detect(candidates: readonly Element[]): AnswerKeyTable | undefined {
    const tables = candidates.filter((el) => el.namespaceURI === NS.w && el.localName === 'tbl');
    const qualifying = tables.filter((table) => AnswerKeyTable.pairsOf(table).length >= 5);
    if (qualifying.length === 0) return undefined;
    return new AnswerKeyTable(qualifying);
  }

  /** Question numbers present in the key, ascending. */
  get questionNumbers(): number[] {
    return [...this.cells.keys()].sort((a, b) => a - b);
  }

  get size(): number {
    return this.cells.size;
  }

  answerOf(questionNumber: number): OptionLetter | undefined {
    const cell = this.cells.get(questionNumber);
    if (!cell) return undefined;
    const text = visibleText(cell).trim().toUpperCase();
    return LETTER_RE.test(text) ? (text as OptionLetter) : undefined;
  }

  /** Replaces the letter of one question, keeping the cell's formatting. */
  setAnswer(questionNumber: number, letter: OptionLetter): void {
    const cell = this.cells.get(questionNumber);
    if (!cell) throw new Error(`Answer key has no cell for question ${questionNumber}`);
    const textNodes = descendants(cell, NS.w, 't');
    if (textNodes.length === 0) {
      AnswerKeyTable.injectLetter(cell, letter);
      return;
    }
    const first = textNodes[0]!;
    first.textContent = letter;
    for (let i = 1; i < textNodes.length; i++) textNodes[i]!.textContent = '';
  }

  private static injectLetter(cell: Element, letter: OptionLetter): void {
    const doc = ownerDocumentOf(cell);
    let paragraph = childElements(cell, NS.w, 'p')[0];
    if (!paragraph) {
      paragraph = createElement(doc, 'w:p', NS.w);
      cell.appendChild(paragraph);
    }
    const run = createElement(doc, 'w:r', NS.w);
    const text = createElement(doc, 'w:t', NS.w);
    text.appendChild(doc.createTextNode(letter));
    run.appendChild(text);
    paragraph.appendChild(run);
  }

  /** Scans a table for `(question number, answer letter)` cell pairs. */
  private static pairsOf(table: Element): LetterCellRef[] {
    const out: LetterCellRef[] = [];
    for (const row of childElements(table, NS.w, 'tr')) {
      const cells = childElements(row, NS.w, 'tc');
      for (let i = 0; i + 1 < cells.length; i++) {
        const numberText = visibleText(cells[i]!).trim();
        const letterText = visibleText(cells[i + 1]!).trim().toUpperCase();
        if (!/^\d{1,4}$/.test(numberText) || !LETTER_RE.test(letterText)) continue;
        out.push({ questionNumber: Number(numberText), cell: cells[i + 1]! });
        i++; // the letter cell cannot also start a pair
      }
    }
    return out;
  }
}

/** True when `node` is a `w:p` whose only text is the given heading (case-insensitive). */
export function isHeadingParagraph(node: Element, headings: readonly string[]): boolean {
  if (!isElement(node) || node.namespaceURI !== NS.w || node.localName !== 'p') return false;
  const text = visibleText(node).replace(/\s+/g, ' ').trim().toUpperCase();
  return headings.some((heading) => heading.toUpperCase() === text);
}
