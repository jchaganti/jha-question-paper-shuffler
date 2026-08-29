import type { OptionLetter } from '../../shared/types';
import type { Element } from '../docx/dom';
import { NS, childElements, createElement, descendants, isElement, ownerDocumentOf, visibleText } from '../docx/xml';

interface LetterCellRef {
  readonly questionNumber: number;
  readonly cell: Element;
}

/** What one table's cells turned out to hold. */
interface TableScan {
  readonly pairs: LetterCellRef[];
  readonly spoiltNumbers: string[];
}

const LETTER_RE = /^[A-D]$/;

/**
 * A question number in the key, with the full stop or bracket authors habitually type
 * after it. "1", "1." and "1)" all name question 1 and nothing else, so all three are
 * read; the paper is not asked to drop a hundred full stops to say what it already says.
 */
const NUMBER_RE = /^(\d{1,4})\s*[.)]?$/;

/**
 * A number cell spoilt by a stray character - "57," where "57." was meant. Unlike a full
 * stop this is a slip, not a convention, so it is reported for correction rather than
 * read through: guessing which characters after a number are decoration and which carry
 * meaning is exactly the kind of guess this tool does not make.
 */
const SPOILT_NUMBER_RE = /^\d{1,4}\s*\S+$/;

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

  /** Number boxes spoilt by a stray character, verbatim - e.g. `["57,"]`. */
  readonly spoiltNumbers: string[] = [];

  private constructor(readonly tables: readonly Element[], scans: readonly TableScan[]) {
    for (const scan of scans) {
      for (const ref of scan.pairs) {
        if (this.cells.has(ref.questionNumber)) {
          if (!this.duplicateNumbers.includes(ref.questionNumber)) {
            this.duplicateNumbers.push(ref.questionNumber);
          }
          continue;
        }
        this.cells.set(ref.questionNumber, ref.cell);
      }
      this.spoiltNumbers.push(...scan.spoiltNumbers);
    }
    this.duplicateNumbers.sort((a, b) => a - b);
    // A key is laid out in columns, so scan order is not question order. Report the boxes
    // in the order the author will look for them.
    this.spoiltNumbers.sort((a, b) => Number.parseInt(a, 10) - Number.parseInt(b, 10));
  }

  /**
   * Finds the answer-key table(s) inside `candidates` (the body nodes at or after
   * the "ANSWER KEY" heading, or the whole body when no heading exists).
   */
  static detect(candidates: readonly Element[]): AnswerKeyTable | undefined {
    const tables = candidates.filter((el) => el.namespaceURI === NS.w && el.localName === 'tbl');
    const scanned = tables.map((table) => ({ table, scan: AnswerKeyTable.scan(table) }));
    const qualifying = scanned.filter((entry) => entry.scan.pairs.length >= 5);
    if (qualifying.length === 0) return undefined;
    return new AnswerKeyTable(
      qualifying.map((entry) => entry.table),
      qualifying.map((entry) => entry.scan),
    );
  }

  /**
   * Question numbers missing from an otherwise unbroken run - the key covers 1 to 100 but
   * has no box for 57. Names the question whose entry needs correcting.
   */
  get gaps(): number[] {
    const numbers = this.questionNumbers;
    const first = numbers[0];
    const last = numbers[numbers.length - 1];
    if (first === undefined || last === undefined) return [];
    const present = new Set(numbers);
    const out: number[] = [];
    for (let n = first; n <= last; n++) if (!present.has(n)) out.push(n);
    return out;
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
  private static scan(table: Element): TableScan {
    const pairs: LetterCellRef[] = [];
    const spoiltNumbers: string[] = [];
    for (const row of childElements(table, NS.w, 'tr')) {
      const cells = childElements(row, NS.w, 'tc');
      for (let i = 0; i + 1 < cells.length; i++) {
        const letterText = visibleText(cells[i + 1]!).trim().toUpperCase();
        if (!LETTER_RE.test(letterText)) continue;
        const numberText = visibleText(cells[i]!).trim();
        const match = NUMBER_RE.exec(numberText);
        if (match) {
          pairs.push({ questionNumber: Number(match[1]), cell: cells[i + 1]! });
          i++; // the letter cell cannot also start a pair
          continue;
        }
        // Next to an answer letter, so it was meant as a question number.
        if (SPOILT_NUMBER_RE.test(numberText)) spoiltNumbers.push(numberText);
      }
    }
    return { pairs, spoiltNumbers };
  }
}

/** True when `node` is a `w:p` whose only text is the given heading (case-insensitive). */
export function isHeadingParagraph(node: Element, headings: readonly string[]): boolean {
  if (!isElement(node) || node.namespaceURI !== NS.w || node.localName !== 'p') return false;
  const text = visibleText(node).replace(/\s+/g, ' ').trim().toUpperCase();
  return headings.some((heading) => heading.toUpperCase() === text);
}
