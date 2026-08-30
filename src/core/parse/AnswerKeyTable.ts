import { OPTION_LETTERS, type OptionLetter } from '../../shared/types';
import type { Element } from '../docx/dom';
import { NS, childElements, createElement, descendants, isElement, ownerDocumentOf, visibleText } from '../docx/xml';

/** The family of symbol a key uses to name one of the four options. */
type AnswerCellScheme = 'letter' | 'digit' | 'roman';

/**
 * How one key box writes its answer, read off that exact box: which family of symbol
 * ("A", "1", "iv"), which case, and whatever decoration surrounds it ("(A)", "1.").
 * Recorded once per box, at parse time, so the new answer can be written back in the same
 * style the author already used - the tool never picks a house style of its own.
 */
interface AnswerCellFormat {
  readonly scheme: AnswerCellScheme;
  readonly upperCase: boolean;
  readonly prefix: string;
  readonly suffix: string;
}

interface LetterCellRef {
  readonly questionNumber: number;
  readonly cell: Element;
  readonly format: AnswerCellFormat;
}

/** What one table's cells turned out to hold. */
interface TableScan {
  readonly pairs: LetterCellRef[];
  readonly spoiltNumbers: string[];
}

/** "i", "ii", "iii", "iv" - the roman numerals a key uses for options A-D, lower case. */
const ROMAN_BY_INDEX = ['i', 'ii', 'iii', 'iv'] as const;

/**
 * Reads one answer box, deducing how *that* box names an option rather than assuming a
 * house style: a bare or decorated letter ("A", "(A)", "a."), a digit naming the option's
 * position ("1", "2)"), or a roman numeral doing the same ("iii)", "(IV)"). Returns the
 * canonical A-D letter plus the exact decoration found, so `setAnswer` can reproduce it.
 */
function parseAnswerCellText(raw: string): { letter: OptionLetter; format: AnswerCellFormat } | undefined {
  const text = raw.trim();
  if (!text) return undefined;

  const bracketed = /^\((.+)\)$/.exec(text);
  const prefix = bracketed ? '(' : '';
  const suffix = bracketed ? ')' : (/[.)]$/.exec(text)?.[0] ?? '');
  const core = (bracketed ? bracketed[1]! : suffix ? text.slice(0, -suffix.length) : text).trim();

  if (/^[A-Da-d]$/.test(core)) {
    return {
      letter: core.toUpperCase() as OptionLetter,
      format: { scheme: 'letter', upperCase: core === core.toUpperCase(), prefix, suffix },
    };
  }
  if (/^[1-4]$/.test(core)) {
    return {
      letter: OPTION_LETTERS[Number(core) - 1]!,
      format: { scheme: 'digit', upperCase: true, prefix, suffix },
    };
  }
  const romanIndex = ROMAN_BY_INDEX.indexOf(core.toLowerCase() as (typeof ROMAN_BY_INDEX)[number]);
  if (romanIndex >= 0) {
    return {
      letter: OPTION_LETTERS[romanIndex]!,
      format: { scheme: 'roman', upperCase: core === core.toUpperCase(), prefix, suffix },
    };
  }
  return undefined;
}

/** Renders `letter` in the exact scheme, case and decoration `format` was read from. */
function renderAnswerCellText(letter: OptionLetter, format: AnswerCellFormat): string {
  const index = OPTION_LETTERS.indexOf(letter);
  const core =
    format.scheme === 'digit'
      ? String(index + 1)
      : format.scheme === 'roman'
        ? format.upperCase
          ? ROMAN_BY_INDEX[index]!.toUpperCase()
          : ROMAN_BY_INDEX[index]!
        : format.upperCase
          ? letter
          : letter.toLowerCase();
  return `${format.prefix}${core}${format.suffix}`;
}

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
  private readonly cells = new Map<number, { cell: Element; format: AnswerCellFormat }>();

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
        this.cells.set(ref.questionNumber, { cell: ref.cell, format: ref.format });
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
    const entry = this.cells.get(questionNumber);
    if (!entry) return undefined;
    return parseAnswerCellText(visibleText(entry.cell).trim())?.letter;
  }

  /**
   * Replaces the answer of one question, keeping the cell's formatting *and* the scheme,
   * case and decoration that box was read with - "(A)" stays "(A)", "iii)" stays a roman
   * numeral, "2" stays a bare digit naming the new option's position.
   */
  setAnswer(questionNumber: number, letter: OptionLetter): void {
    const entry = this.cells.get(questionNumber);
    if (!entry) throw new Error(`Answer key has no cell for question ${questionNumber}`);
    const text = renderAnswerCellText(letter, entry.format);
    const textNodes = descendants(entry.cell, NS.w, 't');
    if (textNodes.length === 0) {
      AnswerKeyTable.injectLetter(entry.cell, text);
      return;
    }
    const first = textNodes[0]!;
    first.textContent = text;
    for (let i = 1; i < textNodes.length; i++) textNodes[i]!.textContent = '';
  }

  private static injectLetter(cell: Element, text: string): void {
    const doc = ownerDocumentOf(cell);
    let paragraph = childElements(cell, NS.w, 'p')[0];
    if (!paragraph) {
      paragraph = createElement(doc, 'w:p', NS.w);
      cell.appendChild(paragraph);
    }
    const run = createElement(doc, 'w:r', NS.w);
    const textEl = createElement(doc, 'w:t', NS.w);
    textEl.appendChild(doc.createTextNode(text));
    run.appendChild(textEl);
    paragraph.appendChild(run);
  }

  /**
   * Scans a table for `(question number, answer)` cell pairs.
   *
   * A key names its options one way throughout, so the scheme used by most of the boxes is
   * taken as this table's scheme and boxes written another way are not read. That is what
   * makes a *key* recognisable rather than any grid of numbers: without it, a table of
   * measurements whose second column happens to hold 1-4 would read as an answer key. A box
   * left out this way shows up as a hole in the key's numbering, which is reported by
   * question number.
   */
  private static scan(table: Element): TableScan {
    const candidates: LetterCellRef[] = [];
    const spoiltNumbers: string[] = [];
    for (const row of childElements(table, NS.w, 'tr')) {
      const cells = childElements(row, NS.w, 'tc');
      for (let i = 0; i + 1 < cells.length; i++) {
        const parsed = parseAnswerCellText(visibleText(cells[i + 1]!).trim());
        if (!parsed) continue;
        const numberText = visibleText(cells[i]!).trim();
        const match = NUMBER_RE.exec(numberText);
        if (match) {
          candidates.push({ questionNumber: Number(match[1]), cell: cells[i + 1]!, format: parsed.format });
          i++; // the answer cell cannot also start a pair
          continue;
        }
        // Next to an answer letter, so it was meant as a question number.
        if (SPOILT_NUMBER_RE.test(numberText)) spoiltNumbers.push(numberText);
      }
    }

    const tally = new Map<AnswerCellScheme, number>();
    for (const ref of candidates) tally.set(ref.format.scheme, (tally.get(ref.format.scheme) ?? 0) + 1);
    const scheme = [...tally.entries()].sort((a, b) => b[1] - a[1])[0]?.[0];
    return { pairs: candidates.filter((ref) => ref.format.scheme === scheme), spoiltNumbers };
  }
}

/** True when `node` is a `w:p` whose only text is the given heading (case-insensitive). */
export function isHeadingParagraph(node: Element, headings: readonly string[]): boolean {
  if (!isElement(node) || node.namespaceURI !== NS.w || node.localName !== 'p') return false;
  const text = visibleText(node).replace(/\s+/g, ' ').trim().toUpperCase();
  return headings.some((heading) => heading.toUpperCase() === text);
}
