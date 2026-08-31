/**
 * Reads the text back out of a report PDF.
 *
 * The report is now a PDF, but the tests still have to prove what it *says* - that the seed
 * is recorded, that answers are quoted in the paper's own vocabulary. The writer emits one
 * uncompressed content stream per page and one text-showing operator per line of a cell, so
 * the text can be recovered by grouping those operators by their y position.
 *
 * This exists to test the real artifact rather than the model behind it: a report that says
 * the right thing but never reaches the file would still pass a model-only assertion.
 */

interface Fragment {
  readonly page: number;
  readonly y: number;
  readonly x: number;
  readonly text: string;
}

const TEXT_OP = /BT [\d.]+ g \/F\d [\d.]+ Tf (-?[\d.]+) (-?[\d.]+) Td \((.*?)\) Tj ET/g;

/**
 * WinAnsi puts Word's smart quotes and dashes in 0x80-0x9F, where Latin-1 has control
 * characters, so those bytes are turned back into the characters they stand for. Every
 * other byte in a string literal means what Latin-1 says it means.
 */
const FROM_WIN_ANSI = new Map<number, string>([
  [0x80, '€'], [0x82, '‚'], [0x83, 'ƒ'], [0x84, '„'], [0x85, '…'],
  [0x86, '†'], [0x87, '‡'], [0x88, 'ˆ'], [0x89, '‰'], [0x8a, 'Š'],
  [0x8b, '‹'], [0x8c, 'Œ'], [0x8e, 'Ž'], [0x91, '‘'], [0x92, '’'],
  [0x93, '“'], [0x94, '”'], [0x95, '•'], [0x96, '–'], [0x97, '—'],
  [0x98, '˜'], [0x99, '™'], [0x9a, 'š'], [0x9b, '›'], [0x9c, 'œ'],
  [0x9e, 'ž'], [0x9f, 'Ÿ'],
]);

function unescape(literal: string): string {
  return [...literal.replace(/\\([()\\])/g, '$1')]
    .map((character) => FROM_WIN_ANSI.get(character.charCodeAt(0)) ?? character)
    .join('');
}

function fragments(pdf: Buffer): Fragment[] {
  const raw = pdf.toString('latin1');
  const streams = [...raw.matchAll(/<< \/Length \d+ >>\nstream\n([\s\S]*?)\nendstream/g)];
  const out: Fragment[] = [];
  streams.forEach((stream, page) => {
    for (const match of stream[1]!.matchAll(TEXT_OP)) {
      out.push({ page, x: Number(match[1]), y: Number(match[2]), text: unescape(match[3]!) });
    }
  });
  return out;
}

/**
 * Every line of the document, in reading order, with the pieces of one line joined by a
 * single space - so a `facts` line reads "Seed march-batch" and a table row reads
 * "28 63 B A A->D, B->A, C->B, D->C".
 */
export function pdfLines(pdf: Buffer): string[] {
  const byLine = new Map<string, Fragment[]>();
  for (const fragment of fragments(pdf)) {
    // y descends down the page, so the key sorts by page and then down the page.
    const key = `${String(fragment.page).padStart(4, '0')}:${String(10000 - Math.round(fragment.y)).padStart(6, '0')}`;
    byLine.set(key, [...(byLine.get(key) ?? []), fragment]);
  }
  return [...byLine.entries()]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([, pieces]) =>
      pieces
        .sort((a, b) => a.x - b.x)
        .map((piece) => piece.text)
        .join(' ')
        .trim(),
    );
}

/** The whole report as one string, for a plain "does it mention this" assertion. */
export function pdfText(pdf: Buffer): string {
  return pdfLines(pdf).join('\n');
}

export function pdfPageCount(pdf: Buffer): number {
  return [...pdf.toString('latin1').matchAll(/\/Type \/Page[^s]/g)].length;
}
