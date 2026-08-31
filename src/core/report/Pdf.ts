/**
 * A very small PDF writer - enough to typeset the generation report and nothing more.
 *
 * Why hand-written rather than a library or Electron's `printToPDF`:
 *
 *  - **No new dependency.** The project ships two (`jszip`, `@xmldom/xmldom`) and a PDF
 *    made of text and rules needs no more than the format's own primitives.
 *  - **The CLI has no Electron.** Every regression run goes through `src/cli`, which runs
 *    on plain Node; `printToPDF` would have made the report an app-only feature and left
 *    the CLI writing something different.
 *  - **Deterministic.** The bytes depend only on the report's content, so the same run
 *    produces the same file - which is the promise the seed already makes about the sets.
 *
 * Only the three standard Type 1 fonts are used (Helvetica, Helvetica-Bold, Courier), so
 * nothing has to be embedded: every PDF reader has them. The price is WinAnsi encoding -
 * see `encode` for what happens to a character outside it.
 */

export type PdfFont = 'regular' | 'bold' | 'mono';

/** Widths of characters 32-126 in 1/1000 em, from the Adobe font metrics. */
const WIDTHS: Record<PdfFont, readonly number[]> = {
  regular: [
    278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556,
    556, 556, 556, 556, 556, 556, 278, 278, 584, 584, 584, 556, 1015, 667, 667, 722, 722, 667, 611, 778,
    722, 278, 500, 667, 556, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 278,
    278, 278, 469, 556, 333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833, 556, 556,
    556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584,
  ],
  bold: [
    278, 333, 474, 556, 556, 889, 722, 238, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556,
    556, 556, 556, 556, 556, 556, 333, 333, 584, 584, 584, 611, 975, 722, 722, 722, 722, 667, 611, 778,
    722, 278, 556, 722, 611, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 333,
    278, 333, 584, 556, 333, 556, 611, 556, 611, 556, 333, 611, 611, 278, 278, 556, 278, 889, 611, 611,
    611, 611, 389, 556, 333, 611, 556, 778, 556, 556, 500, 389, 280, 389, 584,
  ],
  mono: [],
};

/** Courier is monospaced, and anything outside the metric table gets an average width. */
const MONO_WIDTH = 600;
const FALLBACK_WIDTH: Record<PdfFont, number> = { regular: 556, bold: 611, mono: MONO_WIDTH };

const BASE_FONTS: Record<PdfFont, string> = {
  regular: 'Helvetica',
  bold: 'Helvetica-Bold',
  mono: 'Courier',
};
const FONT_KEYS: Record<PdfFont, string> = { regular: 'F1', bold: 'F2', mono: 'F3' };

/**
 * Characters WinAnsiEncoding places outside Latin-1, written the way authors' documents
 * actually produce them (Word's smart quotes and dashes).
 */
const WIN_ANSI_EXTRAS = new Map<number, number>([
  [0x20ac, 0x80], [0x201a, 0x82], [0x0192, 0x83], [0x201e, 0x84], [0x2026, 0x85],
  [0x2020, 0x86], [0x2021, 0x87], [0x02c6, 0x88], [0x2030, 0x89], [0x0160, 0x8a],
  [0x2039, 0x8b], [0x0152, 0x8c], [0x017d, 0x8e], [0x2018, 0x91], [0x2019, 0x92],
  [0x201c, 0x93], [0x201d, 0x94], [0x2022, 0x95], [0x2013, 0x96], [0x2014, 0x97],
  [0x02dc, 0x98], [0x2122, 0x99], [0x0161, 0x9a], [0x203a, 0x9b], [0x0153, 0x9c],
  [0x017e, 0x9e], [0x0178, 0x9f],
]);

/** Width of `text` at `size` points, measured on the bytes that will actually be drawn. */
export function textWidth(text: string, font: PdfFont, size: number): number {
  const bytes = winAnsi(text);
  if (font === 'mono') return (bytes.length * MONO_WIDTH * size) / 1000;
  const table = WIDTHS[font];
  let total = 0;
  for (const code of bytes) {
    total += code >= 32 && code <= 126 ? table[code - 32]! : FALLBACK_WIDTH[font];
  }
  return (total * size) / 1000;
}

/** The longest prefix of `text` that fits in `limit` points, cut at a space when it can be. */
export function wrap(text: string, font: PdfFont, size: number, limit: number): string[] {
  const lines: string[] = [];
  for (const paragraph of text.split('\n')) {
    let line = '';
    for (const word of paragraph.split(' ')) {
      const candidate = line ? `${line} ${word}` : word;
      if (line && textWidth(candidate, font, size) > limit) {
        lines.push(line);
        line = word;
      } else {
        line = candidate;
      }
      // A single word too wide for the column (a long file path) is broken by character,
      // because leaving it to overflow would print it across the neighbouring column.
      while (textWidth(line, font, size) > limit && line.length > 1) {
        let cut = line.length - 1;
        while (cut > 1 && textWidth(line.slice(0, cut), font, size) > limit) cut--;
        lines.push(line.slice(0, cut));
        line = line.slice(cut);
      }
    }
    lines.push(line);
  }
  return lines;
}

/**
 * Characters with an obvious spelling in the printable range. The option mapping is
 * written `A→C`, and an audit trail that printed it as `A?C` would be worse than useless.
 */
const TRANSLITERATE = new Map<number, string>([
  [0x2192, '->'], [0x2190, '<-'], [0x21d2, '=>'],
  [0x00d7, 'x'], [0x2212, '-'], [0x2260, '!='], [0x2264, '<='], [0x2265, '>='],
  [0x00a0, ' '], [0x200b, ''], [0x2009, ' '], [0x00ad, '-'],
]);

/**
 * The WinAnsi bytes a string will be drawn as.
 *
 * A character the encoding has no room for and no spelling in `TRANSLITERATE` becomes `?`.
 * The report quotes option text from the paper, which can hold anything at all; showing
 * `?` states plainly that a character could not be printed, where embedding a Unicode font
 * would mean shipping and subsetting one for an audit trail nobody typesets from.
 *
 * Measuring and drawing both go through here, so a character that spells out as two prints
 * inside the column width that was reserved for it.
 */
function winAnsi(text: string): number[] {
  const bytes: number[] = [];
  for (const character of text) {
    const point = character.codePointAt(0)!;
    const spelled = TRANSLITERATE.get(point);
    if (spelled !== undefined) {
      for (const replacement of spelled) bytes.push(replacement.charCodeAt(0));
      continue;
    }
    bytes.push(point <= 0xff ? point : (WIN_ANSI_EXTRAS.get(point) ?? 0x3f));
  }
  return bytes;
}

/** Turns text into a PDF string literal. */
function encode(text: string): Buffer {
  const bytes: number[] = [0x28]; // (
  for (const code of winAnsi(text)) {
    if (code === 0x28 || code === 0x29 || code === 0x5c) bytes.push(0x5c); // ( ) \
    bytes.push(code);
  }
  bytes.push(0x29); // )
  return Buffer.from(bytes);
}

const number = (value: number): string => (Math.round(value * 100) / 100).toString();

export interface PdfOptions {
  /** Page size in points; A4 by default. */
  readonly width?: number;
  readonly height?: number;
}

/**
 * Draws pages and serialises them.
 *
 * All coordinates are measured **from the top left**, which is how the layout above thinks;
 * PDF's own origin is the bottom left, and the conversion happens here so it happens once.
 */
export class PdfBuilder {
  readonly width: number;
  readonly height: number;
  private readonly pages: Buffer[][] = [];
  private current: Buffer[] = [];

  constructor(options: PdfOptions = {}) {
    this.width = options.width ?? 595.28;
    this.height = options.height ?? 841.89;
    this.addPage();
  }

  get pageCount(): number {
    return this.pages.length;
  }

  addPage(): void {
    this.current = [];
    this.pages.push(this.current);
  }

  /**
   * Draws on a page laid out earlier. Used for the footer, whose "page 3 of 11" cannot be
   * written until the last page exists.
   */
  onPage(index: number, draw: () => void): void {
    const resume = this.current;
    this.current = this.pages[index] ?? resume;
    draw();
    this.current = resume;
  }

  text(x: number, top: number, value: string, font: PdfFont, size: number, gray = 0): void {
    if (value === '') return;
    this.current.push(
      Buffer.concat([
        Buffer.from(`BT ${number(gray)} g /${FONT_KEYS[font]} ${number(size)} Tf ${number(x)} ${number(this.height - top - size)} Td `),
        encode(value),
        Buffer.from(' Tj ET\n'),
      ]),
    );
  }

  /** A horizontal rule, the only kind the report draws. */
  rule(x: number, top: number, width: number, thickness: number, gray: number): void {
    this.current.push(
      Buffer.from(
        `${number(gray)} G ${number(thickness)} w ${number(x)} ${number(this.height - top)} m ` +
          `${number(x + width)} ${number(this.height - top)} l S\n`,
      ),
    );
  }

  fill(x: number, top: number, width: number, height: number, gray: number): void {
    this.current.push(
      Buffer.from(
        `${number(gray)} g ${number(x)} ${number(this.height - top - height)} ${number(width)} ${number(height)} re f\n`,
      ),
    );
  }

  toBuffer(): Buffer {
    const objects: Buffer[] = [];
    const add = (body: string | Buffer): number => {
      objects.push(typeof body === 'string' ? Buffer.from(body) : body);
      return objects.length; // 1-based object number
    };

    // Object 1 is the catalog and object 2 the page tree; both are known only once the
    // page objects have numbers, so they are reserved here and filled in below.
    add('');
    add('');
    const fontIds = (['regular', 'bold', 'mono'] as const).map((font) =>
      add(`<< /Type /Font /Subtype /Type1 /BaseFont /${BASE_FONTS[font]} /Encoding /WinAnsiEncoding >>`),
    );
    const resources =
      `<< /Font << ${(['regular', 'bold', 'mono'] as const)
        .map((font, index) => `/${FONT_KEYS[font]} ${fontIds[index]} 0 R`)
        .join(' ')} >> >>`;

    const pageIds: number[] = [];
    for (const page of this.pages) {
      const content = Buffer.concat(page);
      const streamId = add(
        Buffer.concat([
          Buffer.from(`<< /Length ${content.length} >>\nstream\n`),
          content,
          Buffer.from('\nendstream'),
        ]),
      );
      pageIds.push(
        add(
          `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${number(this.width)} ${number(this.height)}] ` +
            `/Resources ${resources} /Contents ${streamId} 0 R >>`,
        ),
      );
    }

    objects[0] = Buffer.from('<< /Type /Catalog /Pages 2 0 R >>');
    objects[1] = Buffer.from(
      `<< /Type /Pages /Kids [${pageIds.map((id) => `${id} 0 R`).join(' ')}] /Count ${pageIds.length} >>`,
    );

    const parts: Buffer[] = [Buffer.from('%PDF-1.4\n')];
    let offset = parts[0]!.length;
    const offsets: number[] = [];
    objects.forEach((body, index) => {
      offsets.push(offset);
      const chunk = Buffer.concat([Buffer.from(`${index + 1} 0 obj\n`), body, Buffer.from('\nendobj\n')]);
      parts.push(chunk);
      offset += chunk.length;
    });

    const xref = [
      `xref\n0 ${objects.length + 1}\n`,
      '0000000000 65535 f \n',
      ...offsets.map((value) => `${String(value).padStart(10, '0')} 00000 n \n`),
    ].join('');
    parts.push(Buffer.from(xref));
    parts.push(
      Buffer.from(`trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${offset}\n%%EOF\n`),
    );
    return Buffer.concat(parts);
  }
}
