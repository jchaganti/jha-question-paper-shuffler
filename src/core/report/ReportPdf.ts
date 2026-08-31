import { PdfBuilder, textWidth, wrap, type PdfFont } from './Pdf';
import type { ReportBlock, ReportColumn, ReportDocument, ReportFact } from './ReportDocument';

/**
 * Typesets a `ReportDocument` onto A4 pages.
 *
 * The report is an audit trail - a mapping table of 200 rows per set, times however many
 * sets - so the layout is plain and dense on purpose: no colour beyond a green or red
 * verdict, rules instead of boxes, and a repeated header row so a table that runs over
 * several pages can still be read.
 */

const MARGIN = 46;
const BOTTOM = 56; // leaves room for the footer rule and page number
const GAP = 6;

const SIZE = {
  title: 19,
  subtitle: 9,
  heading1: 13,
  heading2: 11,
  body: 9,
  table: 8.5,
  footer: 8,
} as const;

const GRAY = {
  text: 0.1,
  muted: 0.42,
  rule: 0.78,
  headerFill: 0.93,
  ok: 0.35,
  bad: 0.2,
} as const;

const LEADING = 1.35;

/** No column is squeezed below this, however long the widest column's prose is. */
const MIN_COLUMN_WIDTH = 42;

const CELL_PAD = 4;
/**
 * Slack added to a column's natural width. Without it, a cell measured at exactly the
 * column's inner width can still fail the `>` test by one floating-point step and wrap -
 * which is how "Subject" once came out as "Subjec" over "t".
 */
const CELL_SLACK = 1;

export function renderReportPdf(document: ReportDocument): Buffer {
  return new ReportLayout(document).run();
}

class ReportLayout {
  private readonly pdf = new PdfBuilder();
  private readonly left = MARGIN;
  private readonly width: number;
  private y = MARGIN;

  constructor(private readonly document: ReportDocument) {
    this.width = this.pdf.width - MARGIN * 2;
  }

  run(): Buffer {
    this.masthead();
    for (const block of this.document.blocks) this.block(block);
    this.footers();
    return this.pdf.toBuffer();
  }

  private masthead(): void {
    this.pdf.text(this.left, this.y, this.document.title, 'bold', SIZE.title, GRAY.text);
    this.y += SIZE.title * LEADING;
    for (const line of wrap(this.document.subtitle, 'regular', SIZE.subtitle, this.width)) {
      this.pdf.text(this.left, this.y, line, 'regular', SIZE.subtitle, GRAY.muted);
      this.y += SIZE.subtitle * LEADING;
    }
    this.y += 4;
    this.pdf.rule(this.left, this.y, this.width, 1, GRAY.rule);
    this.y += GAP * 2;
  }

  private block(block: ReportBlock): void {
    switch (block.kind) {
      case 'heading':
        return this.heading(block.text, block.level ?? 1);
      case 'paragraph':
        return this.paragraph(block.text);
      case 'facts':
        return this.facts(block.items);
      case 'table':
        return this.table(block.columns, block.rows);
    }
  }

  private heading(text: string, level: 1 | 2): void {
    const size = level === 1 ? SIZE.heading1 : SIZE.heading2;
    // A heading alone at the foot of a page is worse than a slightly short page, so it
    // takes its own height plus a line of whatever follows before it will stay.
    this.ensure(size * LEADING + SIZE.body * LEADING + GAP);
    this.y += GAP;
    this.pdf.text(this.left, this.y, text, 'bold', size, GRAY.text);
    this.y += size * LEADING;
    if (level === 1) {
      this.pdf.rule(this.left, this.y - 2, this.width, 0.6, GRAY.rule);
      this.y += 3;
    }
    this.y += 2;
  }

  private paragraph(text: string): void {
    for (const line of wrap(text, 'regular', SIZE.body, this.width)) {
      this.ensure(SIZE.body * LEADING);
      this.pdf.text(this.left, this.y, line, 'regular', SIZE.body, GRAY.text);
      this.y += SIZE.body * LEADING;
    }
    this.y += GAP / 2;
  }

  private facts(items: readonly ReportFact[]): void {
    const labelWidth =
      Math.max(...items.map((item) => textWidth(`${item.label}  `, 'bold', SIZE.body))) + 2;
    for (const item of items) {
      const font: PdfFont = item.mono ? 'mono' : 'regular';
      const gray = item.tone === 'ok' ? GRAY.ok : item.tone === 'bad' ? GRAY.bad : GRAY.text;
      const lines = wrap(item.value, font, SIZE.body, this.width - labelWidth);
      lines.forEach((line, index) => {
        this.ensure(SIZE.body * LEADING);
        // The label prints once; a value that wrapped keeps its indent, so the column of
        // values stays a column.
        if (index === 0) {
          this.pdf.text(this.left, this.y, item.label, 'bold', SIZE.body, GRAY.muted);
        }
        this.pdf.text(this.left + labelWidth, this.y, line, font, SIZE.body, gray);
        this.y += SIZE.body * LEADING;
      });
    }
    this.y += GAP / 2;
  }

  private table(columns: readonly ReportColumn[], rows: readonly (readonly string[])[]): void {
    const widths = this.columnWidths(columns, rows);
    const cellPad = CELL_PAD;
    const rowHeight = (cells: readonly string[]): number =>
      Math.max(
        ...cells.map((cell, index) => {
          const font: PdfFont = columns[index]?.mono ? 'mono' : 'regular';
          return wrap(cell, font, SIZE.table, widths[index]! - cellPad * 2).length;
        }),
        1,
      ) *
        SIZE.table *
        LEADING +
      cellPad;

    const header = columns.map((column) => column.header);
    const headerHeight = rowHeight(header);
    // A table of short values is narrower than the page, and its rules stop with it.
    const tableWidth = widths.reduce((sum, value) => sum + value, 0);

    const drawRow = (cells: readonly string[], height: number, bold: boolean): void => {
      let x = this.left;
      cells.forEach((cell, index) => {
        const column = columns[index];
        const font: PdfFont = bold ? 'bold' : column?.mono ? 'mono' : 'regular';
        const inner = widths[index]! - cellPad * 2;
        let lineY = this.y + cellPad / 2;
        for (const line of wrap(cell, font, SIZE.table, inner)) {
          const offset =
            column?.align === 'right' ? inner - textWidth(line, font, SIZE.table) : 0;
          this.pdf.text(x + cellPad + offset, lineY, line, font, SIZE.table, GRAY.text);
          lineY += SIZE.table * LEADING;
        }
        x += widths[index]!;
      });
      this.y += height;
    };

    const drawHeader = (): void => {
      this.pdf.fill(this.left, this.y, tableWidth, headerHeight, GRAY.headerFill);
      drawRow(header, headerHeight, true);
      this.pdf.rule(this.left, this.y, tableWidth, 0.7, GRAY.rule);
    };

    this.ensure(headerHeight + rowHeight(rows[0] ?? header));
    drawHeader();

    for (const row of rows) {
      const height = rowHeight(row);
      if (this.y + height > this.pdf.height - BOTTOM) {
        this.pdf.addPage();
        this.y = MARGIN;
        drawHeader();
      }
      drawRow(row, height, false);
      this.pdf.rule(this.left, this.y, tableWidth, 0.3, 0.88);
    }
    this.y += GAP;
  }

  /**
   * Every column gets the width its content wants. `weight` marks the columns that carry
   * prose: those take any room left over, and those give the room back when there is not
   * enough. A table of nothing but short values simply ends before the right margin
   * instead of being stretched across the page.
   */
  private columnWidths(
    columns: readonly ReportColumn[],
    rows: readonly (readonly string[])[],
  ): number[] {
    const natural = columns.map((column, index) => {
      const font: PdfFont = column.mono ? 'mono' : 'regular';
      const widest = rows.reduce(
        (max, row) => Math.max(max, textWidth(row[index] ?? '', font, SIZE.table)),
        textWidth(column.header, 'bold', SIZE.table),
      );
      return widest + CELL_PAD * 2 + CELL_SLACK;
    });

    const total = natural.reduce((sum, value) => sum + value, 0);
    if (total <= this.width) {
      const spare = this.width - total;
      const weights = columns.map((column) => column.weight ?? 0);
      const weightTotal = weights.reduce((sum, value) => sum + value, 0);
      return weightTotal === 0
        ? natural
        : natural.map((value, index) => value + (spare * weights[index]!) / weightTotal);
    }

    // Too wide. A column carrying prose is the one that can lose width without losing
    // anything - it simply takes another line - so the squeeze falls on the weighted
    // columns first, in proportion to what each has to spare. A column of subject names
    // keeps its width: "CHEMISTRY" broken over two lines is worse than a taller reason.
    const widths = [...natural];
    const prose = columns.map((column) => (column.weight ?? 0) > 0);
    for (const eligible of [prose, columns.map(() => true)]) {
      const excess = widths.reduce((sum, value) => sum + value, 0) - this.width;
      if (excess <= 0) break;
      const slack = widths.map((value, index) =>
        eligible[index] ? Math.max(0, value - MIN_COLUMN_WIDTH) : 0,
      );
      const slackTotal = slack.reduce((sum, value) => sum + value, 0);
      if (slackTotal === 0) continue;
      const taken = Math.min(excess, slackTotal);
      slack.forEach((value, index) => {
        widths[index] = widths[index]! - (taken * value) / slackTotal;
      });
    }

    // Every column already at its minimum and the table still too wide: it cannot be laid
    // out honestly, so everything is scaled to fit and every cell wraps.
    const over = widths.reduce((sum, value) => sum + value, 0);
    return over > this.width ? widths.map((value) => (value * this.width) / over) : widths;
  }

  /** Starts a new page unless `height` still fits on this one. */
  private ensure(height: number): void {
    if (this.y + height <= this.pdf.height - BOTTOM) return;
    this.pdf.addPage();
    this.y = MARGIN;
  }

  private footers(): void {
    const total = this.pdf.pageCount;
    const top = this.pdf.height - BOTTOM + 22;
    for (let index = 0; index < total; index++) {
      this.pdf.onPage(index, () => {
        this.pdf.rule(this.left, top - 8, this.width, 0.4, GRAY.rule);
        this.pdf.text(this.left, top, this.document.title, 'regular', SIZE.footer, GRAY.muted);
        const label = `Page ${index + 1} of ${total}`;
        const x = this.left + this.width - textWidth(label, 'regular', SIZE.footer);
        this.pdf.text(x, top, label, 'regular', SIZE.footer, GRAY.muted);
      });
    }
  }
}
