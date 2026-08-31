/**
 * The report is a PDF, so the tool has to typeset one itself. These tests pin the two
 * things that go wrong when a PDF is written by hand: a file no reader will open, and text
 * that does not print the way it was measured.
 */
import { describe, expect, it } from 'vitest';
import { PdfBuilder, textWidth, wrap } from '../src/core/report/Pdf';
import { renderReportPdf } from '../src/core/report/ReportPdf';
import type { ReportDocument } from '../src/core/report/ReportDocument';
import { pdfLines, pdfPageCount, pdfText } from './support/pdfText';

const document = (blocks: ReportDocument['blocks']): ReportDocument => ({
  title: 'Question set generation report',
  subtitle: 'Sample Paper.docx   ·   generated 31-08-2026 at 13:21',
  blocks,
});

describe('the PDF a reader has to open', () => {
  it('has a header, a cross-reference table and a trailer', () => {
    const pdf = renderReportPdf(document([{ kind: 'paragraph', text: 'hello' }]));
    const raw = pdf.toString('latin1');

    expect(raw.startsWith('%PDF-1.4\n')).toBe(true);
    expect(raw.trimEnd().endsWith('%%EOF')).toBe(true);
    expect(raw).toContain('/Type /Catalog');
    expect(raw).toContain('/Type /Pages');
    expect(raw).toContain('xref\n');
    expect(raw).toContain('trailer\n');
  });

  it('points its cross-reference table at where each object really starts', () => {
    // A wrong offset is the classic hand-written-PDF fault: some readers recover from it
    // silently and others refuse the file, so it is checked rather than eyeballed.
    const raw = renderReportPdf(document([{ kind: 'paragraph', text: 'hello' }])).toString('latin1');
    const offsets = [...raw.matchAll(/^(\d{10}) 00000 n $/gm)].map((match) => Number(match[1]));

    expect(offsets.length).toBeGreaterThan(3);
    offsets.forEach((offset, index) => {
      expect(raw.slice(offset, offset + 20)).toMatch(new RegExp(`^${index + 1} 0 obj\\n`));
    });
    expect(Number(/startxref\n(\d+)/.exec(raw)![1])).toBe(raw.indexOf('xref\n0 '));
  });

  it('embeds no font, so every reader has the three it uses', () => {
    const raw = renderReportPdf(document([{ kind: 'paragraph', text: 'hello' }])).toString('latin1');
    expect(raw).toContain('/BaseFont /Helvetica ');
    expect(raw).toContain('/BaseFont /Helvetica-Bold ');
    expect(raw).toContain('/BaseFont /Courier ');
    expect(raw).not.toContain('/FontFile');
  });
});

describe('text that cannot be written in the font', () => {
  it('spells the option-mapping arrow, rather than losing it', () => {
    const pdf = renderReportPdf(
      document([
        {
          kind: 'table',
          columns: [{ header: 'Option mapping' }],
          rows: [['A→C, B→A, C→D, D→B']],
        },
      ]),
    );

    expect(pdfText(pdf)).toContain('A->C, B->A, C->D, D->B');
    expect(pdfText(pdf)).not.toContain('?');
  });

  it('measures what it will print, so a spelled-out character stays inside its column', () => {
    // "→" is drawn as two characters, so a width taken from the original string would
    // reserve too little room and the text would run into the next column.
    expect(textWidth('A→C', 'regular', 10)).toBe(textWidth('A->C', 'regular', 10));
  });

  it('shows a character it can neither encode nor spell as "?"', () => {
    expect(pdfText(renderReportPdf(document([{ kind: 'paragraph', text: 'Ω is 60' }])))).toContain(
      '? is 60',
    );
  });

  it('keeps the punctuation Word produces', () => {
    const text = 'the paper\u2019s "none of these" \u2014 kept';
    expect(pdfText(renderReportPdf(document([{ kind: 'paragraph', text }])))).toContain(text);
  });
});

describe('laying out the report', () => {
  it('wraps a paragraph inside the page rather than off the edge', () => {
    const long = 'word '.repeat(400).trim();
    const lines = pdfLines(renderReportPdf(document([{ kind: 'paragraph', text: long }])));
    const body = lines.filter((line) => line.startsWith('word'));

    expect(body.length).toBeGreaterThan(10);
    for (const line of body) expect(textWidth(line, 'regular', 9)).toBeLessThanOrEqual(503.28);
  });

  it('does not wrap a cell that fits, however exactly it fits', () => {
    // The column is sized from the widest cell, so the widest cell measures as exactly the
    // width available to it - one floating-point step is the difference between one line
    // and "Subjec" over "t".
    const pdf = renderReportPdf(
      document([
        {
          kind: 'table',
          columns: [{ header: 'Q' }, { header: 'Subject' }],
          rows: [
            ['12', 'PHYSICS'],
            ['138', 'CHEMISTRY'],
          ],
        },
      ]),
    );

    const lines = pdfText(pdf);
    expect(lines).toContain('12 PHYSICS');
    expect(lines).toContain('138 CHEMISTRY');
    expect(lines).toContain('Q Subject');
  });

  it('squeezes the prose column, not the one holding a subject name', () => {
    // A table wider than the page has to lose width somewhere. "CHEMISTRY" broken over two
    // lines is worse than a taller reason column, so only the weighted column gives way.
    const reason = 'option-contains-floating-graphic: '.repeat(6);
    const pdf = renderReportPdf(
      document([
        {
          kind: 'table',
          columns: [{ header: 'Q' }, { header: 'Subject' }, { header: 'Reason', weight: 1 }],
          rows: [['82', 'CHEMISTRY', reason]],
        },
      ]),
    );

    const lines = pdfLines(pdf);
    expect(lines.some((line) => line.startsWith('82 CHEMISTRY '))).toBe(true);
    expect(lines).not.toContain('RY');
    // The reason took the squeeze instead, so it runs over several lines.
    expect(lines.filter((line) => line.includes('option-contains-floating-graphic')).length)
      .toBeGreaterThan(1);
  });

  it('repeats the table header on every page a long table runs over', () => {
    const rows = Array.from({ length: 220 }, (_unused, index) => [String(index + 1), 'ALL', 'A->B']);
    const pdf = renderReportPdf(
      document([
        {
          kind: 'table',
          columns: [{ header: 'New Q' }, { header: 'Subject' }, { header: 'Option mapping' }],
          rows,
        },
      ]),
    );

    expect(pdfPageCount(pdf)).toBeGreaterThan(1);
    const headers = pdfLines(pdf).filter((line) => line === 'New Q Subject Option mapping');
    expect(headers.length).toBe(pdfPageCount(pdf));
    // Every row survived the page breaks.
    expect(pdfLines(pdf).filter((line) => /^\d+ ALL A->B$/.test(line))).toHaveLength(220);
  });

  it('numbers every page', () => {
    const rows = Array.from({ length: 220 }, (_unused, index) => [String(index + 1), 'ALL']);
    const pdf = renderReportPdf(
      document([{ kind: 'table', columns: [{ header: 'Q' }, { header: 'Subject' }], rows }]),
    );

    const total = pdfPageCount(pdf);
    const text = pdfText(pdf);
    for (let page = 1; page <= total; page++) expect(text).toContain(`Page ${page} of ${total}`);
  });

  it('is the same file for the same content', () => {
    const blocks: ReportDocument['blocks'] = [
      { kind: 'heading', text: 'This run' },
      { kind: 'facts', items: [{ label: 'Seed', value: 'march-batch', mono: true }] },
    ];
    expect(renderReportPdf(document(blocks)).equals(renderReportPdf(document(blocks)))).toBe(true);
  });
});

describe('wrap', () => {
  it('breaks at spaces', () => {
    expect(wrap('one two three', 'regular', 10, textWidth('one two', 'regular', 10))).toEqual([
      'one two',
      'three',
    ]);
  });

  it('breaks a single word too wide for the column, so it cannot overflow', () => {
    const lines = wrap('C:/papers/a-very-long-file-name.docx', 'mono', 9, 40);
    expect(lines.length).toBeGreaterThan(1);
    for (const line of lines) expect(textWidth(line, 'mono', 9)).toBeLessThanOrEqual(40);
    expect(lines.join('')).toBe('C:/papers/a-very-long-file-name.docx');
  });
});

describe('PdfBuilder', () => {
  it('counts the pages it was asked for', () => {
    const pdf = new PdfBuilder();
    expect(pdf.pageCount).toBe(1);
    pdf.addPage();
    pdf.addPage();
    expect(pdf.pageCount).toBe(3);
    expect(pdfPageCount(pdf.toBuffer())).toBe(3);
  });

  it('draws on an earlier page without disturbing the current one', () => {
    const pdf = new PdfBuilder();
    pdf.text(40, 40, 'first page', 'regular', 10);
    pdf.addPage();
    pdf.text(40, 40, 'second page', 'regular', 10);
    pdf.onPage(0, () => pdf.text(40, 800, 'footer of page one', 'regular', 8));
    pdf.text(40, 60, 'still the second page', 'regular', 10);

    const lines = pdfLines(pdf.toBuffer());
    expect(lines).toEqual(['first page', 'footer of page one', 'second page', 'still the second page']);
  });
});
