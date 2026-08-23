import { describe, expect, it } from 'vitest';
import { DocxPackage } from '../src/core/docx/DocxPackage';
import { NumberingIndex } from '../src/core/parse/NumberingIndex';
import { PaperParseError, PaperParser } from '../src/core/parse/PaperParser';
import { XmlPart } from '../src/core/docx/xml';
import { buildPaper, defaultSections } from './support/PaperFixture';

async function parse(buffer: Buffer) {
  const pkg = await DocxPackage.fromBuffer(buffer);
  return new PaperParser().parsePart(pkg.documentPart(), new NumberingIndex(pkg.numberingPart()));
}

describe('PaperParser', () => {
  it('splits the paper into subjects and numbers questions from the answer key', async () => {
    const paper = await parse(await buildPaper(defaultSections()));

    expect(paper.questionCount).toBe(7);
    expect(paper.sections.map((section) => section.subject)).toEqual(['PHYSICS', 'CHEMISTRY']);
    expect(paper.sections[0]!.blocks.map((block) => block.printedNumber)).toEqual([1, 2, 3, 4]);
    expect(paper.sections[1]!.blocks.map((block) => block.printedNumber)).toEqual([5, 6, 7]);
  });

  it('recognises both question numbering lists and ignores option lettering lists', async () => {
    const paper = await parse(await buildPaper(defaultSections()));
    expect(paper.questionNumIds).toEqual(['1', '2']);
  });

  it('keeps a question stem together with its options, tables and images', async () => {
    const paper = await parse(await buildPaper(defaultSections()));
    const secondQuestion = paper.sections[0]!.blocks[1]!;

    expect(secondQuestion.nodes.length).toBe(3); // stem + two option paragraphs
    expect(secondQuestion.nodes[0]).toBe(secondQuestion.questionParagraph);
  });

  it('leaves subject spacer paragraphs behind so page breaks stay put', async () => {
    const paper = await parse(await buildPaper(defaultSections()));
    expect(paper.sections[0]!.tailNodes.length).toBe(3);
    expect(paper.sections[0]!.blocks[3]!.nodes.length).toBe(5); // stem + 4 auto-lettered options
  });

  it('reads the answer key', async () => {
    const paper = await parse(await buildPaper(defaultSections()));
    expect(paper.answerKey.questionNumbers).toEqual([1, 2, 3, 4, 5, 6, 7]);
    expect(paper.answerKey.answerOf(3)).toBe('C');
    expect(paper.answerKey.answerOf(5)).toBe('A');
  });

  it('rejects a paper whose answer key does not cover every question', async () => {
    const sections = defaultSections();
    const buffer = await buildPaper(sections);
    const pkg = await DocxPackage.fromBuffer(buffer);

    // Drop the last answer-key row to simulate a mismatch.
    const xml = pkg.getPartText('word/document.xml')!;
    const broken = xml.replace(/<w:tr>(?:(?!<w:tr>)[\s\S])*?<\/w:tr>(?=<\/w:tbl>)/, '');
    pkg.setPartText('word/document.xml', broken);

    expect(() =>
      new PaperParser().parsePart(XmlPart.parse(broken), new NumberingIndex(pkg.numberingPart())),
    ).toThrow(PaperParseError);
  });

  it('reports a helpful error when there is no answer key at all', async () => {
    const sections = defaultSections();
    const pkg = await DocxPackage.fromBuffer(await buildPaper(sections));
    const xml = pkg.getPartText('word/document.xml')!.replace(/<w:tbl>[\s\S]*<\/w:tbl>/, '');

    expect(() => new PaperParser().parsePart(XmlPart.parse(xml), new NumberingIndex(pkg.numberingPart()))).toThrow(
      /No answer key found/,
    );
  });
});
