import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { DocxPackage } from '../src/core/docx/DocxPackage';
import type { Element } from '../src/core/docx/dom';
import { NS, childElements, descendants, firstChild } from '../src/core/docx/xml';
import { GenerationService } from '../src/core/generate/GenerationService';
import { PageFlowGuard } from '../src/core/generate/PageFlow';
import { NumberingIndex } from '../src/core/parse/NumberingIndex';
import { PaperParser, isEmptyParagraph } from '../src/core/parse/PaperParser';
import type { ParsedPaper } from '../src/core/parse/PaperModel';
import type { GenerationRequest } from '../src/shared/types';
import { buildPaper, defaultSections } from './support/PaperFixture';

let workingDir = '';
let sourceFile = '';

const service = new GenerationService();

const request = (overrides: Partial<GenerationRequest> = {}): GenerationRequest => ({
  sourceFile,
  shuffleQuestions: true,
  shuffleOptions: true,
  questionExclusions: [],
  optionExclusions: [],
  setCount: 1,
  seed: 'page-flow',
  ...overrides,
});

async function parse(buffer: Buffer): Promise<ParsedPaper> {
  const pkg = await DocxPackage.fromBuffer(buffer);
  return new PaperParser().parsePart(pkg.documentPart(), new NumberingIndex(pkg.numberingPart()));
}

const has = (paragraph: Element, flag: string): boolean => {
  const pPr = firstChild(paragraph, NS.w, 'pPr');
  return !!pPr && !!firstChild(pPr, NS.w, flag);
};

/** Paragraphs of a question, in reading order, tables included. */
function paragraphsOf(nodes: readonly Element[]): Element[] {
  return nodes.flatMap((node) =>
    node.namespaceURI === NS.w && node.localName === 'p' ? [node] : descendants(node, NS.w, 'p'),
  );
}

beforeEach(async () => {
  workingDir = await fs.mkdtemp(path.join(os.tmpdir(), 'shuffler-pageflow-'));
  sourceFile = path.join(workingDir, 'Sample Paper.docx');
  await fs.writeFile(sourceFile, await buildPaper(defaultSections()));
});

afterEach(async () => {
  await fs.rm(workingDir, { recursive: true, force: true });
});

describe('PageFlowGuard', () => {
  it('chains every paragraph of a question to the next, except the last', async () => {
    const paper = await parse(await fs.readFile(sourceFile));
    // Question 3 has one option per line: stem + four options.
    const block = paper.sections[0]!.blocks[2]!;
    new PageFlowGuard().keepBlockWhole(block);

    const paragraphs = paragraphsOf(block.nodes);
    expect(paragraphs).toHaveLength(5);
    expect(paragraphs.map((p) => has(p, 'keepNext'))).toEqual([true, true, true, true, false]);
    // Every paragraph also refuses to split its own lines over a page break.
    expect(paragraphs.every((p) => has(p, 'keepLines'))).toBe(true);
  });

  it('marks a one-paragraph question with keepLines but no keepNext', async () => {
    const paper = await parse(await fs.readFile(sourceFile));
    // Question 1 keeps its stem and all four options on two paragraphs.
    const block = paper.sections[0]!.blocks[0]!;
    new PageFlowGuard().keepBlockWhole(block);

    const paragraphs = paragraphsOf(block.nodes);
    expect(paragraphs).toHaveLength(2);
    expect(has(paragraphs[0]!, 'keepNext')).toBe(true);
    expect(has(paragraphs[1]!, 'keepNext')).toBe(false);
    expect(has(paragraphs[1]!, 'keepLines')).toBe(true);
  });

  it('leaves trailing blank paragraphs out of the chain, so the break can fall in the gap', async () => {
    const paper = await parse(await fs.readFile(sourceFile));
    // The last block of a subject keeps its spacer paragraphs in section.tailNodes, so
    // build a block that ends in blanks by hand: the shape of a question followed by
    // spacing inside a subject.
    const section = paper.sections[0]!;
    const block = section.blocks[2]!;
    const doc = block.nodes[0]!.ownerDocument!;
    const blank = doc.createElementNS(NS.w, 'w:p');
    block.nodes.push(blank);

    new PageFlowGuard().keepBlockWhole(block);

    expect(isEmptyParagraph(blank)).toBe(true);
    expect(has(blank, 'keepNext')).toBe(false);
    expect(has(blank, 'keepLines')).toBe(false);
    // The last real paragraph is still the end of the chain.
    const lastReal = paragraphsOf(block.nodes).filter((p) => !isEmptyParagraph(p)).pop()!;
    expect(has(lastReal, 'keepNext')).toBe(false);
    expect(has(lastReal, 'keepLines')).toBe(true);
  });

  it('inserts the flags where the OOXML schema demands, before numPr', async () => {
    const paper = await parse(await fs.readFile(sourceFile));
    const block = paper.sections[0]!.blocks[2]!;
    new PageFlowGuard().keepBlockWhole(block);

    // The stem paragraph already carries pStyle + numPr; order must stay pStyle,
    // keepNext, keepLines, numPr or Word refuses to open the document.
    const pPr = firstChild(block.questionParagraph, NS.w, 'pPr')!;
    expect(childElements(pPr).map((child) => child.localName)).toEqual([
      'pStyle',
      'keepNext',
      'keepLines',
      'numPr',
    ]);
  });

  it('is idempotent: flags already present are left alone', async () => {
    const paper = await parse(await fs.readFile(sourceFile));
    const block = paper.sections[0]!.blocks[2]!;
    const guard = new PageFlowGuard();
    guard.keepBlockWhole(block);
    guard.keepBlockWhole(block);

    const pPr = firstChild(block.questionParagraph, NS.w, 'pPr')!;
    expect(childElements(pPr).filter((child) => child.localName === 'keepNext')).toHaveLength(1);
    expect(childElements(pPr).filter((child) => child.localName === 'keepLines')).toHaveLength(1);
  });

  it('marks every question of the paper and reports how many', async () => {
    const paper = await parse(await fs.readFile(sourceFile));
    expect(new PageFlowGuard().keepQuestionsWhole(paper)).toBe(7);
    for (const section of paper.sections) {
      for (const block of section.blocks) {
        expect(has(block.questionParagraph, 'keepNext')).toBe(true);
      }
    }
  });

  it('stops a table inside a question from splitting across a page', async () => {
    const paper = await parse(await fs.readFile(sourceFile));
    const block = paper.sections[0]!.blocks[1]!;

    // Insert a two-row table into the question, the way a matching-type question has one.
    const doc = block.nodes[0]!.ownerDocument!;
    const table = doc.createElementNS(NS.w, 'w:tbl');
    for (const text of ['Column I', 'Column II']) {
      const row = doc.createElementNS(NS.w, 'w:tr');
      const cell = doc.createElementNS(NS.w, 'w:tc');
      const paragraph = doc.createElementNS(NS.w, 'w:p');
      const run = doc.createElementNS(NS.w, 'w:r');
      const t = doc.createElementNS(NS.w, 'w:t');
      t.appendChild(doc.createTextNode(text));
      run.appendChild(t);
      paragraph.appendChild(run);
      cell.appendChild(paragraph);
      row.appendChild(cell);
      table.appendChild(row);
    }
    block.nodes.splice(1, 0, table);

    new PageFlowGuard().keepBlockWhole(block);

    const rows = descendants(table, NS.w, 'tr');
    expect(rows).toHaveLength(2);
    for (const row of rows) {
      const trPr = firstChild(row, NS.w, 'trPr');
      expect(trPr && firstChild(trPr, NS.w, 'cantSplit')).toBeTruthy();
      // trPr must precede the cells.
      expect(childElements(row)[0]!.localName).toBe('trPr');
    }
    // Cell paragraphs join the chain, so the table travels with the rest of the question.
    expect(has(descendants(table, NS.w, 'p')[0]!, 'keepNext')).toBe(true);
  });
});

describe('the answer key starts a page of its own', () => {
  const keyHeadingOf = (paper: ParsedPaper): Element => paper.answerKeyNodes[0]!;

  it('adds a page break before the ANSWER KEY heading', async () => {
    const paper = await parse(await fs.readFile(sourceFile));
    expect(new PageFlowGuard().answerKeyOnNewPage(paper)).toBe(true);
    expect(has(keyHeadingOf(paper), 'pageBreakBefore')).toBe(true);
  });

  it('puts the break inside the table when the paper has no ANSWER KEY heading', async () => {
    const buffer = await buildPaper(defaultSections(), { omitAnswerKeyHeading: true });
    const paper = await parse(buffer);

    const table = keyHeadingOf(paper);
    expect(table.localName).toBe('tbl');
    expect(new PageFlowGuard().answerKeyOnNewPage(paper)).toBe(true);
    // A table cannot carry the property, so it goes on the first paragraph inside it.
    expect(has(descendants(table, NS.w, 'p')[0]!, 'pageBreakBefore')).toBe(true);
  });

  it('adds nothing when a typed page break is already there', async () => {
    const buffer = await buildPaper(defaultSections(), {
      beforeAnswerKey: '<w:p><w:r><w:br w:type="page"/></w:r></w:p>',
    });
    const paper = await parse(buffer);

    expect(new PageFlowGuard().answerKeyOnNewPage(paper)).toBe(false);
    expect(has(keyHeadingOf(paper), 'pageBreakBefore')).toBe(false);
  });

  it('adds nothing when a section break is already there', async () => {
    const buffer = await buildPaper(defaultSections(), {
      beforeAnswerKey: '<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr></w:p>',
    });
    const paper = await parse(buffer);

    expect(new PageFlowGuard().answerKeyOnNewPage(paper)).toBe(false);
  });

  it('is idempotent', async () => {
    const paper = await parse(await fs.readFile(sourceFile));
    const guard = new PageFlowGuard();
    expect(guard.answerKeyOnNewPage(paper)).toBe(true);
    expect(guard.answerKeyOnNewPage(paper)).toBe(false);

    const pPr = firstChild(keyHeadingOf(paper), NS.w, 'pPr')!;
    expect(childElements(pPr).filter((child) => child.localName === 'pageBreakBefore')).toHaveLength(1);
  });

  it('inserts the property where the schema demands', async () => {
    const paper = await parse(await fs.readFile(sourceFile));
    // Give the heading the properties a real paper's centred heading carries.
    const heading = keyHeadingOf(paper);
    const doc = heading.ownerDocument!;
    const pPr = doc.createElementNS(NS.w, 'w:pPr');
    pPr.appendChild(doc.createElementNS(NS.w, 'w:jc'));
    pPr.appendChild(doc.createElementNS(NS.w, 'w:rPr'));
    heading.insertBefore(pPr, heading.firstChild);

    new PageFlowGuard().answerKeyOnNewPage(paper);
    expect(childElements(pPr).map((child) => child.localName)).toEqual(['pageBreakBefore', 'jc', 'rPr']);
  });
});

describe('generation with page flow', () => {
  it('keeps every question whole by default and says so in the result', async () => {
    const result = await service.generate(request());
    expect(result.sets[0]!.questionsKeptWhole).toBe(7);
    expect(result.sets[0]!.verification.ok).toBe(true);
    expect(has((await parse(await fs.readFile(result.sets[0]!.filePath))).answerKeyNodes[0]!, 'pageBreakBefore')).toBe(
      true,
    );

    const generated = await parse(await fs.readFile(result.sets[0]!.filePath));
    for (const section of generated.sections) {
      for (const block of section.blocks) {
        const paragraphs = paragraphsOf(block.nodes).filter((p) => !isEmptyParagraph(p));
        expect(paragraphs.every((p) => has(p, 'keepLines'))).toBe(true);
        expect(paragraphs.slice(0, -1).every((p) => has(p, 'keepNext'))).toBe(true);
        expect(has(paragraphs[paragraphs.length - 1]!, 'keepNext')).toBe(false);
      }
    }
  });

  it('leaves the page flow untouched when the option is switched off', async () => {
    const result = await service.generate(request({ keepQuestionsWhole: false }));
    expect(result.sets[0]!.questionsKeptWhole).toBe(0);

    const generated = await parse(await fs.readFile(result.sets[0]!.filePath));
    const marked = generated.sections
      .flatMap((section) => section.blocks)
      .flatMap((block) => paragraphsOf(block.nodes))
      .filter((paragraph) => has(paragraph, 'keepNext') || has(paragraph, 'keepLines'));
    expect(marked).toEqual([]);
  });

  it('still starts the answer key on a new page when that option is switched off', async () => {
    // The two are independent: the key is never printed under the last question.
    const result = await service.generate(request({ keepQuestionsWhole: false }));
    const generated = await parse(await fs.readFile(result.sets[0]!.filePath));
    expect(has(generated.answerKeyNodes[0]!, 'pageBreakBefore')).toBe(true);
  });

  it('does not disturb question order, options or the answer key', async () => {
    const withFlow = await service.generate(request({ seed: 'same' }));
    const withoutFlow = await service.generate(request({ seed: 'same', keepQuestionsWhole: false }));

    const read = async (file: string): Promise<string[]> => {
      const paper = await parse(await fs.readFile(file));
      return paper.sections.flatMap((section) =>
        section.blocks.map(
          (block) =>
            `${block.printedNumber}:${block.questionParagraph.textContent?.trim()}` +
            `=${paper.answerKey.answerOf(block.printedNumber)}`,
        ),
      );
    };

    expect(await read(withFlow.sets[0]!.filePath)).toEqual(await read(withoutFlow.sets[0]!.filePath));
  });

  it('records the setting in the generation report', async () => {
    const result = await service.generate(request());
    const report = await fs.readFile(result.reportFile, 'utf8');
    expect(report).toContain('Keep each question on one page: yes');
    expect(report).toContain('Questions marked to stay whole on one page: 7');
  });
});
