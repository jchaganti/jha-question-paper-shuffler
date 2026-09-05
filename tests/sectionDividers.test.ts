/**
 * A subject is not always the unit a question may move in.
 *
 * Papers divide a subject into "SECTION A (All questions are compulsory)" and "SECTION B
 * (Attempt any 10 questions)", and sometimes into parts above those. Those divisions carry
 * rules of their own: a compulsory question moved into the attempt-any-10 section changes
 * what the candidate is asked to do. So the divisions bound the shuffle, and their heading
 * paragraphs stay exactly where they are.
 */
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { DocxPackage } from '../src/core/docx/DocxPackage';
import { NS, visibleText } from '../src/core/docx/xml';
import { GenerationService } from '../src/core/generate/GenerationService';
import { NumberingIndex } from '../src/core/parse/NumberingIndex';
import { PaperParser } from '../src/core/parse/PaperParser';
import { buildPaper, type FixtureQuestion, type FixtureSection } from './support/PaperFixture';

const question = (n: number, headingBefore?: string): FixtureQuestion => ({
  stem: `Question ${n}`,
  optionParagraphs: [`(A)\tone ${n}\t(B)\ttwo ${n}\t(C)\tthree ${n}\t(D)\tfour ${n}`],
  answer: 'A',
  ...(headingBefore === undefined ? {} : { headingBefore }),
});

/** `count` questions in one subject, with `heading` announced before question `at`. */
function divided(subject: string, heading: string, at: number, count = 6): FixtureSection[] {
  return [
    {
      subject,
      startNumber: 1,
      numId: '1',
      questions: Array.from({ length: count }, (_unused, i) =>
        question(i + 1, i + 1 === at ? heading : undefined),
      ),
    },
  ];
}

/** One subject whose questions carry the given headings, by question number. */
function withHeadings(subject: string, headings: Record<number, string>, count: number): FixtureSection[] {
  return [
    {
      subject,
      startNumber: 1,
      numId: '1',
      questions: Array.from({ length: count }, (_unused, i) => question(i + 1, headings[i + 1])),
    },
  ];
}

async function parse(sections: readonly FixtureSection[]) {
  const pkg = await DocxPackage.fromBuffer(await buildPaper(sections));
  const numbering = new NumberingIndex(pkg.numberingPart());
  return new PaperParser().parsePart(pkg.documentPart(), numbering);
}

/** The label and question numbers of every run of questions the paper splits into. */
const layout = (paper: Awaited<ReturnType<typeof parse>>): [string, number[]][] =>
  paper.sections.map((section) => [section.label, section.blocks.map((b) => b.printedNumber)]);

describe('a subject divided into sections', () => {
  it('becomes one run of questions per section', async () => {
    const paper = await parse(divided('PHYSICS', 'SECTION B (Attempt any 10 questions)', 4));

    expect(layout(paper)).toEqual([
      ['PHYSICS', [1, 2, 3]],
      ['PHYSICS - SECTION B', [4, 5, 6]],
    ]);
  });

  it('names the first run too when it has its own heading', async () => {
    const paper = await parse(
      withHeadings(
        'PHYSICS',
        { 1: 'SECTION A (All questions are compulsory)', 4: 'SECTION B (Attempt any 10 questions)' },
        6,
      ),
    );

    expect(layout(paper)).toEqual([
      ['PHYSICS - SECTION A', [1, 2, 3]],
      ['PHYSICS - SECTION B', [4, 5, 6]],
    ]);
  });

  it('reads a dash, a repeated subject name and a roman numeral the same way', async () => {
    const labelOf = async (heading: string): Promise<string> =>
      layout(await parse(divided('PHYSICS', heading, 4)))[1]![0];

    expect(await labelOf('SECTION - B')).toBe('PHYSICS - SECTION B');
    expect(await labelOf('PHYSICS PART 2')).toBe('PHYSICS - PART 2');
    expect(await labelOf('PART II')).toBe('PHYSICS - PART II');
  });

  it('nests sections under parts, and a new part restarts them', async () => {
    const paper = await parse(
      withHeadings(
        'BIOLOGY',
        {
          1: 'BIOLOGY PART 1',
          3: 'SECTION - B (Attempt any 10 questions)',
          4: 'BIOLOGY PART 2',
          5: 'SECTION - B (Attempt any 10 questions)',
        },
        6,
      ),
    );

    // PART 2 clears PART 1's section: question 4 is in "PART 2", not "PART 1 SECTION B".
    expect(layout(paper)).toEqual([
      ['BIOLOGY - PART 1', [1, 2]],
      ['BIOLOGY - PART 1 SECTION B', [3]],
      ['BIOLOGY - PART 2', [4]],
      ['BIOLOGY - PART 2 SECTION B', [5, 6]],
    ]);
  });

  it('does not open a run for a heading with no questions under it', async () => {
    // "BIOLOGY PART 1" sits immediately above "SECTION - A", as it does in a real paper.
    // Two headings, one run of questions, named by both.
    const paper = await parse(
      withHeadings(
        'BIOLOGY',
        { 1: 'BIOLOGY PART 1\nSECTION - A (All questions are Compulsory)' },
        6,
      ),
    );

    expect(layout(paper)).toEqual([['BIOLOGY - PART 1 SECTION A', [1, 2, 3, 4, 5, 6]]]);
  });

  it('leaves a subject with no such headings as one run', async () => {
    const paper = await parse(divided('PHYSICS', 'Read the following passage', 4));

    expect(layout(paper)).toEqual([['PHYSICS', [1, 2, 3, 4, 5, 6]]]);
  });

  it('drops the stand-in subject name from the label when no subject was recognised', async () => {
    const paper = await parse(divided('MODEL TEST PAPER', 'SECTION B (Attempt any 10 questions)', 4));

    expect(layout(paper).map(([label]) => label)).toEqual(['ALL', 'SECTION B']);
  });
});

describe('generating from a paper divided into sections', () => {
  const paper = (): FixtureSection[] =>
    divided('PHYSICS', 'SECTION B (Attempt any 10 questions)', 5, 8);

  it('never moves a question across the divide, and leaves the heading in place', async () => {
    const workingDir = await fs.mkdtemp(path.join(os.tmpdir(), 'shuffler-sections-'));
    try {
      const sourceFile = path.join(workingDir, 'Paper.docx');
      await fs.writeFile(sourceFile, await buildPaper(paper()));
      const result = await new GenerationService().generate({
        sourceFile,
        shuffleQuestions: true,
        shuffleOptions: true,
        questionExclusions: [],
        optionExclusions: [],
        setCount: 3,
        seed: 'sections',
      });

      for (const set of result.sets) {
        expect(set.verification.ok).toBe(true);
        // 1-4 are the first section, 5-8 the second. A mapping that crosses that line is
        // the defect this whole file exists to prevent.
        const crossed = set.mappings.filter((m) => m.newNumber <= 4 !== m.originalNumber <= 4);
        expect(crossed).toEqual([]);
      }

      const written = await fs.readFile(path.join(result.outputFolder, result.sets[0]!.fileName));
      const pkg = await DocxPackage.fromBuffer(written);
      const numbering = new NumberingIndex(pkg.numberingPart());
      const generated = new PaperParser().parsePart(pkg.documentPart(), numbering);

      expect(layout(generated)).toEqual([
        ['PHYSICS', [1, 2, 3, 4]],
        ['PHYSICS - SECTION B', [5, 6, 7, 8]],
      ]);
      // The heading is still one paragraph, still between the two runs of questions.
      const headings = generated.sections[1]!.headerNodes
        .filter((node) => node.namespaceURI === NS.w && node.localName === 'p')
        .map((node) => visibleText(node).trim())
        .filter((text) => text !== '');
      expect(headings).toEqual(['SECTION B (Attempt any 10 questions)']);
    } finally {
      await fs.rm(workingDir, { recursive: true, force: true });
    }
  });

  it('reports each section separately in the dry run', async () => {
    const workingDir = await fs.mkdtemp(path.join(os.tmpdir(), 'shuffler-sections-'));
    try {
      const sourceFile = path.join(workingDir, 'Paper.docx');
      await fs.writeFile(sourceFile, await buildPaper(paper()));
      const report = await new GenerationService().dryRun({
        sourceFile,
        shuffleQuestions: true,
        shuffleOptions: true,
        questionExclusions: [],
        optionExclusions: [],
        setCount: 2,
      });

      expect(report.groups).toEqual([
        { group: 'PHYSICS', questionCount: 4, movable: 4, optionsShuffled: 4 },
        { group: 'PHYSICS - SECTION B', questionCount: 4, movable: 4, optionsShuffled: 4 },
      ]);
    } finally {
      await fs.rm(workingDir, { recursive: true, force: true });
    }
  });
});
