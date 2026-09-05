/**
 * What the user is told about a question that could not be read.
 *
 * A skipped question is only useful if the person who typed the paper can act on it, so
 * every refusal owes two things: a `detail` saying what is wrong with *this* question,
 * naming the option, and a `fix` saying what to change in Word. Neither may fall back on
 * the `SkipReason` slug, which is a name for the code and means nothing to a reader.
 *
 * The refusals themselves are covered where the behaviour they guard lives; this file
 * pins the shape and the wording that reach the dry run, the report and the UI.
 */
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { DocxPackage } from '../src/core/docx/DocxPackage';
import { GenerationService } from '../src/core/generate/GenerationService';
import { ReportWriter } from '../src/core/generate/ReportWriter';
import { OptionSetParser } from '../src/core/options/OptionSetParser';
import { NumberingIndex } from '../src/core/parse/NumberingIndex';
import { PaperParser } from '../src/core/parse/PaperParser';
import { SKIP_REASON_LABEL, groupSkippedOptions } from '../src/shared/skipReasons';
import type { GenerationRequest, SkipReason } from '../src/shared/types';
import { pdfText } from './support/pdfText';
import { buildPaper, type FixtureQuestion, type FixtureSection } from './support/PaperFixture';

function paperWith(questions: readonly FixtureQuestion[]): FixtureSection[] {
  const filler = Array.from({ length: 6 }, (_unused, index) => ({
    stem: `Filler question ${index + 1}`,
    optionParagraphs: ['(A)\tone\t(B)\ttwo\t(C)\tthree\t(D)\tfour'],
    answer: 'A' as const,
  }));
  return [{ subject: 'PHYSICS', startNumber: 1, numId: '1', questions: [...questions, ...filler] }];
}

/** A question whose options are inside a table - a refusal shared by several questions. */
const inTable = (n: number): FixtureQuestion => ({
  stem: `Question ${n} with its options in a table`,
  optionParagraphs: [],
  rawOptionParagraph:
    '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>(A) one</w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:p><w:r><w:t>(B) two</w:t></w:r></w:p></w:tc></w:tr>' +
    '<w:tr><w:tc><w:p><w:r><w:t>(C) three</w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:p><w:r><w:t>(D) four</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
  answer: 'A',
});

/** An option split by a paragraph break - a different refusal, one question. */
const spilled: FixtureQuestion = {
  stem: 'Which statements are correct?',
  optionParagraphs: ['(A)\tone', '(B)\ttwo and it', 'runs on', '(C)\tthree', '(D)\tfour'],
  answer: 'A',
};

async function refusalFor(question: FixtureQuestion) {
  const pkg = await DocxPackage.fromBuffer(await buildPaper(paperWith([question])));
  const numbering = new NumberingIndex(pkg.numberingPart());
  const paper = new PaperParser().parsePart(pkg.documentPart(), numbering);
  const parsed = new OptionSetParser(numbering).parse(paper.sections[0]!.blocks[0]!);
  if (parsed.ok) throw new Error('expected this question to be refused');
  return parsed;
}

describe('every refusal', () => {
  it('has a plain-language heading for its reason', () => {
    // A slug reaching the page is the failure this guards: every reason needs a sentence.
    for (const [reason, label] of Object.entries(SKIP_REASON_LABEL)) {
      expect(label, reason).not.toBe(reason);
      expect(label, reason).toContain(' ');
      expect(label.length, reason).toBeGreaterThan(15);
    }
  });

  it('says what is wrong and what to change, in separate sentences', async () => {
    for (const question of [inTable(1), spilled]) {
      const refusal = await refusalFor(question);
      expect(refusal.detail.length).toBeGreaterThan(30);
      expect(refusal.fix.length).toBeGreaterThan(30);
      // The fix is an instruction, not a restatement of the problem.
      expect(refusal.fix).not.toBe(refusal.detail);
    }
  });

  it('names a key or a menu the author can actually find', async () => {
    expect((await refusalFor(spilled)).fix).toContain('Shift+Enter');
    expect((await refusalFor(inTable(1))).fix).toContain('Convert to Text');
  });
});

describe('gathering skipped questions by problem', () => {
  const item = (questionNumber: number, reason: SkipReason) => ({
    questionNumber,
    subject: 'PHYSICS',
    reason,
    detail: `Q${questionNumber} detail`,
    fix: `fix for ${reason}`,
  });

  it('states each fix once, with the questions it covers', () => {
    const groups = groupSkippedOptions([
      item(3, 'options-inside-table'),
      item(9, 'option-spans-paragraphs'),
      item(1, 'options-inside-table'),
    ]);

    expect(groups).toHaveLength(2);
    // Most-affected first: that is the one worth fixing before the next paper.
    expect(groups[0]!.label).toBe(SKIP_REASON_LABEL['options-inside-table']);
    expect(groups[0]!.questionNumbers).toEqual([1, 3]);
    expect(groups[0]!.fix).toBe('fix for options-inside-table');
    expect(groups[1]!.questionNumbers).toEqual([9]);
  });

  it('keeps the detail of each question under the group', () => {
    const groups = groupSkippedOptions([item(3, 'options-inside-table'), item(1, 'options-inside-table')]);

    expect(groups[0]!.questions.map((q) => q.detail)).toEqual(['Q1 detail', 'Q3 detail']);
  });
});

describe('what reaches the run', () => {
  let workingDir = '';
  let sourceFile = '';
  const service = new GenerationService();
  const request = (): GenerationRequest => ({
    sourceFile,
    shuffleQuestions: false,
    shuffleOptions: true,
    questionExclusions: [],
    optionExclusions: [],
    setCount: 2,
    seed: 'skip-explanations',
  });

  beforeEach(async () => {
    workingDir = await fs.mkdtemp(path.join(os.tmpdir(), 'shuffler-skips-'));
    sourceFile = path.join(workingDir, 'Paper.docx');
    await fs.writeFile(sourceFile, await buildPaper(paperWith([inTable(1), inTable(2), spilled])));
  });

  afterEach(async () => {
    await fs.rm(workingDir, { recursive: true, force: true });
  });

  it('groups the dry run by problem, worst first', async () => {
    const report = await service.dryRun(request());

    expect(report.paper.unshufflableGroups.map((group) => group.label)).toEqual([
      SKIP_REASON_LABEL['options-inside-table'],
      SKIP_REASON_LABEL['option-spans-paragraphs'],
    ]);
    expect(report.paper.unshufflableGroups[0]!.questionNumbers).toEqual([1, 2]);
  });

  it('writes the problem, the questions and the fix into the report file', async () => {
    const result = await service.generate(request());

    expect(pdfText(await fs.readFile(result.reportFile))).toContain('Options that could not be shuffled');
    const cells = new ReportWriter()
      .build(request(), result)
      .blocks.filter((block) => block.kind === 'table')
      .flatMap((block) => (block as { rows: readonly (readonly string[])[] }).rows)
      .flat();
    // Every question in the table group says the same thing, so the heading carries it.
    expect(cells.some((cell) => cell.startsWith(SKIP_REASON_LABEL['options-inside-table']))).toBe(true);
    expect(cells.some((cell) => cell.includes('a table cell is not a line of text'))).toBe(true);
    expect(cells.some((cell) => cell.includes('Convert to Text'))).toBe(true);
    expect(cells).toContain('1, 2');
    // The split option is a group of one, so its own sentence is there as well.
    expect(cells.some((cell) => cell.includes('Option (B) runs on to a paragraph of its own'))).toBe(true);
  });

  it('never shows the reason slug to the user', async () => {
    const result = await service.generate(request());
    const text = pdfText(await fs.readFile(result.reportFile));

    for (const reason of Object.keys(SKIP_REASON_LABEL)) {
      expect(text, reason).not.toContain(reason);
    }
  });
});
