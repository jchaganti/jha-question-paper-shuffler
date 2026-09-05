/**
 * Reporting of option layouts that the parser can read only by falling back on a
 * heuristic. These questions are shuffled normally; the note exists so that whoever types
 * the paper can correct it. Every case here used to be reported as a *skip* before the
 * heuristics were added, and must stay visible now that it succeeds.
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
import { groupLayoutNotes, questionsWithLayoutNotes } from '../src/shared/layoutNotes';
import type { GenerationRequest, OptionLayoutIssue, QuestionLayoutNote } from '../src/shared/types';
import { pdfText } from './support/pdfText';
import { LIST, buildPaper, defaultSections, type FixtureQuestion, type FixtureSection } from './support/PaperFixture';

/** A paper whose first question is the one under test, padded so the key is detectable. */
function paperWith(question: FixtureQuestion): FixtureSection[] {
  const filler = Array.from({ length: 6 }, (_unused, index) => ({
    stem: `Filler question ${index + 1}`,
    optionParagraphs: ['(A)\tone\t(B)\ttwo\t(C)\tthree\t(D)\tfour'],
    answer: 'A' as const,
  }));
  return [{ subject: 'PHYSICS', startNumber: 1, numId: '1', questions: [question, ...filler] }];
}

/** The notes reported for the first question of a paper built from `question`. */
async function notesFor(question: FixtureQuestion): Promise<readonly { issue: string; detail: string }[]> {
  const pkg = await DocxPackage.fromBuffer(await buildPaper(paperWith(question)));
  const numbering = new NumberingIndex(pkg.numberingPart());
  const paper = new PaperParser().parsePart(pkg.documentPart(), numbering);
  const parsed = new OptionSetParser(numbering).parse(paper.sections[0]!.blocks[0]!);
  if (!parsed.ok) throw new Error(`expected the options to parse, got ${parsed.reason}: ${parsed.detail}`);
  return parsed.notes;
}

const issues = (notes: readonly { issue: string }[]): string[] => notes.map((note) => note.issue);

describe('option layouts that are read but reported', () => {
  it('reports option (A) lettered by Word while (B), (C), (D) are typed', async () => {
    const notes = await notesFor({
      stem: 'Which quantity is a vector?',
      optionParagraphs: ['speed', '(B)\tmass', '(C)\tvelocity', '(D)\ttime'],
      answer: 'C',
      letteredFirstOption: true,
    });

    expect(issues(notes)).toEqual(['mixed-auto-and-typed-labels']);
    expect(notes[0]!.detail).toContain('lettered by Word');
    expect(notes[0]!.detail).toContain('(B), (C), (D)');
  });

  it('reports a label with no tab in front of it, naming the label', async () => {
    const notes = await notesFor({
      stem: 'Which statements are incorrect?',
      optionParagraphs: ['(A)\t(i) and (ii)\t(B)\t(i) and (iii)\t(C)\t(i), (iv) and (v) (D)\t(iii) and (v)'],
      answer: 'A',
    });

    expect(issues(notes)).toEqual(['label-not-after-tab']);
    expect(notes[0]!.detail).toContain('(D)');
    expect(notes[0]!.detail).not.toContain('(C)');
    expect(notes[0]!.fix).toContain('Tab');
  });

  it('reports labels that mix upper and lower case', async () => {
    const notes = await notesFor({
      stem: 'Pick one',
      optionParagraphs: ['(A)\tone\t(b)\ttwo\t(C)\tthree\t(d)\tfour'],
      answer: 'A',
    });

    expect(issues(notes)).toEqual(['mixed-label-case']);
    expect(notes[0]!.detail).toContain('(A) (b) (C) (d)');
  });

  it('reports more than one four-item lettered list, saying which was used', async () => {
    const notes = await notesFor({
      stem: 'Match the columns',
      // A four-item statement list "A. ..." above the four bracketed options.
      secondLetteredList: ['statement one', 'statement two', 'statement three', 'statement four'],
      secondLetteredListId: LIST.plainStatements,
      optionParagraphs: ['first', 'second', 'third', 'fourth'],
      answer: 'B',
      autoLettered: true,
    });

    expect(issues(notes)).toEqual(['several-lettered-lists']);
    expect(notes[0]!.detail).toContain('2 lettered lists');
    expect(notes[0]!.detail).toContain('bracketed');
  });

  it('says nothing for a clean typed layout', async () => {
    const notes = await notesFor({
      stem: 'What is the SI unit of force?',
      optionParagraphs: ['(A)\tnewton\t(B)\tjoule\t(C)\twatt\t(D)\tpascal'],
      answer: 'A',
    });

    expect(notes).toEqual([]);
  });

  it('says nothing for a clean auto-lettered layout', async () => {
    const notes = await notesFor({
      stem: 'Pick the correct statement',
      optionParagraphs: ['first', 'second', 'third', 'fourth'],
      answer: 'D',
      autoLettered: true,
    });

    expect(notes).toEqual([]);
  });

  it('says nothing for a paper written consistently in lower case', async () => {
    // Reading this needs a later search pass, but the document itself is consistent.
    const notes = await notesFor({
      stem: 'Pick one',
      optionParagraphs: ['(a)\tone\t(b)\ttwo\t(c)\tthree\t(d)\tfour'],
      answer: 'A',
    });

    expect(notes).toEqual([]);
  });

  it('reports both problems when a question has two', async () => {
    const notes = await notesFor({
      stem: 'Pick one',
      // Mixed case, and the tab before (d) replaced by a comma and a space.
      optionParagraphs: ['(A)\tone\t(b)\ttwo\t(C)\tthree, (d)\tfour'],
      answer: 'A',
    });

    expect(issues(notes).sort()).toEqual(['label-not-after-tab', 'mixed-label-case']);
  });
});

describe('grouping for display', () => {
  const note = (questionNumber: number, issue: OptionLayoutIssue): QuestionLayoutNote => ({
    questionNumber,
    subject: 'PHYSICS',
    issue,
    detail: 'detail',
    fix: `fix for ${issue}`,
  });

  it('groups by problem, most affected first, with sorted question numbers', () => {
    const groups = groupLayoutNotes([
      note(9, 'label-not-after-tab'),
      note(4, 'mixed-auto-and-typed-labels'),
      note(1, 'mixed-auto-and-typed-labels'),
      note(7, 'mixed-auto-and-typed-labels'),
    ]);

    expect(groups.map((group) => group.issue)).toEqual(['mixed-auto-and-typed-labels', 'label-not-after-tab']);
    expect(groups[0]!.questionNumbers).toEqual([1, 4, 7]);
    expect(groups[0]!.fix).toBe('fix for mixed-auto-and-typed-labels');
    expect(groups[0]!.label).toContain('lettered by Word');
  });

  it('counts a question once even when it carries several notes', () => {
    const notes = [note(5, 'label-not-after-tab'), note(5, 'mixed-label-case'), note(8, 'mixed-label-case')];
    expect(questionsWithLayoutNotes(notes)).toEqual([5, 8]);
    expect(groupLayoutNotes(notes)).toHaveLength(2);
  });

  it('returns nothing for a clean paper', () => {
    expect(groupLayoutNotes([])).toEqual([]);
    expect(questionsWithLayoutNotes([])).toEqual([]);
  });
});

describe('end to end', () => {
  let workingDir = '';
  let sourceFile = '';

  const service = new GenerationService();

  beforeEach(async () => {
    workingDir = await fs.mkdtemp(path.join(os.tmpdir(), 'shuffler-notes-'));
    sourceFile = path.join(workingDir, 'Sample Paper.docx');

    // The default paper is clean; make question 3 the "(A) by Word, (B)(C)(D) typed" shape.
    const sections = defaultSections();
    sections[0]!.questions[2] = {
      stem: 'Which quantity is a vector?',
      optionParagraphs: ['speed', '(B)\tmass', '(C)\tvelocity', '(D)\ttime'],
      answer: 'C',
      letteredFirstOption: true,
    };
    await fs.writeFile(sourceFile, await buildPaper(sections));
  });

  afterEach(async () => {
    await fs.rm(workingDir, { recursive: true, force: true });
  });

  const request = (overrides: Partial<GenerationRequest> = {}): GenerationRequest => ({
    sourceFile,
    shuffleQuestions: true,
    shuffleOptions: true,
    questionExclusions: [],
    optionExclusions: [],
    setCount: 2,
    seed: 'notes',
    ...overrides,
  });

  it('surfaces the question in the dry run without skipping it', async () => {
    const report = await service.dryRun(request());

    expect(report.paper.layoutNotes.map((note) => note.questionNumber)).toEqual([3]);
    expect(report.paper.layoutNotes[0]!.issue).toBe('mixed-auto-and-typed-labels');
    expect(report.paper.layoutNotes[0]!.subject).toBe('PHYSICS');
    // Still shuffled: it is not in the "cannot be shuffled" list, and it is counted.
    expect(report.optionsKeptByTool.map((item) => item.questionNumber)).not.toContain(3);
    expect(report.optionsToShuffle).toBe(7);
  });

  it('reports it even when the user excluded that question from option shuffling', async () => {
    // The document is still worth fixing, whatever the settings say.
    const report = await service.dryRun(request({ optionExclusions: [3] }));
    expect(report.paper.layoutNotes.map((note) => note.questionNumber)).toEqual([3]);
  });

  it('still generates and verifies that question', async () => {
    const result = await service.generate(request());
    expect(result.sets[0]!.verification.ok).toBe(true);
    expect(result.paper.layoutNotes.map((note) => note.questionNumber)).toEqual([3]);
  });

  it('writes the problem, the questions and the fix into the report file', async () => {
    const result = await service.generate(request());

    // The heading proves the section reached the written PDF; the cells are asserted on
    // the document behind it, because a table cell wraps over several lines on the page.
    expect(pdfText(await fs.readFile(result.reportFile))).toContain(
      'Shuffled, but worth correcting in the Word document',
    );
    const cells = new ReportWriter()
      .build(request(), result)
      .blocks.filter((block) => block.kind === 'table')
      .flatMap((block) => (block as { rows: readonly (readonly string[])[] }).rows)
      .flat();
    expect(cells).toContain('Some options are lettered by Word and the rest typed by hand');
    expect(cells.some((cell) => cell.includes('Letter all four options the same way'))).toBe(true);
  });

  it('leaves the section out of the report for a clean paper', async () => {
    await fs.writeFile(sourceFile, await buildPaper(defaultSections()));
    const result = await service.generate(request());
    const headings = new ReportWriter()
      .build(request(), result)
      .blocks.filter((block) => block.kind === 'heading')
      .map((block) => (block as { text: string }).text);

    expect(result.paper.layoutNotes).toEqual([]);
    expect(headings).not.toContain('Shuffled, but worth correcting in the Word document');
  });

  it('does not confuse an auto-lettered option list with the statement list above it', async () => {
    // Both lists exist and both have four items, so the tool must report the choice it made.
    const sections = defaultSections();
    sections[0]!.questions[3] = {
      stem: 'Match the columns',
      secondLetteredList: ['statement one', 'statement two', 'statement three', 'statement four'],
      secondLetteredListId: LIST.plainStatements,
      optionParagraphs: ['first', 'second', 'third', 'fourth'],
      answer: 'B',
      autoLettered: true,
    };
    await fs.writeFile(sourceFile, await buildPaper(sections));

    const report = await service.dryRun(request());
    const note = report.paper.layoutNotes.find((item) => item.questionNumber === 4);
    expect(note?.issue).toBe('several-lettered-lists');
    expect(LIST.bracketedOptions).toBe('90');
  });
});
