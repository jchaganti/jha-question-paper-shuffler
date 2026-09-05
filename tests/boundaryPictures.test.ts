/**
 * A floating picture anchored at the boundary between two questions.
 *
 * Word draws such a picture downwards from the paragraph it is anchored to, so the picture
 * at the end of question N is very often the artwork of question N+1. Nothing in the file
 * says which - so the two questions are held together instead of being separated.
 *
 * The case this exists for: "Alternating Current" question 42, whose option graphs are
 * anchored in question 41's last paragraph. Before this, shuffling the questions apart left
 * question 42 with labels and no graphs.
 */
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { DocxPackage } from '../src/core/docx/DocxPackage';
import { visibleText } from '../src/core/docx/xml';
import { GenerationService } from '../src/core/generate/GenerationService';
import { questionsPinnedByPictures, pinnedQuestionNumbers } from '../src/core/parse/BoundaryPictures';
import { NumberingIndex } from '../src/core/parse/NumberingIndex';
import { PaperParser } from '../src/core/parse/PaperParser';
import type { GenerationRequest } from '../src/shared/types';
import { buildPaper, type FixtureQuestion, type FixtureSection } from './support/PaperFixture';

/** A plain question, numbered so the fixtures below read as 1..n. */
const plain = (index: number): FixtureQuestion => ({
  stem: `Plain question ${index}`,
  optionParagraphs: ['(A)\tone\t(B)\ttwo\t(C)\tthree\t(D)\tfour'],
  answer: 'A',
});

/** Eight questions, with a floating picture at the end of the one at `at` (0-based). */
function sectionsWithPictureAfter(at: number, count = 8): FixtureSection[] {
  const questions = Array.from({ length: count }, (_unused, index) => {
    if (index !== at) return plain(index + 1);
    return {
      stem: `Question with a picture after it ${index + 1}`,
      // The picture sits in the option line itself - which is this question's last
      // paragraph, exactly as it does in the sample paper.
      optionParagraphs: ['(A)\tone\t(B)\ttwo\t(C)\tthree\t(D)\tfour[[float:graph]]'],
      answer: 'A' as const,
    };
  });
  return [{ subject: 'PHYSICS', startNumber: 1, numId: '1', questions }];
}

async function pinnedIn(sections: FixtureSection[]) {
  const pkg = await DocxPackage.fromBuffer(await buildPaper(sections));
  const paper = new PaperParser().parsePart(pkg.documentPart(), new NumberingIndex(pkg.numberingPart()));
  return questionsPinnedByPictures(paper);
}

describe('finding the pictures anchored across a question boundary', () => {
  it('holds the question with the picture and the one after it', async () => {
    const pinned = await pinnedIn(sectionsWithPictureAfter(2));

    expect(pinned).toHaveLength(1);
    expect(pinned[0]!.questionNumbers).toEqual([3, 4]);
    expect(pinned[0]!.subject).toBe('PHYSICS');
    expect(pinned[0]!.detail).toContain('question 3');
    expect(pinned[0]!.detail).toContain('question 4');
    expect(pinned[0]!.fix).toContain('In line with text');
  });

  it('holds only itself when the picture is at the end of the last question', async () => {
    const pinned = await pinnedIn(sectionsWithPictureAfter(7));

    expect(pinned).toHaveLength(1);
    expect(pinned[0]!.questionNumbers).toEqual([8]);
  });

  it('says nothing about a picture that is not in the last paragraph', async () => {
    // A picture among the options but with an option line after it is drawn over this
    // question's own paragraphs, which travel with it. Nothing to hold.
    const sections: FixtureSection[] = [
      {
        subject: 'PHYSICS',
        startNumber: 1,
        numId: '1',
        questions: [
          {
            stem: 'A picture between the option lines',
            optionParagraphs: ['(A)\tone[[float:graph]]\t(B)\ttwo', '(C)\tthree\t(D)\tfour'],
            answer: 'A',
          },
          ...Array.from({ length: 6 }, (_unused, i) => plain(i + 2)),
        ],
      },
    ];

    expect(await pinnedIn(sections)).toEqual([]);
  });

  it('says nothing about a paper with no floating pictures', async () => {
    const sections: FixtureSection[] = [
      { subject: 'PHYSICS', startNumber: 1, numId: '1', questions: Array.from({ length: 7 }, (_u, i) => plain(i + 1)) },
    ];

    expect(await pinnedIn(sections)).toEqual([]);
  });

  it('lists every held question once, in order', () => {
    const numbers = pinnedQuestionNumbers([
      { questionNumbers: [8, 9], subject: 'PHYSICS', detail: '', fix: '' },
      { questionNumbers: [3, 4], subject: 'PHYSICS', detail: '', fix: '' },
      { questionNumbers: [4, 5], subject: 'PHYSICS', detail: '', fix: '' },
    ]);

    expect(numbers).toEqual([3, 4, 5, 8, 9]);
  });
});

describe('generating a paper that has one', () => {
  let workingDir = '';
  let sourceFile = '';
  const service = new GenerationService();

  const request = (overrides: Partial<GenerationRequest> = {}): GenerationRequest => ({
    sourceFile,
    shuffleQuestions: true,
    shuffleOptions: true,
    questionExclusions: [],
    optionExclusions: [],
    setCount: 2,
    seed: 'boundary',
    ...overrides,
  });

  beforeEach(async () => {
    workingDir = await fs.mkdtemp(path.join(os.tmpdir(), 'shuffler-boundary-'));
    sourceFile = path.join(workingDir, 'Sample Paper.docx');
    await fs.writeFile(sourceFile, await buildPaper(sectionsWithPictureAfter(2)));
  });

  afterEach(async () => {
    await fs.rm(workingDir, { recursive: true, force: true });
  });

  /** The stem of each question of a generated paper, in printed order. */
  async function stems(file: string): Promise<string[]> {
    const pkg = await DocxPackage.fromBuffer(await fs.readFile(file));
    const paper = new PaperParser().parsePart(pkg.documentPart(), new NumberingIndex(pkg.numberingPart()));
    return paper.sections.flatMap((section) =>
      section.blocks.map((block) => visibleText(block.questionParagraph).replace(/\s+/g, ' ').trim()),
    );
  }

  it('keeps both questions at their original positions, with the rest shuffled around them', async () => {
    const before = await stems(sourceFile);
    const result = await service.generate(request({ setCount: 4 }));

    for (const set of result.sets) {
      const after = await stems(set.filePath);
      expect(after[2], set.fileName).toBe(before[2]);
      expect(after[3], set.fileName).toBe(before[3]);
      // The point of holding only the pair: everything else still moves.
      expect(after).not.toEqual(before);
    }
  });

  it('reports the pair, and counts them out of the questions free to move', async () => {
    const report = await service.dryRun(request());

    expect(report.questionsKeptByTool).toHaveLength(1);
    expect(report.questionsKeptByTool[0]!.questionNumbers).toEqual([3, 4]);
    expect(report.groups[0]!.movable).toBe(report.paper.questionCount - 2);
  });

  it('says nothing about them when questions are not being shuffled', async () => {
    const report = await service.dryRun(request({ shuffleQuestions: false }));

    expect(report.questionsKeptByTool).toEqual([]);
    // Still known to the paper summary, which is about the paper rather than the request.
    expect(report.paper.pinnedQuestions).toHaveLength(1);
  });

  it('does not report them as something the user asked for', async () => {
    const report = await service.dryRun(request());

    // A question the tool pins counts as the tool's doing, never as the user's - and its
    // number is never reported back as one this paper does not have.
    expect(report.questionAccounting.keptByUser).toEqual([]);
    expect(report.questionAccounting.keptByTool.length).toBeGreaterThan(0);
    expect(report.optionAccounting.keptByUser).toEqual([]);
    expect(report.warnings.join(' ')).not.toContain('does not have');
  });

  it('still verifies', async () => {
    const result = await service.generate(request({ setCount: 2 }));
    for (const set of result.sets) {
      expect(set.verification.ok, JSON.stringify(set.verification.checks)).toBe(true);
    }
  });
});
