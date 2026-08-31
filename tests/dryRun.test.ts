import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { GenerationService } from '../src/core/generate/GenerationService';
import type { GenerationRequest, ProgressEvent } from '../src/shared/types';
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
  setCount: 3,
  seed: 'dry-run',
  ...overrides,
});

beforeEach(async () => {
  workingDir = await fs.mkdtemp(path.join(os.tmpdir(), 'shuffler-dry-'));
  sourceFile = path.join(workingDir, 'Sample Paper.docx');
  await fs.writeFile(sourceFile, await buildPaper(defaultSections()));
});

afterEach(async () => {
  await fs.rm(workingDir, { recursive: true, force: true });
});

describe('dryRun', () => {
  it('writes nothing at all', async () => {
    await service.dryRun(request());
    expect(await fs.readdir(workingDir)).toEqual(['Sample Paper.docx']);
  });

  // The fixture has 7 questions, all of them shufflable: question 4 is auto-lettered by
  // Word rather than carrying typed "(A)" labels, which is supported too.
  it('reports what would be shuffled with the current settings', async () => {
    const report = await service.dryRun(request());

    expect(report.setCount).toBe(3);
    expect(report.questionsEligibleToMove).toBe(7);
    expect(report.optionsToShuffle).toBe(7);
    expect(report.optionsKeptByTool).toEqual([]);
    expect(report.subjects).toEqual([
      { subject: 'PHYSICS', questionCount: 4, movable: 4, optionsShuffled: 4 },
      { subject: 'CHEMISTRY', questionCount: 3, movable: 3, optionsShuffled: 3 },
    ]);
  });

  it('subtracts the exclusion lists', async () => {
    const report = await service.dryRun(request({ questionExclusions: [1, 2], optionExclusions: [5, 6] }));

    expect(report.questionsEligibleToMove).toBe(5);
    expect(report.optionsToShuffle).toBe(5);
    expect(report.optionsKeptByUser).toEqual([5, 6]);
    expect(report.subjects[0]).toMatchObject({ subject: 'PHYSICS', movable: 2 });
    expect(report.subjects[1]).toMatchObject({ subject: 'CHEMISTRY', optionsShuffled: 1 });
  });

  it('reports nothing movable or shuffled when a toggle is off', async () => {
    expect((await service.dryRun(request({ shuffleQuestions: false }))).questionsEligibleToMove).toBe(0);
    expect((await service.dryRun(request({ shuffleOptions: false }))).optionsToShuffle).toBe(0);
  });

  it('suggests position-dependent questions and stops suggesting once excluded', async () => {
    const before = await service.dryRun(request());
    // Question 5 has a "None of these" option.
    expect(before.suggestedForExclusion.map((item) => item.questionNumber)).toContain(5);

    const after = await service.dryRun(request({ optionExclusions: [5] }));
    expect(after.suggestedForExclusion.map((item) => item.questionNumber)).not.toContain(5);
  });

  it('never suggests a question whose options are already left alone', async () => {
    const report = await service.dryRun(request());
    const skipped = report.optionsKeptByTool.map((item) => item.questionNumber);
    for (const suggestion of report.suggestedForExclusion) {
      expect(skipped).not.toContain(suggestion.questionNumber);
    }
  });

  it('warns about exclusion numbers that do not exist in the paper', async () => {
    const report = await service.dryRun(request({ questionExclusions: [3, 400], optionExclusions: [999] }));
    expect(report.warnings.join(' ')).toMatch(/no question 400/);
    expect(report.warnings.join(' ')).toMatch(/no question 999/);
    expect(report.warnings.join(' ')).not.toMatch(/no question 3\b/);
  });

  it('warns when no subject headings were recognised', async () => {
    const [physics, chemistry] = defaultSections();
    const file = path.join(workingDir, 'No headings.docx');
    await fs.writeFile(
      file,
      await buildPaper([
        { ...physics!, subject: 'Part I' },
        { ...chemistry!, subject: 'Part II' },
      ]),
    );

    const report = await service.dryRun(request({ sourceFile: file }));
    expect(report.warnings.join(' ')).toMatch(/No subject headings were recognised/);
  });

  it('validates the request just like generate does', async () => {
    await expect(service.dryRun(request({ setCount: 0 }))).rejects.toThrow(/between 1 and 100/);
    await expect(
      service.dryRun(request({ shuffleQuestions: false, shuffleOptions: false })),
    ).rejects.toThrow(/at least one/);
  });

  it('predicts the option count that generate then produces', async () => {
    const settings = request({ setCount: 1, optionExclusions: [5] });
    const report = await service.dryRun(settings);
    const result = await service.generate(settings);
    expect(result.sets[0]!.optionsShuffled).toBe(report.optionsToShuffle);
  });

  it('names options the way the paper names them, not always A-D', async () => {
    // A paper written "(1) (2) (3) (4)" must never be told about an "Option (A)": the
    // reader would go looking for a label that is not on the page.
    await fs.writeFile(
      sourceFile,
      await buildPaper(
        [
          {
            subject: 'PHYSICS',
            startNumber: 1,
            numId: '1',
            questions: [
              // Three options, so the reason names the label it could not find.
              { stem: 'A short question', optionParagraphs: ['(1)\tone\t(2)\ttwo\t(3)\tthree'], answer: 'A' },
              ...Array.from({ length: 6 }, (_unused, index) => ({
                stem: `Filler question ${index + 1}`,
                optionParagraphs: ['(1)\tone\t(2)\ttwo\t(3)\tthree\t(4)\tfour'],
                answer: 'A' as const,
              })),
            ],
          },
        ],
        { answerKeyAnswerFormat: 'digit' },
      ),
    );

    const report = await service.dryRun(request({ setCount: 1 }));
    const details = report.optionsKeptByTool.map((item) => item.detail).join(' ');

    expect(report.optionsKeptByTool.length).toBeGreaterThan(0);
    expect(details).toMatch(/1,2,3,4/);
    expect(details).not.toMatch(/\(?[ABCD]\)?[,)]/);
  });
});

describe('progress reporting', () => {
  it('reports every set, in order, ending at 100%', async () => {
    const events: ProgressEvent[] = [];
    const result = await service.generate(request({ setCount: 3 }), (event) => events.push(event));

    expect(result.sets).toHaveLength(3);

    const overall = events.map((event) => event.fraction ?? 0);
    for (let i = 1; i < overall.length; i++) {
      expect(overall[i]!, `event ${i} went backwards`).toBeGreaterThanOrEqual(overall[i - 1]!);
    }
    expect(overall[overall.length - 1]).toBe(1);

    for (const setNumber of [1, 2, 3]) {
      const forSet = events.filter((event) => event.setNumber === setNumber && event.stage !== 'done');
      expect(forSet.length, `set ${setNumber} reported no progress`).toBeGreaterThan(0);
      expect(forSet.every((event) => event.setCount === 3)).toBe(true);
      expect(Math.max(...forSet.map((event) => event.setFraction ?? 0))).toBe(1);
    }

    expect(events.map((event) => event.stage)).toContain('planning');
    expect(events.map((event) => event.stage)).toContain('options');
    expect(events.map((event) => event.stage)).toContain('ordering');
    expect(events.map((event) => event.stage)).toContain('writing');
    expect(events.map((event) => event.stage)).toContain('verifying');
    expect(events[events.length - 1]!.stage).toBe('done');
  });

  it('keeps each set inside its own slice of the overall bar', async () => {
    const events: ProgressEvent[] = [];
    await service.generate(request({ setCount: 2 }), (event) => events.push(event));

    for (const event of events) {
      if (!event.setNumber || event.setFraction === undefined || event.stage === 'done') continue;
      const lower = (event.setNumber - 1) / 2;
      expect(event.fraction!).toBeGreaterThanOrEqual(lower);
      expect(event.fraction!).toBeLessThanOrEqual(event.setNumber / 2);
    }
  });
});
