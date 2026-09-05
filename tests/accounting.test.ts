import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import {
  OPTION_LABELS,
  QUESTION_LABELS,
  accountFor,
  accountingLine,
  unknownNumbers,
} from '../src/shared/accounting';
import { GenerationService } from '../src/core/generate/GenerationService';
import type { GenerationRequest, ShuffleAccounting } from '../src/shared/types';
import { buildPaper, defaultSections } from './support/PaperFixture';

/**
 * The dry run's arithmetic.
 *
 * A report that says "options would be shuffled for 77 of 100 questions" and then names
 * only 3 problems has left 20 questions unexplained, and reads as a contradiction even
 * though every number in it is right. These tests hold the report to a stronger promise:
 * every question of the paper is in exactly one of the three outcomes, always.
 */
describe('accountFor', () => {
  const numbers = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
  const labels = OPTION_LABELS;

  /** The promise the whole report rests on. */
  const addsUp = (accounting: ShuffleAccounting): boolean =>
    accounting.shuffled + accounting.keptByTool.length + accounting.keptByUser.length ===
    accounting.total;

  it('splits the paper into three counts that add up to it', () => {
    const accounting = accountFor({
      questionNumbers: numbers,
      active: true,
      keptByTool: [3],
      keptByUser: [7, 9],
      labels,
    });

    expect(accounting).toMatchObject({ total: 10, shuffled: 7, keptByTool: [3], keptByUser: [7, 9] });
    expect(addsUp(accounting)).toBe(true);
  });

  it('blames the tool, not the user, for a question the tool cannot do', () => {
    // Taking 3 off the user's list would change nothing, so blaming the list would send
    // the user to the wrong place - and counting it twice would break the arithmetic.
    const accounting = accountFor({
      questionNumbers: numbers,
      active: true,
      keptByTool: [3],
      keptByUser: [3, 7],
      labels,
    });

    expect(accounting.keptByTool).toEqual([3]);
    expect(accounting.keptByUser).toEqual([7]);
    expect(addsUp(accounting)).toBe(true);
  });

  it('ignores numbers this paper does not have', () => {
    const accounting = accountFor({
      questionNumbers: numbers,
      active: true,
      keptByTool: [],
      keptByUser: [7, 102, 193],
      labels,
    });

    expect(accounting.keptByUser).toEqual([7]);
    expect(accounting.shuffled).toBe(9);
    expect(addsUp(accounting)).toBe(true);
    expect(unknownNumbers(numbers, [7, 102, 193])).toEqual([102, 193]);
  });

  it('sorts and de-duplicates whatever the user typed', () => {
    const accounting = accountFor({
      questionNumbers: numbers,
      active: true,
      keptByTool: [],
      keptByUser: [9, 2, 9],
      labels,
    });

    expect(accounting.keptByUser).toEqual([2, 9]);
    expect(addsUp(accounting)).toBe(true);
  });

  it('shuffles nothing and keeps no list when the shuffle is switched off', () => {
    const accounting = accountFor({
      questionNumbers: numbers,
      active: false,
      keptByTool: [3],
      keptByUser: [7],
      labels,
    });

    expect(accounting).toEqual({ total: 10, active: false, shuffled: 0, keptByTool: [], keptByUser: [] });
    expect(accounting.summary).toBeUndefined();
  });

  it('states the sum in words, leaving out whatever is empty', () => {
    const line = (keptByTool: number[], keptByUser: number[]): string | undefined =>
      accountingLine(
        accountFor({ questionNumbers: numbers, active: true, keptByTool, keptByUser, labels }),
        labels,
      );

    expect(line([3], [7, 9])).toBe(
      '7 with their options shuffled + 1 the tool cannot read + 2 you asked to keep in order = 10',
    );
    expect(line([], [])).toBe('10 with their options shuffled = 10');
    expect(line([3], [])).toBe('9 with their options shuffled + 1 the tool cannot read = 10');
  });

  it('names the two shuffles differently, because they are kept for different reasons', () => {
    expect(QUESTION_LABELS.keptByTool).not.toBe(OPTION_LABELS.keptByTool);
    for (const labels of [QUESTION_LABELS, OPTION_LABELS]) {
      for (const label of Object.values(labels)) expect(label).not.toMatch(/^\s*$/);
    }
  });
});

describe('the dry run accounts for every question', () => {
  let workingDir = '';
  let file = '';
  const service = new GenerationService();

  beforeAll(async () => {
    workingDir = await fs.mkdtemp(path.join(os.tmpdir(), 'shuffler-accounting-'));
    file = path.join(workingDir, 'Paper.docx');
    await fs.writeFile(file, await buildPaper(defaultSections()));
  });

  afterAll(async () => {
    await fs.rm(workingDir, { recursive: true, force: true });
  });

  const request = (overrides: Partial<GenerationRequest> = {}): GenerationRequest => ({
    sourceFile: file,
    setCount: 2,
    shuffleQuestions: true,
    shuffleOptions: true,
    questionExclusions: [],
    optionExclusions: [],
    ...overrides,
  });

  it('matches the section table it prints beside it', async () => {
    const report = await service.dryRun(request({ questionExclusions: [1], optionExclusions: [5, 6] }));

    // The headline totals and the per-section columns are two views of one split, so they
    // must agree exactly - this is what makes the report readable as arithmetic.
    expect(report.questionsEligibleToMove).toBe(report.questionAccounting.shuffled);
    expect(report.optionsToShuffle).toBe(report.optionAccounting.shuffled);
    expect(report.groups.reduce((sum, group) => sum + group.questionCount, 0)).toBe(
      report.paper.questionCount,
    );
  });

  it('leaves no question unexplained, whatever the settings', async () => {
    const settings: Partial<GenerationRequest>[] = [
      {},
      { questionExclusions: [1, 2], optionExclusions: [5, 6] },
      // Numbers from some other paper, which is what a list left over from one looks like.
      { optionExclusions: [102, 112, 117] },
      { shuffleQuestions: false },
      { shuffleOptions: false },
    ];

    for (const overrides of settings) {
      const report = await service.dryRun(request(overrides));
      for (const accounting of [report.questionAccounting, report.optionAccounting]) {
        expect(accounting.total).toBe(report.paper.questionCount);
        const explained =
          accounting.shuffled + accounting.keptByTool.length + accounting.keptByUser.length;
        expect(explained).toBe(accounting.active ? accounting.total : 0);
      }
    }
  });

  it('says the sum in words for the user, and says nothing when a shuffle is off', async () => {
    const on = await service.dryRun(request());
    expect(on.questionAccounting.summary).toMatch(/free to move/);
    expect(on.optionAccounting.summary).toMatch(/options shuffled/);
    // Every summary ends in the paper's own question count, so it can be checked by eye.
    expect(on.optionAccounting.summary).toMatch(new RegExp(`= ${on.paper.questionCount}$`));

    const off = await service.dryRun(request({ shuffleOptions: false }));
    expect(off.optionAccounting.summary).toBeUndefined();
  });

  it('lists every question number of the paper, so foreign numbers can be spotted', async () => {
    const report = await service.dryRun(request());
    expect(report.paper.questionNumbers).toHaveLength(report.paper.questionCount);
    expect([...report.paper.questionNumbers].sort((a, b) => a - b)).toEqual([
      ...report.paper.questionNumbers,
    ]);
  });
});
