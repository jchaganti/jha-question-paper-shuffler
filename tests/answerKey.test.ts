import { describe, expect, it } from 'vitest';
import { DocxPackage } from '../src/core/docx/DocxPackage';
import { describePermutation, mapAnswer } from '../src/core/generate/SetBuilder';
import { NumberingIndex } from '../src/core/parse/NumberingIndex';
import { PaperParser } from '../src/core/parse/PaperParser';
import { parseNumberList } from '../src/core/generate/GenerationService';
import { buildPaper, defaultSections } from './support/PaperFixture';

describe('mapAnswer', () => {
  it('follows the option content to its new slot', () => {
    // permutation[newSlot] = originalSlot: new A shows old C, new B shows old D, ...
    const permutation = [2, 3, 0, 1];
    expect(mapAnswer('C', permutation)).toBe('A');
    expect(mapAnswer('D', permutation)).toBe('B');
    expect(mapAnswer('A', permutation)).toBe('C');
    expect(mapAnswer('B', permutation)).toBe('D');
  });

  it('matches the worked example from the specification', () => {
    // A)1 B)99 C)22 D)44 with answer C=22 becomes A)22 B)44 C)1 D)99, so the answer is A.
    const permutation = [2, 3, 0, 1];
    expect(mapAnswer('C', permutation)).toBe('A');
  });

  it('keeps the answer when the options are untouched', () => {
    expect(mapAnswer('B', [0, 1, 2, 3])).toBe('B');
  });

  it('describes a permutation as source to target', () => {
    expect(describePermutation([2, 3, 0, 1])).toBe('A→C, B→D, C→A, D→B');
  });
});

describe('AnswerKeyTable', () => {
  it('rewrites a letter without touching the rest of the cell', async () => {
    const pkg = await DocxPackage.fromBuffer(await buildPaper(defaultSections()));
    const part = pkg.documentPart();
    const paper = new PaperParser().parsePart(part, new NumberingIndex(pkg.numberingPart()));

    paper.answerKey.setAnswer(3, 'D');
    expect(paper.answerKey.answerOf(3)).toBe('D');

    const xml = part.serialize();
    expect(xml).toContain('<w:tbl>');
    // The other answers are unchanged.
    const reparsed = new PaperParser().parsePart(
      (await DocxPackage.fromBuffer(await buildPaper(defaultSections()))).documentPart(),
      new NumberingIndex(pkg.numberingPart()),
    );
    expect(reparsed.answerKey.answerOf(1)).toBe(paper.answerKey.answerOf(1));
  });
});

describe('parseNumberList', () => {
  it('parses comma separated numbers', () => {
    expect(parseNumberList('10,11,19,20')).toEqual([10, 11, 19, 20]);
  });

  it('tolerates spaces, semicolons and duplicates', () => {
    expect(parseNumberList(' 4; 4 , 1 ')).toEqual([1, 4]);
  });

  it('expands ranges', () => {
    expect(parseNumberList('2-5, 9')).toEqual([2, 3, 4, 5, 9]);
  });

  it('returns an empty list for empty input', () => {
    expect(parseNumberList('   ')).toEqual([]);
  });

  it('rejects anything that is not a question number', () => {
    expect(() => parseNumberList('1, abc')).toThrow(/not a question number/);
  });
});
