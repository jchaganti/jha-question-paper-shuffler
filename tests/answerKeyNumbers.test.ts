/**
 * How a question number may be written in the answer key.
 *
 * Authors habitually type the ordinal punctuation after the number - "1." or "1)" - and
 * that names question 1 and nothing else, so it is read. A stray character is a slip
 * rather than a convention, so it is reported by the box it spoilt instead of being read
 * through: the tool never decides for itself which characters after a number are
 * decoration and which carry meaning.
 */
import { describe, expect, it } from 'vitest';
import { DocxPackage } from '../src/core/docx/DocxPackage';
import { NumberingIndex } from '../src/core/parse/NumberingIndex';
import { PaperParser, PaperParseError } from '../src/core/parse/PaperParser';
import { buildPaper, defaultSections, type FixtureOptions } from './support/PaperFixture';

async function parse(options: FixtureOptions) {
  const pkg = await DocxPackage.fromBuffer(await buildPaper(defaultSections(), options));
  return new PaperParser().parsePart(pkg.documentPart(), new NumberingIndex(pkg.numberingPart()));
}

describe('a question number written with its ordinal punctuation', () => {
  it('is read when the key writes "1."', async () => {
    const paper = await parse({ answerKeyNumberSuffix: '.' });

    expect(paper.answerKey.questionNumbers).toEqual([1, 2, 3, 4, 5, 6, 7]);
    expect(paper.answerKey.answerOf(3)).toBe('C');
  });

  it('is read when the key writes "1)"', async () => {
    const paper = await parse({ answerKeyNumberSuffix: ')' });

    expect(paper.answerKey.questionNumbers).toEqual([1, 2, 3, 4, 5, 6, 7]);
    expect(paper.answerKey.answerOf(5)).toBe('A');
  });

  it('still writes the new letter back into the right box', async () => {
    const paper = await parse({ answerKeyNumberSuffix: '.' });

    paper.answerKey.setAnswer(3, 'D');

    expect(paper.answerKey.answerOf(3)).toBe('D');
    expect(paper.answerKey.answerOf(2)).toBe('B');
  });
});

describe('a question number spoilt by a stray character', () => {
  it('names the question the key has lost', async () => {
    await expect(parse({ answerKeyNumberSuffix: '.', spoiltKeyNumbers: [3] })).rejects.toThrow(
      /answer key is incomplete[\s\S]*no entry for question 3/,
    );
  });

  it('quotes the box to correct, rather than reading through it', async () => {
    await expect(parse({ answerKeyNumberSuffix: '.', spoiltKeyNumbers: [3] })).rejects.toThrow(
      /"3," - a stray character was typed after the number/,
    );
  });

  it('is reported as a key fault, not as unrecognised question numbering', async () => {
    // The numbering is perfectly good here; only the key is short. Reporting the numbering
    // would send the author looking in the wrong half of the document.
    const error = await parse({ answerKeyNumberSuffix: '.', spoiltKeyNumbers: [3] }).catch((e) => e);

    expect(error).toBeInstanceOf(PaperParseError);
    expect(String(error)).not.toMatch(/numbering/i);
  });

  it('reports every spoilt box, not just the first', async () => {
    const error = await parse({ answerKeyNumberSuffix: '.', spoiltKeyNumbers: [3, 6] }).catch((e) => e);

    expect(String(error)).toMatch(/question 3 and 6/);
    // In question order, not the order the columns happen to be scanned in.
    expect(String(error)).toMatch(/Number boxes in the key read "3," and "6,"/);
  });
});
