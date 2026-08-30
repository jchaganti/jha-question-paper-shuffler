/**
 * How the answer key names the correct option.
 *
 * Papers do not agree on this: some write the bare letter ("A"), some bracket it ("(A)"),
 * some write it in lower case, and some name the option by its *position* instead of its
 * letter - a digit ("1", "2.") or a roman numeral ("i)", "(iv)"). None of these is a
 * "wrong" way to fill in a key, so the tool deduces which one a given paper uses from the
 * box itself and writes the new answer back in that same style - it never imposes a house
 * style of its own.
 */
import { describe, expect, it } from 'vitest';
import { DocxPackage } from '../src/core/docx/DocxPackage';
import type { Element } from '../src/core/docx/dom';
import { NS, childElements, visibleText } from '../src/core/docx/xml';
import { NumberingIndex } from '../src/core/parse/NumberingIndex';
import { PaperParser } from '../src/core/parse/PaperParser';
import { buildPaper, defaultSections, type AnswerKeyAnswerFormat, type FixtureOptions } from './support/PaperFixture';

async function parse(options: FixtureOptions) {
  const pkg = await DocxPackage.fromBuffer(await buildPaper(defaultSections(), options));
  return new PaperParser().parsePart(pkg.documentPart(), new NumberingIndex(pkg.numberingPart()));
}

/**
 * The exact text in the box next to question `questionNumber`, whichever key table holds
 * it. Steps in number/answer pairs: in a digit key the *answers* are numbers too, so
 * scanning every cell would match an answer box as if it were a question number.
 */
function answerCellText(tables: readonly Element[], questionNumber: number): string | undefined {
  for (const table of tables) {
    for (const row of childElements(table, NS.w, 'tr')) {
      const cells = childElements(row, NS.w, 'tc');
      for (let i = 0; i + 1 < cells.length; i += 2) {
        const match = /^(\d{1,4})/.exec(visibleText(cells[i]!).trim());
        if (match && Number(match[1]) === questionNumber) return visibleText(cells[i + 1]!).trim();
      }
    }
  }
  return undefined;
}

describe.each<[string, AnswerKeyAnswerFormat, string, string]>([
  // [label, format, question 3's box (answer C), question 1's box (answer A)]
  ['a bare upper-case letter', 'letter', 'C', 'A'],
  ['a bare lower-case letter', 'letter-lower', 'c', 'a'],
  ['a bracketed letter', 'letter-bracketed', '(C)', '(A)'],
  ['a bare digit naming the option position', 'digit', '3', '1'],
  ['a digit with ordinal punctuation', 'digit-dot', '3.', '1.'],
  ['a roman numeral naming the option position', 'roman', 'iii)', 'i)'],
  ['a bracketed roman numeral', 'roman-bracketed', '(iii)', '(i)'],
])('a key that writes its answer as %s', (_label, format, expectedQ3Text, expectedQ1Text) => {
  it('is read as the letter it names', async () => {
    const paper = await parse({ answerKeyAnswerFormat: format });

    // Question 3's answer is C in every fixture section (see defaultSections).
    expect(paper.answerKey.answerOf(3)).toBe('C');
    // The literal box confirms the fixture actually wrote the format under test.
    expect(answerCellText(paper.answerKey.tables, 3)).toBe(expectedQ3Text);
  });

  it('writes the new answer back in the same style', async () => {
    const paper = await parse({ answerKeyAnswerFormat: format });

    paper.answerKey.setAnswer(3, 'D');

    expect(paper.answerKey.answerOf(3)).toBe('D');
    // Untouched boxes keep both their answer and their style.
    expect(paper.answerKey.answerOf(1)).toBe('A');
    expect(answerCellText(paper.answerKey.tables, 1)).toBe(expectedQ1Text);
  });
});

describe('a key whose answers name the option by position', () => {
  it('maps every digit to the letter at that position', async () => {
    const paper = await parse({ answerKeyAnswerFormat: 'digit' });

    expect(paper.answerKey.answerOf(1)).toBe('A');
    expect(paper.answerKey.answerOf(2)).toBe('B');
    expect(paper.answerKey.answerOf(3)).toBe('C');
    expect(paper.answerKey.answerOf(4)).toBe('D');
  });

  it('maps every roman numeral to the letter at that position', async () => {
    const paper = await parse({ answerKeyAnswerFormat: 'roman-bracketed' });

    expect(paper.answerKey.answerOf(1)).toBe('A');
    expect(paper.answerKey.answerOf(2)).toBe('B');
    expect(paper.answerKey.answerOf(3)).toBe('C');
    expect(paper.answerKey.answerOf(4)).toBe('D');
  });
});

describe('setAnswer rendering', () => {
  it('renders every option letter in a roman-numeral key', async () => {
    const paper = await parse({ answerKeyAnswerFormat: 'roman' });

    paper.answerKey.setAnswer(1, 'A');
    paper.answerKey.setAnswer(2, 'B');
    paper.answerKey.setAnswer(3, 'C');
    paper.answerKey.setAnswer(4, 'D');

    expect(answerCellText(paper.answerKey.tables, 1)).toBe('i)');
    expect(answerCellText(paper.answerKey.tables, 2)).toBe('ii)');
    expect(answerCellText(paper.answerKey.tables, 3)).toBe('iii)');
    expect(answerCellText(paper.answerKey.tables, 4)).toBe('iv)');
  });

  it('renders every option position in a digit key', async () => {
    const paper = await parse({ answerKeyAnswerFormat: 'digit-dot' });

    paper.answerKey.setAnswer(1, 'D');
    paper.answerKey.setAnswer(2, 'C');
    paper.answerKey.setAnswer(3, 'B');
    paper.answerKey.setAnswer(4, 'A');

    expect(answerCellText(paper.answerKey.tables, 1)).toBe('4.');
    expect(answerCellText(paper.answerKey.tables, 2)).toBe('3.');
    expect(answerCellText(paper.answerKey.tables, 3)).toBe('2.');
    expect(answerCellText(paper.answerKey.tables, 4)).toBe('1.');
  });
});
