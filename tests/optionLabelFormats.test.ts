/**
 * How a paper labels its four options.
 *
 * Papers do not agree: most write letters `(A)...(D)`, but whole papers are written with
 * digits `(1)...(4)` or roman numerals `(i)...(iv)`. The reader deduces which scheme the
 * question in front of it uses instead of assuming letters, and the labels themselves are
 * never moved, so each set comes back written the way its author wrote it.
 *
 * The schemes are tried letters, then digits, then roman numerals - the order of
 * decreasing certainty that a run of labels really is the options. Questions routinely
 * list *items* as `(i)...(iv)` or `(a)...(d)` above answers written another way, so the
 * scheme most easily confused with an item list is asked last. The tests below pin both
 * halves: the schemes that must be read, and the item lists that must not be mistaken for
 * options.
 */
import { describe, expect, it } from 'vitest';
import { DocxPackage } from '../src/core/docx/DocxPackage';
import { visibleText } from '../src/core/docx/xml';
import { OptionSetParser } from '../src/core/options/OptionSetParser';
import { NumberingIndex } from '../src/core/parse/NumberingIndex';
import { PaperParser } from '../src/core/parse/PaperParser';
import { buildPaper, type FixtureQuestion, type FixtureSection } from './support/PaperFixture';

/** A paper whose first question is the one under test, padded so the key is detectable. */
function paperWith(question: FixtureQuestion): FixtureSection[] {
  const filler = Array.from({ length: 6 }, (_unused, index) => ({
    stem: `Filler question ${index + 1}`,
    optionParagraphs: ['(A)\tone\t(B)\ttwo\t(C)\tthree\t(D)\tfour'],
    answer: 'A' as const,
  }));
  return [{ subject: 'PHYSICS', startNumber: 1, numId: '1', questions: [question, ...filler] }];
}

async function readFirstQuestion(question: FixtureQuestion) {
  const pkg = await DocxPackage.fromBuffer(await buildPaper(paperWith(question)));
  const numbering = new NumberingIndex(pkg.numberingPart());
  const paper = new PaperParser().parsePart(pkg.documentPart(), numbering);
  return new OptionSetParser(numbering).parse(paper.sections[0]!.blocks[0]!);
}

/** Option contents, in slot order A..D, as plain text. */
function contents(parsed: Awaited<ReturnType<typeof readFirstQuestion>>): string[] {
  if (!parsed.ok) throw new Error(`expected the options to parse, got ${parsed.reason}: ${parsed.detail}`);
  return parsed.options.signatures.map((signature) => signature.split('|')[0]!);
}

describe.each([
  ['digits', '(1)\tnewton\t(2)\tjoule\t(3)\twatt\t(4)\tpascal', '(5)\tnone of these'],
  ['digits without an opening bracket', '1)\tnewton\t2)\tjoule\t3)\twatt\t4)\tpascal', '5)\tnone of these'],
  ['lower-case roman numerals', '(i)\tnewton\t(ii)\tjoule\t(iii)\twatt\t(iv)\tpascal', '(v)\tnone of these'],
  ['upper-case roman numerals', '(I)\tnewton\t(II)\tjoule\t(III)\twatt\t(IV)\tpascal', '(V)\tnone of these'],
])('options labelled with %s', (_label, optionParagraph, fifthOption) => {
  it('reads all four options in order', async () => {
    const parsed = await readFirstQuestion({
      stem: 'What is the SI unit of force?',
      optionParagraphs: [optionParagraph],
      answer: 'A',
    });

    expect(contents(parsed)).toEqual(['newton', 'joule', 'watt', 'pascal']);
  });

  it('refuses a fifth option rather than folding it into the fourth', async () => {
    // The fifth label must be recognised as a label of this scheme, or its text would ride
    // along with option four and be moved around with it.
    const parsed = await readFirstQuestion({
      stem: 'What is the SI unit of force?',
      optionParagraphs: [optionParagraph, fifthOption],
      answer: 'A',
    });

    expect(parsed.ok).toBe(false);
    if (parsed.ok) return;
    expect(parsed.reason).toBe('unexpected-label-sequence');
  });
});

describe('a paper that labels its options with digits', () => {
  it('shuffles the contents while the "(1)...(4)" labels stay where they are', async () => {
    const pkg = await DocxPackage.fromBuffer(
      await buildPaper(
        paperWith({
          stem: 'What is the SI unit of force?',
          optionParagraphs: ['(1)\tnewton\t(2)\tjoule\t(3)\twatt\t(4)\tpascal'],
          answer: 'A',
        }),
      ),
    );
    const numbering = new NumberingIndex(pkg.numberingPart());
    const paper = new PaperParser().parsePart(pkg.documentPart(), numbering);
    const block = paper.sections[0]!.blocks[0]!;
    const parsed = new OptionSetParser(numbering).parse(block);
    if (!parsed.ok) throw new Error('expected the options to parse');

    parsed.options.apply([3, 2, 1, 0]);

    const text = block.nodes
      .filter((node) => node !== block.questionParagraph)
      .map((node) => visibleText(node))
      .join(' ');
    // Same labels, in the same order, with the contents reversed behind them.
    expect(text.replace(/\s+/g, ' ').trim()).toBe('(1) pascal (2) watt (3) joule (4) newton');
  });
});

describe('an item list that must not be mistaken for the options', () => {
  it('prefers the lettered answers over a roman item list in the same paragraph', async () => {
    // "(A) (i) and (ii) ... (D) (iii) and (v)" - a full run of roman numerals sits in the
    // option *contents*, but the answers are lettered, so the letters win.
    const parsed = await readFirstQuestion({
      stem: 'Which statements are incorrect?',
      optionParagraphs: ['(A)\t(i) and (ii)\t(B)\t(iii) and (iv)\t(C)\t(i), (iii) and (iv)\t(D)\t(ii) and (iv)'],
      answer: 'A',
    });

    expect(contents(parsed)).toEqual(['(i) and (ii)', '(iii) and (iv)', '(i), (iii) and (iv)', '(ii) and (iv)']);
  });

  it('prefers the lettered answers over a digit item list in the same paragraph', async () => {
    const parsed = await readFirstQuestion({
      stem: 'Which statements are correct?',
      optionParagraphs: ['(A)\t(1) and (2)\t(B)\t(3) and (4)\t(C)\t(1), (3) and (4)\t(D)\t(2) and (4)'],
      answer: 'A',
    });

    expect(contents(parsed)).toEqual(['(1) and (2)', '(3) and (4)', '(1), (3) and (4)', '(2) and (4)']);
  });

  it('prefers the digit answers over a roman item list above them', async () => {
    const parsed = await readFirstQuestion({
      stem: 'Match the columns',
      optionParagraphs: [
        '(i)\tcopper\t(ii)\tzinc\t(iii)\tiron\t(iv)\tlead',
        '(1)\ti-a, ii-b\t(2)\ti-b, ii-c\t(3)\ti-c, ii-d\t(4)\ti-d, ii-a',
      ],
      answer: 'A',
    });

    expect(contents(parsed)).toEqual(['i-a, ii-b', 'i-b, ii-c', 'i-c, ii-d', 'i-d, ii-a']);
  });
});
