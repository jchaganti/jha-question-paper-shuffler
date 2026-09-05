/**
 * The label-finding rules that let real, inconsistently typed papers be read.
 *
 * Each rule exists because a specific shape appeared in the sample papers, and each test
 * also pins the *limit* of the rule - the case it must keep refusing, so that a relaxed
 * reading never invents a wrong option boundary.
 */
import { describe, expect, it } from 'vitest';
import { DocxPackage } from '../src/core/docx/DocxPackage';
import { NS, firstChild, visibleText } from '../src/core/docx/xml';
import { OptionSetParser } from '../src/core/options/OptionSetParser';
import { NumberingIndex } from '../src/core/parse/NumberingIndex';
import { PaperParser } from '../src/core/parse/PaperParser';
import type { QuestionBlock } from '../src/core/parse/PaperModel';
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
  const block = paper.sections[0]!.blocks[0]!;
  return { block, parsed: new OptionSetParser(numbering).parse(block) };
}

/** Option contents, A..D, as plain text. */
function contents(parsed: Awaited<ReturnType<typeof readFirstQuestion>>['parsed']): string[] {
  if (!parsed.ok) throw new Error(`expected the options to parse, got ${parsed.reason}: ${parsed.detail}`);
  return parsed.options.signatures.map((signature) => signature.split('|')[0]!);
}

const visibleOptionText = (block: QuestionBlock): string =>
  block.nodes
    .filter((node) => node !== block.questionParagraph && node.namespaceURI === NS.w && node.localName === 'p')
    .map((node) => visibleText(node))
    .join(' | ')
    .replace(/\t/g, '›');

describe('a label after punctuation instead of a tab', () => {
  // "(A) (i) and (ii) (B) ... (C) (i), (iv) and (v) (D) ..." - the tab before (D) is missing.
  it('is accepted when the label follows a non-letter', async () => {
    const { parsed } = await readFirstQuestion({
      stem: 'Which statements are incorrect?',
      optionParagraphs: ['(A)\t(i) and (ii)\t(B)\t(i) and (iii)\t(C)\t(i), (iv) and (v) (D)\t(iii) and (v)'],
      answer: 'A',
    });

    expect(contents(parsed)).toEqual(['(i) and (ii)', '(i) and (iii)', '(i), (iv) and (v)', '(iii) and (v)']);
  });

  it('is still refused when a letter comes first, so option cross-references survive', async () => {
    const { parsed } = await readFirstQuestion({
      stem: 'Which reactant is limiting?',
      optionParagraphs: ['(A)\tCaCO3\t(B)\tHCl', '(C)\tBoth (A) and (B)\t(D)\tNone of these'],
      answer: 'C',
    });

    // Four options, and the "(A)"/"(B)" inside option C are part of its text, not labels.
    expect(contents(parsed)).toEqual(['caco3', 'hcl', 'both (a) and (b)', 'none of these']);
  });

  it('is still refused when the label is glued to a digit', async () => {
    // The shape of a floating equation picture whose text box runs into the label.
    const { parsed } = await readFirstQuestion({
      stem: 'Which one is a wrong statement?',
      optionParagraphs: ['(A)\tfirst', '(B)\tsecond', '1s22s2(C)\tthird', '(D)\tfourth'],
      answer: 'A',
    });

    expect(parsed.ok).toBe(false);
    if (parsed.ok) return;
    expect(parsed.reason).toBe('unexpected-label-sequence');
    expect(parsed.detail).toContain('"ABD"');
  });
});

describe('two options on one line pushed apart with spaces instead of a tab', () => {
  // "…High power loss in transmission              (B) Low power loss in transmission".
  // The labels are all present and in order; only the separator in front of them is wrong.
  it('is refused, naming the labels to put a tab in front of', async () => {
    const { parsed } = await readFirstQuestion({
      stem: 'If the power factor is low, it means:',
      optionParagraphs: [
        '(A)\tHigh power loss              (B)\tLow power loss',
        '(C)\tMay be high or low              (D)\tNone of these',
      ],
      answer: 'A',
    });

    expect(parsed.ok).toBe(false);
    if (parsed.ok) return;
    expect(parsed.reason).toBe('label-after-spaces-not-tab');
    expect(parsed.detail).toContain('(B), (D)');
    expect(parsed.fix).toContain('press Tab');
  });

  it('is not blamed when the space run does not account for every option', async () => {
    // (C) is glued to the digits of a floating equation, which is the real fault; the space
    // run before (D) must not be reported as though fixing it would make the question read.
    const { parsed } = await readFirstQuestion({
      stem: 'Which one is a wrong statement?',
      optionParagraphs: ['(A)\tfirst', '(B)\tsecond', '1s22s2(C)\tthird              (D)\tfourth'],
      answer: 'A',
    });

    expect(parsed.ok).toBe(false);
    if (parsed.ok) return;
    expect(parsed.reason).toBe('unexpected-label-sequence');
  });

  it('leaves a single space in prose alone, so option cross-references still read', async () => {
    const { parsed } = await readFirstQuestion({
      stem: 'Which reactant is limiting?',
      optionParagraphs: ['(A)\tCaCO3\t(B)\tHCl', '(C)\tBoth (A) and (B)\t(D)\tNone of these'],
      answer: 'C',
    });

    expect(contents(parsed)).toEqual(['caco3', 'hcl', 'both (a) and (b)', 'none of these']);
  });
});

describe('upper case labels are preferred over lower case', () => {
  // A match-the-columns question lists items as "(a)...(d)" and answers as "(A)...(D)".
  it('reads the answer options, not the item list', async () => {
    const { parsed } = await readFirstQuestion({
      stem: 'Match the following',
      optionParagraphs: [
        '(a)\tSilica cell wall\t(i)\tGonyaulax',
        '(b)\tCellulosic plates\t(ii)\tDiatoms',
        '(c)\tChitinous wall\t(iii)\tMarine amoeba',
        '(d)\tSilica shells\t(iv)\tTrichomonas',
        '(A)\tii iv i iii\t(B)\tiv iii ii i',
        '(C)\tii i iv iii\t(D)\tiii i iv ii',
      ],
      answer: 'C',
    });

    expect(contents(parsed)).toEqual(['ii iv i iii', 'iv iii ii i', 'ii i iv iii', 'iii i iv ii']);
  });

  it('still reads a paper whose options are lower case only', async () => {
    const { parsed } = await readFirstQuestion({
      stem: 'A question',
      optionParagraphs: ['(a)\tone\t(b)\ttwo', '(c)\tthree\t(d)\tfour'],
      answer: 'B',
    });

    expect(contents(parsed)).toEqual(['one', 'two', 'three', 'four']);
  });

  it('leaves the item list untouched when the options are shuffled', async () => {
    const { block, parsed } = await readFirstQuestion({
      stem: 'Match the following',
      optionParagraphs: ['(a)\tSilica cell wall', '(A)\tii iv\t(B)\tiv iii', '(C)\tii i\t(D)\tiii i'],
      answer: 'C',
    });
    if (!parsed.ok) throw new Error('should parse');

    parsed.options.apply([3, 2, 1, 0]);
    const text = visibleOptionText(block);

    expect(text).toContain('(a)›Silica cell wall');
    expect(text).toContain('(A)›iii i');
    expect(text).toContain('(D)›ii iv');
  });
});

describe('a first option lettered by Word', () => {
  const HYBRID: FixtureQuestion = {
    stem: 'Read the following and select the correct option.',
    optionParagraphs: [
      'Both statements are correct and 2 explains 1',
      '(B)\tBoth statements are correct but 2 does not explain 1',
      '(C)\tStatement 1 is correct and statement 2 is incorrect',
      '(D)\tBoth statements are incorrect',
    ],
    answer: 'C',
    letteredFirstOption: true,
  };

  it('is recognised when only (B), (C) and (D) are typed', async () => {
    const { parsed } = await readFirstQuestion(HYBRID);
    expect(contents(parsed)).toEqual([
      'both statements are correct and 2 explains 1',
      'both statements are correct but 2 does not explain 1',
      'statement 1 is correct and statement 2 is incorrect',
      'both statements are incorrect',
    ]);
  });

  it('moves content into and out of the Word-lettered paragraph, keeping its numbering', async () => {
    const { block, parsed } = await readFirstQuestion(HYBRID);
    if (!parsed.ok) throw new Error('should parse');

    parsed.options.apply([2, 1, 0, 3]);
    const paragraphs = block.nodes.filter(
      (node) => node !== block.questionParagraph && node.namespaceURI === NS.w && node.localName === 'p',
    );

    // The lettered paragraph now holds what option (C) had, and is still a list item.
    expect(visibleText(paragraphs[0]!).trim()).toBe('Statement 1 is correct and statement 2 is incorrect');
    const pPr = firstChild(paragraphs[0]!, NS.w, 'pPr');
    expect(pPr && firstChild(pPr, NS.w, 'numPr')).toBeTruthy();
    // ...and the typed "(C)" label is still typed, now in front of the first option's text.
    expect(visibleText(paragraphs[2]!).replace(/\t/g, '›').trim()).toBe(
      '(C)›Both statements are correct and 2 explains 1',
    );
  });

  it('handles the whole question living in one lettered paragraph', async () => {
    // "(A)" comes from the list; "One" is option A; (B), (C), (D) are typed in the same run.
    const { block, parsed } = await readFirstQuestion({
      stem: 'Which of the following statements are true?',
      optionParagraphs: ['One\t\t(B)\tFour\t\t(C)\tTwo\t\t(D)\tThree'],
      answer: 'D',
      letteredFirstOption: true,
    });

    expect(contents(parsed)).toEqual(['one', 'four', 'two', 'three']);
    if (!parsed.ok) return;

    parsed.options.apply([3, 0, 1, 2]);
    expect(visibleOptionText(block).trim()).toBe('Three››(B)›One››(C)›Four››(D)›Two');
  });

  it('is refused when two lettered lists could both supply option (A)', async () => {
    const { parsed } = await readFirstQuestion({
      stem: 'Ambiguous question',
      optionParagraphs: [
        'first candidate for A',
        '(B)\tsecond',
        '(C)\tthird',
        '(D)\tfourth',
      ],
      answer: 'A',
      letteredFirstOption: true,
      // A second single-item "(%1)" list before (B) makes the choice ambiguous.
      secondLetteredList: ['another candidate for A'],
    });

    expect(parsed.ok).toBe(false);
    if (parsed.ok) return;
    expect(parsed.detail).toContain('BCD');
  });

  it('is refused when nothing supplies option (A)', async () => {
    const { parsed } = await readFirstQuestion({
      stem: 'Missing option A',
      optionParagraphs: ['(B)\tsecond', '(C)\tthird', '(D)\tfourth'],
      answer: 'B',
    });

    expect(parsed.ok).toBe(false);
    if (parsed.ok) return;
    expect(parsed.reason).toBe('unexpected-label-sequence');
    expect(parsed.detail).toContain('"BCD"');
  });
});

describe('failure messages', () => {
  it('blame the floating picture when one sits in the option area', async () => {
    const { parsed } = await readFirstQuestion({
      stem: 'Which diagram is right?',
      optionParagraphs: ['(A)\tfirst', '(B)\tsecond', '(D)\tfourth'],
      answer: 'A',
      rawOptionParagraph:
        '<w:p><w:pPr/><w:r><w:drawing><wp:anchor distT="0"><wp:extent cx="1" cy="1"/></wp:anchor></w:drawing></w:r>' +
        '<w:r><w:t xml:space="preserve">(C) </w:t></w:r><w:r><w:tab/><w:t>third</w:t></w:r></w:p>',
    });

    expect(parsed.ok).toBe(false);
    if (parsed.ok) return;
    expect(parsed.reason).toBe('option-contains-floating-graphic');
    expect(parsed.detail).toContain('floating picture');
  });
});
