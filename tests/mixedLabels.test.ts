/**
 * Papers letter some options with Word's numbering and type the rest by hand. Every mixture
 * of the two is one paper feature, so the reader treats it as one: typed labels and lettered
 * paragraphs are laid out in document order and have to make exactly four options.
 *
 * The second half of this file pins the boundary of that reading - the shapes where the
 * merge does *not* add up, which are refused rather than guessed at.
 */
import { describe, expect, it } from 'vitest';
import { DocxPackage } from '../src/core/docx/DocxPackage';
import { NS, descendants, visibleText } from '../src/core/docx/xml';
import { NumberingIndex } from '../src/core/parse/NumberingIndex';
import { PaperParser } from '../src/core/parse/PaperParser';
import { OptionSetParser } from '../src/core/options/OptionSetParser';
import type { QuestionBlock } from '../src/core/parse/PaperModel';
import { LIST, buildPaper, type FixtureQuestion, type FixtureSection } from './support/PaperFixture';

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

type Parsed = Awaited<ReturnType<typeof readFirstQuestion>>['parsed'];

/** The four option contents, in slot order A-D. */
function contents(parsed: Parsed): string[] {
  if (!parsed.ok) throw new Error(`expected the options to parse, got ${parsed.reason}: ${parsed.detail}`);
  return parsed.options.signatures.map((signature) => signature.split('|')[0]!);
}

const optionText = (block: QuestionBlock): string =>
  block.nodes
    .filter((node) => node !== block.questionParagraph && node.namespaceURI === NS.w && node.localName === 'p')
    .map((node) => visibleText(node).replace(/\s+/g, ' ').trim())
    .filter((text) => text !== '')
    .join(' | ');

describe('options part lettered by Word and part typed', () => {
  it('reads three lettered paragraphs and one typed label', async () => {
    // The shape four questions of MTP-2-XI-2023 use: Word letters (A) (B) (C), the author
    // typed the (D).
    const { parsed } = await readFirstQuestion({
      stem: 'Select the correct statement.',
      optionParagraphs: ['\tfungi', '\teuglena', '\tgonyaulax', '(D)\tslime moulds'],
      autoLetteredFirst: 3,
      answer: 'B',
    });

    expect(contents(parsed)).toEqual(['fungi', 'euglena', 'gonyaulax', 'slime moulds']);
  });

  it('reads two lettered and two typed', async () => {
    const { parsed } = await readFirstQuestion({
      stem: 'Assertion and Reason.',
      optionParagraphs: ['\tboth true, R explains A', '\tboth true, R does not explain A', '(C)\tA true, R false', '(D)\tA false, R true'],
      autoLetteredFirst: 2,
      answer: 'A',
    });

    expect(contents(parsed)).toEqual([
      'both true, r explains a',
      'both true, r does not explain a',
      'a true, r false',
      'a false, r true',
    ]);
  });

  it('reads a lettered option that shares its line with a typed one', async () => {
    // "(A) 1 amp <tab> (B) 1.5 amp" is one paragraph: Word draws the (A) at the start of the
    // line, so the lettered option comes before the typed label inside the same paragraph.
    const { parsed } = await readFirstQuestion({
      stem: 'The current in the circuit is',
      optionParagraphs: ['1 amp\t(B)\t1.5 amp', '(C)\t2 amp\t(D)\t2.5 amp'],
      autoLetteredFirst: 1,
      answer: 'A',
    });

    expect(contents(parsed)).toEqual(['1 amp', '1.5 amp', '2 amp', '2.5 amp']);
  });

  it('shuffles them, leaving every label where it is', async () => {
    const { block, parsed } = await readFirstQuestion({
      stem: 'Select the correct statement.',
      optionParagraphs: ['\tfungi', '\teuglena', '\tgonyaulax', '(D)\tslime moulds'],
      autoLetteredFirst: 3,
      answer: 'B',
    });
    if (!parsed.ok) throw new Error('expected the options to parse');

    parsed.options.apply([3, 2, 1, 0]);

    // The typed "(D)" has not moved; only the answers have.
    expect(optionText(block)).toBe('slime moulds | gonyaulax | euglena | (D) fungi');
  });

  it('reports the mixture, so the author can make the paper consistent', async () => {
    const { parsed } = await readFirstQuestion({
      stem: 'Select the correct statement.',
      optionParagraphs: ['\tfungi', '\teuglena', '\tgonyaulax', '(D)\tslime moulds'],
      autoLetteredFirst: 3,
      answer: 'B',
    });
    if (!parsed.ok) throw new Error('expected the options to parse');

    const note = parsed.notes.find((item) => item.issue === 'mixed-auto-and-typed-labels');
    expect(note?.detail).toBe(
      'The first, second, third options are lettered by Word, but (D) is typed into the text.',
    );
  });

  it('prefers the lettered list whose case matches the typed labels', async () => {
    // Two lettered lists: an "(a) (b)" statement list, and the "(A)" of the options. The
    // typed labels are upper case, so the upper-case list is the options.
    const { parsed } = await readFirstQuestion({
      stem: 'Consider two situations',
      optionParagraphs: ['\tVa > Vb', '(B)\tIa > Ib', '(C)\tVa = Vb', '(D)\tVa < Vb'],
      autoLetteredFirst: 1,
      secondLetteredList: ['air filled', 'mica filled'],
      secondLetteredListId: LIST.lowerStatements,
      answer: 'A',
    });

    expect(contents(parsed)).toEqual(['va > vb', 'ia > ib', 'va = vb', 'va < vb']);
  });

  it('refuses when the lettered paragraphs do not account for the missing options', async () => {
    // Three lettered statements and three typed labels: seven entries, not four.
    const { parsed } = await readFirstQuestion({
      stem: 'Which of the following is correct?',
      optionParagraphs: ['(B)\ttwo', '(C)\tthree', '(D)\tfour'],
      secondLetteredList: ['statement one', 'statement two', 'statement three'],
      answer: 'B',
    });

    expect(parsed.ok).toBe(false);
    if (parsed.ok) return;
    expect(parsed.reason).toBe('unexpected-label-sequence');
  });

  it('refuses when the typed labels are out of order', async () => {
    const { parsed } = await readFirstQuestion({
      stem: 'Which of the following is correct?',
      optionParagraphs: ['\tone', '(D)\tfour', '(B)\ttwo', '(C)\tthree'],
      autoLetteredFirst: 1,
      answer: 'A',
    });

    expect(parsed.ok).toBe(false);
  });
});

describe('an option split across two paragraphs', () => {
  const spilled = (): FixtureQuestion => ({
    stem: 'Which statements are correct?',
    // Enter was pressed inside option (B), so its text runs on to a paragraph of its own.
    optionParagraphs: ['(A)\tboth true, B explains A', '(B)\tboth true, but B is not the', 'correct explanation', '(C)\tA true, B false', '(D)\tboth false'],
    answer: 'A',
  });

  it('is refused, and the message says which key to press', async () => {
    const { parsed } = await readFirstQuestion(spilled());

    expect(parsed.ok).toBe(false);
    if (parsed.ok) return;
    expect(parsed.reason).toBe('option-spans-paragraphs');
    expect(parsed.detail).toContain('Option (B)');
    expect(parsed.fix).toContain('Shift+Enter');
  });

  it('shuffles once the break is a line break instead', async () => {
    // Shift+Enter prints the same two lines but keeps the option in one paragraph, so its
    // text has a line to move to. This is the fix the message names.
    const { block, parsed } = await readFirstQuestion({
      stem: 'Which statements are correct?',
      optionParagraphs: [
        '(A)\tboth true, B explains A',
        '(B)\tboth true, but B is not the[[br]]correct explanation',
        '(C)\tA true, B false',
        '(D)\tboth false',
      ],
      answer: 'A',
    });
    if (!parsed.ok) throw new Error(`expected the options to parse, got ${parsed.reason}`);

    // The signature collapses the break to a space; both lines are one option's content.
    expect(contents(parsed)).toEqual([
      'both true, b explains a',
      'both true, but b is not the correct explanation',
      'a true, b false',
      'both false',
    ]);

    parsed.options.apply([1, 0, 2, 3]);

    // Both lines of option B travelled together into slot A - and the break came with them,
    // so slot A still prints on two lines.
    expect(optionText(block)).toBe(
      '(A) both true, but B is not the correct explanation | (B) both true, B explains A | (C) A true, B false | (D) both false',
    );
    const slotA = block.nodes.filter((node) => node.namespaceURI === NS.w && node.localName === 'p')[1]!;
    expect(descendants(slotA, NS.w, 'br')).toHaveLength(1);
  });
});
