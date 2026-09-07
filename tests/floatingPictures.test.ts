/**
 * Floating pictures anchored among the options.
 *
 * A floating picture is placed from the *page*, not from the text beside it. Two things
 * follow, and both are what these tests pin:
 *
 *  - Its text box is not part of the sentence the reader sees, so it must not be counted
 *    when looking for option labels. Counting it put a diagram's stray letters in front of
 *    the next label ("O A B Cl" before "(2)") and hid it, which is why whole questions
 *    used to be reported as "labels read as 134".
 *  - It never moves. So it stays exactly where it is while the answers move around it,
 *    and a picture wedged *between the words of one answer* is refused instead, because
 *    there the words would leave without it.
 */
import { describe, expect, it } from 'vitest';
import { DocxPackage } from '../src/core/docx/DocxPackage';
import { NS, descendants, visibleText } from '../src/core/docx/xml';
import { OptionSetParser } from '../src/core/options/OptionSetParser';
import { NumberingIndex } from '../src/core/parse/NumberingIndex';
import { PaperParser } from '../src/core/parse/PaperParser';
import type { QuestionBlock } from '../src/core/parse/PaperModel';
import {
  buildPaper,
  floatingPictureParagraph,
  type FixtureQuestion,
  type FixtureSection,
} from './support/PaperFixture';

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

function contents(parsed: Awaited<ReturnType<typeof readFirstQuestion>>['parsed']): string[] {
  if (!parsed.ok) throw new Error(`expected the options to parse, got ${parsed.reason}: ${parsed.detail}`);
  return parsed.options.signatures.map((signature) => signature.split('|')[0]!);
}

/** Which option paragraph (by index) anchors each floating picture, and its caption. */
function anchorsPerParagraph(block: QuestionBlock): string[] {
  const out: string[] = [];
  block.nodes.forEach((node, index) => {
    if (node.namespaceURI !== NS.w || node.localName !== 'p') return;
    for (const drawing of descendants(node, NS.w, 'drawing')) {
      if (descendants(drawing, NS.wp, 'anchor').length > 0) out.push(`${index}:${visibleText(drawing).trim()}`);
    }
  });
  return out.sort();
}

describe('a diagram anchored in front of an option label', () => {
  // The exact shape of FST-1 Q12: "(1) 3:5:7" then a paragraph whose picture caption runs
  // straight into "(2) 3:4:5".
  const question: FixtureQuestion = {
    stem: 'Three identical particles are joined together',
    optionParagraphs: [
      '(A)\t3 : 5 : 7',
      '[[float:O    A    B    Cl]](B)\t3 : 4 : 5',
      '(C)\t7 : 11 : 6',
      '(D)\t3 : 5 : 6',
    ],
    answer: 'D',
  };

  it('does not hide the label behind it', async () => {
    const { parsed } = await readFirstQuestion(question);

    expect(contents(parsed)).toEqual(['3 : 5 : 7', '3 : 4 : 5', '7 : 11 : 6', '3 : 5 : 6']);
  });

  it('leaves the picture in its own paragraph when the answers move', async () => {
    const { block, parsed } = await readFirstQuestion(question);
    const before = anchorsPerParagraph(block);

    if (!parsed.ok) throw new Error('expected the options to parse');
    parsed.options.apply([3, 2, 1, 0]);

    expect(anchorsPerParagraph(block)).toEqual(before);
    expect(before).toHaveLength(1);
  });

  it('moves the answers around it', async () => {
    const { block, parsed } = await readFirstQuestion(question);
    if (!parsed.ok) throw new Error('expected the options to parse');

    parsed.options.apply([3, 2, 1, 0]);

    const text = block.nodes
      .filter((node) => node !== block.questionParagraph)
      .map((node) => visibleText(node))
      .join(' | ');
    expect(text.replace(/[\t ]+/g, ' ')).toContain('(A) 3 : 5 : 6');
    expect(text.replace(/[\t ]+/g, ' ')).toContain('(D) 3 : 5 : 7');
    // The caption is still in the paragraph that carries option (B).
    expect(text.replace(/[\t ]+/g, ' ')).toContain('O A B Cl(B) 7 : 11 : 6');
  });

  it('is reported so the author can check it', async () => {
    const { parsed } = await readFirstQuestion(question);
    if (!parsed.ok) throw new Error('expected the options to parse');

    const note = parsed.notes.find((entry) => entry.issue === 'floating-picture-in-option-area');
    expect(note).toBeDefined();
    expect(note!.fix).toMatch(/In Line with Text/);
  });
});

describe('a diagram in a paragraph after the last option', () => {
  // FST-1 Q70/Q74: the next question's artwork is anchored at the end of this block, which
  // used to read as "option (D) continues on another paragraph".
  const question: FixtureQuestion = {
    stem: 'Which of the following is correct?',
    optionParagraphs: ['(A)\tone', '(B)\ttwo', '(C)\tthree', '(D)\tfour'],
    rawOptionParagraph: floatingPictureParagraph('CH3 HA HE2 Alcoholic KOH'),
    answer: 'A',
  };

  it('does not make the last option span paragraphs', async () => {
    const { parsed } = await readFirstQuestion(question);

    expect(contents(parsed)).toEqual(['one', 'two', 'three', 'four']);
  });

  it('leaves the picture in its own paragraph when the answers move', async () => {
    const { block, parsed } = await readFirstQuestion(question);
    const before = anchorsPerParagraph(block);
    if (!parsed.ok) throw new Error('expected the options to parse');

    parsed.options.apply([1, 0, 3, 2]);

    expect(anchorsPerParagraph(block)).toEqual(before);
  });
});

describe('material trailing the option list', () => {
  // FST-1 Q35: the last option is followed by two spacer paragraphs and a "SECTION B"
  // instruction, which the last option swallowed because nothing bounds it.
  it('does not become part of the last option', async () => {
    const { parsed } = await readFirstQuestion({
      stem: 'The current through the circuit is',
      optionParagraphs: ['(A)\t3A', '(B)\t1A', '(C)\t4A', '(D)\t0', '', '', 'SECTION B (Attempt any 10 questions)'],
      answer: 'A',
    });

    expect(contents(parsed)).toEqual(['3a', '1a', '4a', '0']);
  });

  it('stays exactly where it is when the options move', async () => {
    const { block, parsed } = await readFirstQuestion({
      stem: 'The current through the circuit is',
      optionParagraphs: ['(A)\t3A', '(B)\t1A', '(C)\t4A', '(D)\t0', '', '', 'SECTION B (Attempt any 10 questions)'],
      answer: 'A',
    });
    if (!parsed.ok) throw new Error('expected the options to parse');

    parsed.options.apply([3, 2, 1, 0]);

    const paragraphs = block.nodes
      .filter((node) => node !== block.questionParagraph && node.localName === 'p')
      .map((node) => visibleText(node).replace(/\s+/g, ' ').trim());
    // The spacers and the "SECTION B" instruction are no longer part of this question at
    // all: the instruction opens the next section, so they belong to that section's header
    // and cannot travel with a question even in principle. See `sectionDividers.test.ts`.
    expect(paragraphs).toEqual(['(A) 0', '(B) 4A', '(C) 1A', '(D) 3A']);
  });

  it('still refuses an option that genuinely runs on to the next paragraph', async () => {
    // No blank paragraph in between, so this is a continuation, not trailing material.
    const { parsed } = await readFirstQuestion({
      stem: 'Pick the correct statement',
      optionParagraphs: ['(A)\tone', '(B)\ttwo', '(C)\tthree', '(D)\tfour', 'and the rest of option D'],
      answer: 'A',
    });

    expect(parsed.ok).toBe(false);
    if (parsed.ok) return;
    expect(parsed.reason).toBe('option-spans-paragraphs');
  });
});

describe('a picture wedged between the words of one answer', () => {
  it('is refused, because the words would move and the picture would not', async () => {
    const { parsed } = await readFirstQuestion({
      stem: 'Identify the diagram',
      optionParagraphs: [
        '(A)\tthe shape [[float:diagram]] shown above',
        '(B)\ttwo',
        '(C)\tthree',
        '(D)\tfour',
      ],
      answer: 'A',
    });

    expect(parsed.ok).toBe(false);
    if (parsed.ok) return;
    expect(parsed.reason).toBe('option-contains-floating-graphic');
    expect(parsed.detail).toMatch(/in the middle of its words/);
    expect(parsed.fix).toMatch(/In Line with Text/);
  });
});
