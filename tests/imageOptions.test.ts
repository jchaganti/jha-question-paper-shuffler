/**
 * Options that are pictures rather than words.
 *
 * An option's answer is not always text - a chemistry paper's four options are often four
 * structures. Whether such an option can be shuffled turns on one thing only, and it is a
 * fact stated in the document rather than a judgement:
 *
 *  - **Inline** ("In line with text"): the picture sits in the text flow and moves with the
 *    run that holds it, so the option shuffles like any other.
 *  - **Floating** (anchored, "Square"/"Tight"/...): the picture is placed from the page.
 *    Which option each one belongs to cannot be read from the file - the pictures and the
 *    labels sit in different paragraphs, in an order that differs from question to
 *    question - so the question is refused with a message naming the fix.
 */
import { describe, expect, it } from 'vitest';
import { DocxPackage } from '../src/core/docx/DocxPackage';
import { NS, descendants, visibleText } from '../src/core/docx/xml';
import { OptionSetParser } from '../src/core/options/OptionSetParser';
import { NumberingIndex } from '../src/core/parse/NumberingIndex';
import { PaperParser } from '../src/core/parse/PaperParser';
import type { QuestionBlock } from '../src/core/parse/PaperModel';
import { buildPaper, type FixtureQuestion, type FixtureSection } from './support/PaperFixture';

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

/** The picture ids in each option paragraph, in document order. */
function picturesPerParagraph(block: QuestionBlock): string[] {
  const out: string[] = [];
  for (const node of block.nodes) {
    if (node === block.questionParagraph || node.localName !== 'p') continue;
    for (const docPr of descendants(node, NS.wp, 'docPr')) {
      out.push((docPr.getAttribute('name') ?? '').replace('Picture ', ''));
    }
  }
  return out;
}

describe('options whose answers are inline pictures', () => {
  const question: FixtureQuestion = {
    stem: 'Which structure is correct?',
    optionParagraphs: ['(A)\t[[image:rIdA]]', '(B)\t[[image:rIdB]]', '(C)\t[[image:rIdC]]', '(D)\t[[image:rIdD]]'],
    answer: 'A',
  };

  it('are read as content, not as an empty option', async () => {
    const { parsed } = await readFirstQuestion(question);

    expect(parsed.ok).toBe(true);
  });

  it('move between slots when the options are shuffled', async () => {
    const { block, parsed } = await readFirstQuestion(question);
    if (!parsed.ok) throw new Error('expected the options to parse');
    expect(picturesPerParagraph(block)).toEqual(['rIdA', 'rIdB', 'rIdC', 'rIdD']);

    // permutation[newSlot] = originalSlot: new A shows old D, new B shows old C, ...
    parsed.options.apply([3, 2, 1, 0]);

    expect(picturesPerParagraph(block)).toEqual(['rIdD', 'rIdC', 'rIdB', 'rIdA']);
  });

  it('keeps each label in place while its picture changes', async () => {
    const { block, parsed } = await readFirstQuestion(question);
    if (!parsed.ok) throw new Error('expected the options to parse');

    parsed.options.apply([3, 2, 1, 0]);

    const labels = block.nodes
      .filter((node) => node !== block.questionParagraph && node.localName === 'p')
      .map((node) => visibleText(node).trim());
    expect(labels).toEqual(['(A)', '(B)', '(C)', '(D)']);
  });

  it('tells them apart, so a picture is never duplicated or lost', async () => {
    const { block, parsed } = await readFirstQuestion(question);
    if (!parsed.ok) throw new Error('expected the options to parse');

    parsed.options.apply([1, 0, 3, 2]);

    expect([...picturesPerParagraph(block)].sort()).toEqual(['rIdA', 'rIdB', 'rIdC', 'rIdD']);
  });
});

describe('a text option mixed with an inline picture', () => {
  it('moves the words and the picture together', async () => {
    const { block, parsed } = await readFirstQuestion({
      stem: 'Which structure is correct?',
      optionParagraphs: ['(A)\tthe ring [[image:rIdA]]', '(B)\ttwo', '(C)\tthree', '(D)\tfour'],
      answer: 'A',
    });
    if (!parsed.ok) throw new Error('expected the options to parse');

    parsed.options.apply([1, 2, 3, 0]);

    const paragraphs = block.nodes
      .filter((node) => node !== block.questionParagraph && node.localName === 'p')
      .map((node) => visibleText(node).replace(/\s+/g, ' ').trim());
    // Option (A)'s words moved to slot (D); its picture went with them.
    expect(paragraphs[3]).toBe('(D) the ring');
    expect(picturesPerParagraph(block)).toEqual(['rIdA']);
  });
});

describe('options whose answers are floating pictures', () => {
  it('are refused with the change that makes them work', async () => {
    const { parsed } = await readFirstQuestion({
      stem: 'Which structure is correct?',
      optionParagraphs: [
        '[[float:structure A]](A)\t',
        '[[float:structure B]](B)\t',
        '[[float:structure C]](C)\t',
        '[[float:structure D]](D)\t',
      ],
      answer: 'A',
    });

    expect(parsed.ok).toBe(false);
    if (parsed.ok) return;
    expect(parsed.reason).toBe('option-contains-floating-graphic');
    expect(parsed.detail).toMatch(/options are floating pictures/);
    expect(parsed.detail).toMatch(/In line with text/);
  });
});
