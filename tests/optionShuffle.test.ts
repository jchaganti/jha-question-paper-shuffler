import { describe, expect, it } from 'vitest';
import { DocxPackage } from '../src/core/docx/DocxPackage';
import { visibleText } from '../src/core/docx/xml';
import { OptionBlockParser, slotSignature } from '../src/core/options/OptionBlockParser';
import { OptionShuffleApplier } from '../src/core/options/OptionShuffleApplier';
import { NumberingIndex } from '../src/core/parse/NumberingIndex';
import { PaperParser } from '../src/core/parse/PaperParser';
import type { QuestionBlock } from '../src/core/parse/PaperModel';
import { buildPaper, defaultSections, floatingPictureOptionParagraph } from './support/PaperFixture';

async function blocks(sections = defaultSections()): Promise<Map<number, QuestionBlock>> {
  const pkg = await DocxPackage.fromBuffer(await buildPaper(sections));
  const paper = new PaperParser().parsePart(pkg.documentPart(), new NumberingIndex(pkg.numberingPart()));
  const out = new Map<number, QuestionBlock>();
  for (const section of paper.sections) {
    for (const block of section.blocks) out.set(block.printedNumber, block);
  }
  return out;
}

const optionText = (block: QuestionBlock): string =>
  block.nodes
    .slice(1)
    .map((node) => visibleText(node))
    .join(' | ')
    .replace(/\t/g, '›');

describe('OptionBlockParser', () => {
  it('reads four options typed on one line', async () => {
    const block = (await blocks()).get(1)!;
    const parsed = new OptionBlockParser().parse(block);

    expect(parsed.ok).toBe(true);
    if (!parsed.ok) return;
    expect(parsed.block.slots.map((slot) => slot.letter)).toEqual(['A', 'B', 'C', 'D']);
    expect(parsed.block.slots.map((slot) => slot.coreAtoms.map((atom) => atom.text).join('').trim())).toEqual([
      'newton',
      'joule',
      'watt',
      'pascal',
    ]);
  });

  it('reads options spread over two and four paragraphs', async () => {
    const all = await blocks();
    for (const questionNumber of [2, 3]) {
      const parsed = new OptionBlockParser().parse(all.get(questionNumber)!);
      expect(parsed.ok, `question ${questionNumber}`).toBe(true);
    }
  });

  it('refuses options that Word letters automatically', async () => {
    const parsed = new OptionBlockParser().parse((await blocks()).get(4)!);
    expect(parsed.ok).toBe(false);
    if (parsed.ok) return;
    expect(parsed.reason).toBe('options-not-found');
  });

  it('refuses an option whose whole answer is a floating picture, and says how to fix it', async () => {
    // An option may perfectly well *be* a picture - but only an inline one can be moved.
    // `floatingPictures.test.ts` covers the cases that *are* now shuffled, and
    // `imageOptions.test.ts` covers inline images, which shuffle like any other content.
    const sections = defaultSections();
    const physics = sections[0]!;
    const withPicture = {
      ...physics,
      questions: [
        {
          stem: 'Identify the diagram',
          optionParagraphs: ['(A)\tfirst', '(B)\tsecond', '(C)\tthird'],
          rawOptionParagraph: floatingPictureOptionParagraph('(D)').replace('<w:r><w:t>diagram</w:t></w:r>', ''),
          answer: 'A' as const,
        },
        ...physics.questions.slice(1),
      ],
    };
    const parsed = new OptionBlockParser().parse((await blocks([withPicture, sections[1]!])).get(1)!);

    expect(parsed.ok).toBe(false);
    if (parsed.ok) return;
    expect(parsed.reason).toBe('option-contains-floating-graphic');
    expect(parsed.detail).toMatch(/Option \(D\) has no text of its own/);
    expect(parsed.detail).toMatch(/options are floating pictures/);
    expect(parsed.detail).toMatch(/In line with text/);
  });

  it('does not treat "(A)" inside option text as a label', async () => {
    const sections = defaultSections();
    const section = {
      ...sections[0]!,
      questions: [
        {
          stem: 'Which reactant is limiting?',
          optionParagraphs: ['(A)\tCaCO3\t(B)\tHCl', '(C)\tBoth (A) and (B)\t(D)\tNone of these'],
          answer: 'C' as const,
        },
        ...sections[0]!.questions.slice(1),
      ],
    };
    const parsed = new OptionBlockParser().parse((await blocks([section, sections[1]!])).get(1)!);

    expect(parsed.ok).toBe(true);
    if (!parsed.ok) return;
    expect(parsed.block.slots[2]!.coreAtoms.map((atom) => atom.text).join('').trim()).toBe('Both (A) and (B)');
  });
});

describe('OptionShuffleApplier', () => {
  it('moves option text between labels while labels and tabs stay put', async () => {
    const block = (await blocks()).get(1)!;
    expect(optionText(block)).toBe('(A)›newton›(B)›joule›(C)›watt›(D)›pascal');

    const parser = new OptionBlockParser();
    const parsed = parser.parse(block);
    if (!parsed.ok) throw new Error('fixture should parse');
    new OptionShuffleApplier().apply(parsed.block, [2, 0, 3, 1]);

    expect(optionText(block)).toBe('(A)›watt›(B)›newton›(C)›pascal›(D)›joule');
  });

  it('keeps the paragraph layout when options span several paragraphs', async () => {
    const block = (await blocks()).get(2)!;
    const parsed = new OptionBlockParser().parse(block);
    if (!parsed.ok) throw new Error('fixture should parse');
    new OptionShuffleApplier().apply(parsed.block, [3, 2, 1, 0]);

    expect(optionText(block)).toBe('(A)›g/2›(B)›2g | (C)›g›(D)›zero');
  });

  it('moves embedded OLE equations with their option', async () => {
    const block = (await blocks()).get(7)!;
    const parser = new OptionBlockParser();
    const before = parser.parse(block);
    if (!before.ok) throw new Error('fixture should parse');
    const signatures = before.block.slots.map((slot) => slotSignature(slot.coreAtoms));

    new OptionShuffleApplier().apply(before.block, [1, 2, 3, 0]);

    const after = parser.parse(block);
    if (!after.ok) throw new Error('should still parse after shuffling');
    const shuffled = after.block.slots.map((slot) => slotSignature(slot.coreAtoms));

    // Slot A now holds what slot B had, and every equation is still present exactly once.
    expect(shuffled[0]).toBe(signatures[1]);
    expect([...shuffled].sort()).toEqual([...signatures].sort());
    expect(shuffled[0]).toContain('rel:id=rId101');
  });

  it('is reversible: applying the inverse permutation restores the original', async () => {
    const block = (await blocks()).get(1)!;
    const parser = new OptionBlockParser();
    const original = optionText(block);

    const first = parser.parse(block);
    if (!first.ok) throw new Error('fixture should parse');
    new OptionShuffleApplier().apply(first.block, [1, 2, 3, 0]);

    const second = parser.parse(block);
    if (!second.ok) throw new Error('should still parse');
    new OptionShuffleApplier().apply(second.block, [3, 0, 1, 2]);

    expect(optionText(block)).toBe(original);
  });
});
