/**
 * Shuffling must not change a single character of a question - it only re-orders whole
 * option contents. These tests pin the two rules that make that true:
 *
 *  1. Whitespace at the edges of an option belongs to the *slot*, not to the content, so
 *     the space that separates one option from the next label never travels away.
 *  2. A label's opening bracket typed with Insert > Symbol is read from the font's own
 *     encoding, so it belongs to the label. Only a bracket from a picture font, which says
 *     nothing about what it prints, leaves the previous option's end unknown - and that
 *     question is refused rather than guessed at.
 */
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { DocxPackage } from '../src/core/docx/DocxPackage';
import type { Element } from '../src/core/docx/dom';
import { NS, childElements, visibleText } from '../src/core/docx/xml';
import { GenerationService } from '../src/core/generate/GenerationService';
import { OptionSetParser } from '../src/core/options/OptionSetParser';
import { NumberingIndex } from '../src/core/parse/NumberingIndex';
import { PaperParser } from '../src/core/parse/PaperParser';
import type { QuestionBlock } from '../src/core/parse/PaperModel';
import type { GenerationRequest } from '../src/shared/types';
import {
  buildPaper,
  defaultSections,
  symbolBracketOptionParagraph,
  symbolContentOptionParagraph,
  type FixtureQuestion,
  type FixtureSection,
} from './support/PaperFixture';

/** A paper whose first question is the one under test, padded so the key is detectable. */
function paperWith(question: FixtureQuestion): FixtureSection[] {
  const filler = Array.from({ length: 6 }, (_unused, index) => ({
    stem: `Filler question ${index + 1}`,
    optionParagraphs: ['(A)\tone\t(B)\ttwo\t(C)\tthree\t(D)\tfour'],
    answer: 'A' as const,
  }));
  return [{ subject: 'PHYSICS', startNumber: 1, numId: '1', questions: [question, ...filler] }];
}

async function firstQuestion(sections: readonly FixtureSection[]) {
  const pkg = await DocxPackage.fromBuffer(await buildPaper(sections));
  const numbering = new NumberingIndex(pkg.numberingPart());
  const paper = new PaperParser().parsePart(pkg.documentPart(), numbering);
  const block = paper.sections[0]!.blocks[0]!;
  return { block, parsed: new OptionSetParser(numbering).parse(block) };
}

/** Visible text of the option region, tabs shown as arrows. */
const optionText = (block: QuestionBlock): string =>
  block.nodes
    .filter((node) => node !== block.questionParagraph && node.namespaceURI === NS.w && node.localName === 'p')
    .map((node) => visibleText(node))
    .join(' | ')
    .replace(/\t/g, '→');

/**
 * Option text with each `w:sym` shown as `glyph`. A symbol has no text of its own in the
 * document, so reading the paragraph normally cannot show whether one stayed with its
 * number or was left behind in the old column.
 */
function optionTextWithSymbols(block: QuestionBlock, glyph = 'Ω'): string {
  const render = (node: Element): string => {
    if (node.namespaceURI === NS.w && node.localName === 'sym') return glyph;
    if (node.namespaceURI === NS.w && node.localName === 't') return node.textContent ?? '';
    if (node.namespaceURI === NS.w && node.localName === 'tab') return '→';
    return childElements(node).map(render).join('');
  };
  return block.nodes
    .filter((node) => node !== block.questionParagraph && node.namespaceURI === NS.w && node.localName === 'p')
    .map((node) => render(node).trim())
    .join(' | ');
}

/** Multiset of words - invariant when only whole option contents are re-ordered. */
const words = (block: QuestionBlock): string[] =>
  block.nodes
    .filter((node) => node !== block.questionParagraph)
    .flatMap((node) => visibleText(node).replace(/\s+/g, ' ').trim().toLowerCase().split(' '))
    .filter((word) => word !== '')
    .sort();

describe('the separator between an option and the next label', () => {
  it('stays put when the option contents move', async () => {
    // The tab before (D) is missing, so a space separates "(iii)" from "(D)". That space
    // is part of the same run as the option text - the case that used to glue them.
    const { block, parsed } = await firstQuestion(
      paperWith({
        stem: 'Which statements are correct?',
        optionParagraphs: ['(A)\tonly (i)\t(B)\tonly (iv)\t(C)\t(i), (ii) and (iii) (D)\tall of the above'],
        answer: 'A',
      }),
    );
    if (!parsed.ok) throw new Error(`expected the options to parse, got ${parsed.reason}`);

    const before = words(block);
    parsed.options.apply([3, 2, 1, 0]);

    // The lone space stays where the author put it - in front of (D) - whatever content
    // lands in slot (C). Before the fix it travelled with "(i), (ii) and (iii)" and left
    // "only (iv)(D)" glued together.
    expect(optionText(block)).toBe(
      '(A)→all of the above→(B)→(i), (ii) and (iii)→(C)→only (iv) (D)→only (i)',
    );
    // Nothing gained, lost or glued: "above(B)" would show up here as a merged word.
    expect(words(block)).toEqual(before);
  });

  it('keeps the option text itself free of the separator', async () => {
    const { parsed } = await firstQuestion(
      paperWith({
        stem: 'Pick one',
        optionParagraphs: ['(A)\tone \t(B)\ttwo \t(C)\tthree \t(D)\tfour'],
        answer: 'A',
      }),
    );
    if (!parsed.ok) throw new Error('expected the options to parse');

    // Signatures are the option *contents*; the padding spaces are not part of them.
    expect(parsed.options.signatures.map((s) => s.split('|')[0])).toEqual(['one', 'two', 'three', 'four']);
  });

  it('leaves a tidy paper byte-for-byte identical in its option text', async () => {
    const { block, parsed } = await firstQuestion(
      paperWith({
        stem: 'What is the SI unit of force?',
        optionParagraphs: ['(A)\tnewton\t(B)\tjoule\t(C)\twatt\t(D)\tpascal'],
        answer: 'A',
      }),
    );
    if (!parsed.ok) throw new Error('expected the options to parse');

    parsed.options.apply([1, 0, 3, 2]);
    expect(optionText(block)).toBe('(A)→joule→(B)→newton→(C)→pascal→(D)→watt');
  });
});

describe('a label whose bracket was typed in the Symbol font', () => {
  const paper = (): FixtureSection[] =>
    paperWith({
      stem: 'Which of these is correct?',
      optionParagraphs: [],
      rawOptionParagraph: symbolBracketOptionParagraph(['first', 'second', 'third', 'none of these']),
      answer: 'D',
    });

  it('is read: Symbol 0x28 is a left parenthesis, so the label is "(A)"', async () => {
    const { parsed } = await firstQuestion(paper());

    if (!parsed.ok) throw new Error(`expected the options to parse, got ${parsed.reason}`);
    // No signature mentions the `w:sym`: the bracket went to the label, not to an option.
    expect(parsed.options.signatures).toEqual(['first|', 'second|', 'third|', 'none of these|']);
  });

  it('leaves the bracket where it is when the contents move', async () => {
    const { block, parsed } = await firstQuestion(paper());
    if (!parsed.ok) throw new Error('expected the options to parse');

    parsed.options.apply([3, 2, 1, 0]);

    // The bracket belongs to the label, not to the option before it, so all four stay put.
    expect(optionTextWithSymbols(block, '(')).toBe(
      '(A) →none of these→(B) →third→(C) →second→(D) →first→',
    );
  });

  it('shuffles like any other question, and still reads back', async () => {
    const workingDir = await fs.mkdtemp(path.join(os.tmpdir(), 'shuffler-sym-'));
    try {
      const sourceFile = path.join(workingDir, 'Paper.docx');
      await fs.writeFile(sourceFile, await buildPaper(paper()));
      const request: GenerationRequest = {
        sourceFile,
        shuffleQuestions: true,
        shuffleOptions: true,
        questionExclusions: [],
        optionExclusions: [],
        setCount: 2,
        seed: 'symbol',
      };
      const result = await new GenerationService().generate(request);

      expect(result.sets[0]!.verification.ok).toBe(true);
      expect(result.sets[0]!.optionsShuffled).toBe(7);
    } finally {
      await fs.rm(workingDir, { recursive: true, force: true });
    }
  });
});

describe('a label whose bracket comes from a picture font', () => {
  const paper = (): FixtureSection[] =>
    paperWith({
      stem: 'Which of these is correct?',
      optionParagraphs: [],
      // Wingdings numbers drawings, not characters: 0x28 is a glyph, and no reading of the
      // file can say it was meant as a bracket.
      rawOptionParagraph: symbolBracketOptionParagraph(
        ['first', 'second', 'third', 'none of these'],
        'Wingdings',
      ),
      answer: 'D',
    });

  it('is refused rather than shuffled with a guessed boundary', async () => {
    const { parsed } = await firstQuestion(paper());

    expect(parsed.ok).toBe(false);
    if (parsed.ok) return;
    expect(parsed.reason).toBe('label-bracket-is-a-symbol');
    expect(parsed.detail).toContain('Wingdings');
  });

  it('is listed in the dry run so the author can correct the document', async () => {
    const workingDir = await fs.mkdtemp(path.join(os.tmpdir(), 'shuffler-sym-'));
    try {
      const sourceFile = path.join(workingDir, 'Paper.docx');
      await fs.writeFile(sourceFile, await buildPaper(paper()));
      const report = await new GenerationService().dryRun({
        sourceFile,
        shuffleQuestions: true,
        shuffleOptions: true,
        questionExclusions: [],
        optionExclusions: [],
        setCount: 2,
      });

      const kept = report.optionsKeptByTool.find((item) => item.questionNumber === 1);
      expect(kept?.reason).toBe('label-bracket-is-a-symbol');
    } finally {
      await fs.rm(workingDir, { recursive: true, force: true });
    }
  });

  it('does not stop the rest of the paper being shuffled', async () => {
    const workingDir = await fs.mkdtemp(path.join(os.tmpdir(), 'shuffler-sym-'));
    try {
      const sourceFile = path.join(workingDir, 'Paper.docx');
      await fs.writeFile(sourceFile, await buildPaper(paper()));
      const request: GenerationRequest = {
        sourceFile,
        shuffleQuestions: true,
        shuffleOptions: true,
        questionExclusions: [],
        optionExclusions: [],
        setCount: 2,
        seed: 'symbol',
      };
      const result = await new GenerationService().generate(request);

      expect(result.sets[0]!.verification.ok).toBe(true);
      expect(result.sets[0]!.optionsShuffled).toBe(6);
    } finally {
      await fs.rm(workingDir, { recursive: true, force: true });
    }
  });
});

describe('an option whose content is a symbol', () => {
  const paper = (): FixtureSection[] =>
    paperWith({
      stem: 'What is the resistance?',
      optionParagraphs: [],
      // "60Ω  (B) 50Ω  (C) 80Ω  (D) 100Ω" - every symbol sits right before the next label.
      rawOptionParagraph: symbolContentOptionParagraph(['60', '50', '80', '100']),
      answer: 'A',
    });

  it('is shuffled: a bracketed label has its bracket, so the symbol is content', async () => {
    const { parsed } = await firstQuestion(paper());

    if (!parsed.ok) throw new Error(`expected the options to parse, got ${parsed.reason}`);
    expect(parsed.options.signatures).toEqual([
      '60ω|sym:Symbol:F057',
      '50ω|sym:Symbol:F057',
      '80ω|sym:Symbol:F057',
      '100ω|sym:Symbol:F057',
    ]);
  });

  it('carries its symbol with it when it moves', async () => {
    const { block, parsed } = await firstQuestion(paper());
    if (!parsed.ok) throw new Error('expected the options to parse');

    parsed.options.apply([3, 2, 1, 0]);

    // Every unit stays welded to its number: "100Ω" must not become "100" next to a
    // stray Ω left behind in the old column.
    expect(optionTextWithSymbols(block)).toBe('(A) →100Ω→(B) →80Ω→(C) →50Ω→(D) →60Ω→');
  });

  it('survives being written out and read back', async () => {
    const workingDir = await fs.mkdtemp(path.join(os.tmpdir(), 'shuffler-symcontent-'));
    try {
      const sourceFile = path.join(workingDir, 'Paper.docx');
      await fs.writeFile(sourceFile, await buildPaper(paper()));
      const result = await new GenerationService().generate({
        sourceFile,
        shuffleQuestions: true,
        shuffleOptions: true,
        questionExclusions: [],
        optionExclusions: [],
        setCount: 2,
        seed: 'symbol-content',
      });

      // Verification re-parses the generated file: a question the tool shuffles must still
      // read as four options afterwards, or the shuffle was based on a boundary that the
      // shuffle itself destroyed.
      expect(result.sets[0]!.verification.ok).toBe(true);
      expect(result.sets[0]!.optionsShuffled).toBe(7);
    } finally {
      await fs.rm(workingDir, { recursive: true, force: true });
    }
  });
});

describe('whole-paper fidelity', () => {
  let workingDir = '';
  let sourceFile = '';

  beforeEach(async () => {
    workingDir = await fs.mkdtemp(path.join(os.tmpdir(), 'shuffler-fidelity-'));
    sourceFile = path.join(workingDir, 'Paper.docx');
    await fs.writeFile(sourceFile, await buildPaper(defaultSections()));
  });

  afterEach(async () => {
    await fs.rm(workingDir, { recursive: true, force: true });
  });

  it('changes no word of any question, only their order', async () => {
    const result = await new GenerationService().generate({
      sourceFile,
      shuffleQuestions: true,
      shuffleOptions: true,
      questionExclusions: [],
      optionExclusions: [],
      setCount: 2,
      seed: 'fidelity',
    });

    const read = async (file: string): Promise<Map<number, string[]>> => {
      const pkg = await DocxPackage.fromBuffer(await fs.readFile(file));
      const parsed = new PaperParser().parsePart(pkg.documentPart(), new NumberingIndex(pkg.numberingPart()));
      const out = new Map<number, string[]>();
      for (const section of parsed.sections) {
        for (const block of section.blocks) out.set(block.printedNumber, words(block));
      }
      return out;
    };

    const original = await read(sourceFile);
    for (const set of result.sets) {
      const generated = await read(set.filePath);
      for (const mapping of set.mappings) {
        expect(generated.get(mapping.newNumber)).toEqual(original.get(mapping.originalNumber));
      }
    }
  });
});
