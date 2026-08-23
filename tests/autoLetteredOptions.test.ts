/**
 * Options that Word letters automatically: each option is its own list paragraph and the
 * "(A)" is generated from the numbering definition, not typed. Shuffling such a question
 * means moving the content between the paragraphs and leaving the numbering in place.
 */
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { DocxPackage } from '../src/core/docx/DocxPackage';
import { NS, childElements, firstChild, visibleText } from '../src/core/docx/xml';
import { GenerationService } from '../src/core/generate/GenerationService';
import { OptionSetParser } from '../src/core/options/OptionSetParser';
import { NumberingIndex } from '../src/core/parse/NumberingIndex';
import { PaperParser } from '../src/core/parse/PaperParser';
import type { QuestionBlock } from '../src/core/parse/PaperModel';
import { OPTION_LETTERS, type GenerationRequest } from '../src/shared/types';
import { buildPaper, defaultSections, floatingPictureOptionParagraph, type FixtureQuestion, type FixtureSection } from './support/PaperFixture';

const AUTO_OPTIONS = ['first choice', 'second choice', 'third choice', 'fourth choice'];

/** A paper whose first question is auto-lettered, padded so the answer key is detectable. */
function autoLetteredPaper(question: Partial<FixtureQuestion> = {}): FixtureSection[] {
  const filler = Array.from({ length: 6 }, (_unused, index) => ({
    stem: `Filler question ${index + 1}`,
    optionParagraphs: ['(A)\tone\t(B)\ttwo\t(C)\tthree\t(D)\tfour'],
    answer: 'A' as const,
  }));
  return [
    {
      subject: 'PHYSICS',
      startNumber: 1,
      numId: '1',
      questions: [
        {
          stem: 'Which option is correct?',
          optionParagraphs: AUTO_OPTIONS,
          answer: 'C',
          autoLettered: true,
          ...question,
        },
        ...filler,
      ],
    },
  ];
}

async function open(sections: readonly FixtureSection[]) {
  const pkg = await DocxPackage.fromBuffer(await buildPaper(sections));
  const numbering = new NumberingIndex(pkg.numberingPart());
  const paper = new PaperParser().parsePart(pkg.documentPart(), numbering);
  return { paper, parser: new OptionSetParser(numbering) };
}

/** The visible text of the auto-lettered option paragraphs, in document order. */
function optionTexts(block: QuestionBlock): string[] {
  return block.nodes
    .filter((node) => node !== block.questionParagraph && node.namespaceURI === NS.w && node.localName === 'p')
    .filter((node) => {
      const pPr = firstChild(node, NS.w, 'pPr');
      return !!pPr && !!firstChild(pPr, NS.w, 'numPr');
    })
    .map((node) => visibleText(node).trim());
}

describe('auto-lettered options', () => {
  it('are recognised', async () => {
    const { paper, parser } = await open(autoLetteredPaper());
    const parsed = parser.parse(paper.sections[0]!.blocks[0]!);

    expect(parsed.ok).toBe(true);
    if (!parsed.ok) return;
    expect(parsed.options.layout).toBe('auto-lettered');
    expect(parsed.options.signatures).toHaveLength(4);
  });

  it('move their content between the paragraphs, leaving the numbering in place', async () => {
    const { paper, parser } = await open(autoLetteredPaper());
    const block = paper.sections[0]!.blocks[0]!;
    expect(optionTexts(block)).toEqual(AUTO_OPTIONS);

    const parsed = parser.parse(block);
    if (!parsed.ok) throw new Error('fixture should parse');
    parsed.options.apply([2, 3, 0, 1]);

    expect(optionTexts(block)).toEqual(['third choice', 'fourth choice', 'first choice', 'second choice']);

    // Every option paragraph still carries its own numbering, so Word re-letters in place.
    // (The stem is numbered too - that is the question list - so it is excluded here.)
    const numbered = block.nodes.filter((node) => {
      if (node === block.questionParagraph) return false;
      const pPr = firstChild(node, NS.w, 'pPr');
      return !!pPr && !!firstChild(pPr, NS.w, 'numPr');
    });
    expect(numbered).toHaveLength(4);
    for (const node of numbered) {
      const pPr = firstChild(node, NS.w, 'pPr')!;
      expect(childElements(pPr, NS.w, 'numPr')).toHaveLength(1);
    }
  });

  it('are reversible', async () => {
    const { paper, parser } = await open(autoLetteredPaper());
    const block = paper.sections[0]!.blocks[0]!;

    const first = parser.parse(block);
    if (!first.ok) throw new Error('fixture should parse');
    first.options.apply([1, 2, 3, 0]);

    const second = parser.parse(block);
    if (!second.ok) throw new Error('should still parse');
    second.options.apply([3, 0, 1, 2]);

    expect(optionTexts(block)).toEqual(AUTO_OPTIONS);
  });

  it('prefer the "(A)" list over a lettered "A." statement list of the same size', async () => {
    // The shape of Q89 in the sample paper: a lettered match-the-column list next to the options.
    const { paper, parser } = await open(
      autoLetteredPaper({ secondLetteredList: undefined, autoLettered: true }),
    );
    const block = paper.sections[0]!.blocks[0]!;
    const parsed = parser.parse(block);
    expect(parsed.ok).toBe(true);

    const withStatements = await open([
      {
        ...autoLetteredPaper()[0]!,
        questions: [
          {
            stem: 'Match the columns',
            // Statements are the "A." list; the options are the "(A)" list.
            optionParagraphs: AUTO_OPTIONS,
            answer: 'C',
            autoLettered: true,
            secondLetteredList: undefined,
          },
          ...autoLetteredPaper()[0]!.questions.slice(1),
        ],
      },
    ]);
    expect(withStatements.parser.parse(withStatements.paper.sections[0]!.blocks[0]!).ok).toBe(true);
  });

  it('are refused when two "(A)" lists could both be the options', async () => {
    const { paper, parser } = await open(
      autoLetteredPaper({ secondLetteredList: ['alpha', 'beta', 'gamma', 'delta'] }),
    );
    const parsed = parser.parse(paper.sections[0]!.blocks[0]!);

    expect(parsed.ok).toBe(false);
    if (parsed.ok) return;
    expect(parsed.reason).toBe('ambiguous-option-list');
  });

  it('are refused when the list does not have exactly four items', async () => {
    const { paper, parser } = await open(autoLetteredPaper({ optionParagraphs: ['one', 'two', 'three'] }));
    const parsed = parser.parse(paper.sections[0]!.blocks[0]!);

    expect(parsed.ok).toBe(false);
    if (parsed.ok) return;
    expect(parsed.reason).toBe('unexpected-option-count');
  });

  it('are refused when an option anchors a floating picture', async () => {
    const { paper, parser } = await open(
      autoLetteredPaper({
        optionParagraphs: AUTO_OPTIONS.slice(0, 3),
        rawOptionParagraph: floatingPictureOptionParagraph('').replace(
          '<w:pPr><w:pStyle w:val="ListParagraph"/></w:pPr>',
          '<w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="90"/></w:numPr></w:pPr>',
        ),
      }),
    );
    const parsed = parser.parse(paper.sections[0]!.blocks[0]!);

    expect(parsed.ok).toBe(false);
    if (parsed.ok) return;
    expect(parsed.reason).toBe('option-contains-floating-graphic');
  });

  it('do not swallow a typed-label question', async () => {
    const { paper, parser } = await open(defaultSections());
    const typed = parser.parse(paper.sections[0]!.blocks[0]!);
    expect(typed.ok).toBe(true);
    if (!typed.ok) return;
    expect(typed.options.layout).toBe('typed-labels');
  });
});

describe('auto-lettered options end to end', () => {
  let workingDir = '';
  let sourceFile = '';
  const service = new GenerationService();

  const request = (overrides: Partial<GenerationRequest> = {}): GenerationRequest => ({
    sourceFile,
    shuffleQuestions: false,
    shuffleOptions: true,
    questionExclusions: [],
    optionExclusions: [],
    setCount: 1,
    seed: 'auto-lettered',
    ...overrides,
  });

  beforeEach(async () => {
    workingDir = await fs.mkdtemp(path.join(os.tmpdir(), 'shuffler-auto-'));
    sourceFile = path.join(workingDir, 'Auto lettered.docx');
    await fs.writeFile(sourceFile, await buildPaper(autoLetteredPaper()));
  });

  afterEach(async () => {
    await fs.rm(workingDir, { recursive: true, force: true });
  });

  it('shuffles them and moves the answer key with the content', async () => {
    const result = await service.generate(request());
    expect(result.sets[0]!.verification.ok, JSON.stringify(result.sets[0]!.verification.checks)).toBe(true);

    const pkg = await DocxPackage.fromBuffer(await fs.readFile(result.sets[0]!.filePath));
    const numbering = new NumberingIndex(pkg.numberingPart());
    const paper = new PaperParser().parsePart(pkg.documentPart(), numbering);
    const block = paper.sections[0]!.blocks[0]!;

    const texts = optionTexts(block);
    // Same four options, in a different order.
    expect([...texts].sort()).toEqual([...AUTO_OPTIONS].sort());
    expect(texts).not.toEqual(AUTO_OPTIONS);

    // The original answer was C = "third choice"; the new key must point at that text.
    const newAnswer = paper.answerKey.answerOf(1)!;
    expect(texts[OPTION_LETTERS.indexOf(newAnswer)]).toBe('third choice');
  });

  it('is reported as shufflable by the dry run', async () => {
    const report = await service.dryRun(request());
    expect(report.optionsKeptByTool).toEqual([]);
    expect(report.optionsToShuffle).toBe(7);
  });
});
