import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { DocxPackage } from '../src/core/docx/DocxPackage';
import { GenerationService } from '../src/core/generate/GenerationService';
import { OptionBlockParser, slotSignature } from '../src/core/options/OptionBlockParser';
import { NumberingIndex } from '../src/core/parse/NumberingIndex';
import { PaperParser } from '../src/core/parse/PaperParser';
import { OPTION_LETTERS, type GenerationRequest } from '../src/shared/types';
import { buildPaper, defaultSections } from './support/PaperFixture';

let workingDir = '';
let sourceFile = '';

const service = new GenerationService();

const request = (overrides: Partial<GenerationRequest> = {}): GenerationRequest => ({
  sourceFile,
  shuffleQuestions: true,
  shuffleOptions: true,
  questionExclusions: [],
  optionExclusions: [],
  setCount: 2,
  seed: 'end-to-end',
  ...overrides,
});

/** Reads a generated paper back: question order, options and key, per position. */
async function readPaper(file: string) {
  const pkg = await DocxPackage.fromBuffer(await fs.readFile(file));
  const paper = new PaperParser().parsePart(pkg.documentPart(), new NumberingIndex(pkg.numberingPart()));
  const parser = new OptionBlockParser();
  return paper.sections.flatMap((section) =>
    section.blocks.map((block) => {
      const parsed = parser.parse(block);
      return {
        subject: section.subject,
        number: block.printedNumber,
        stem: block.questionParagraph.textContent?.trim() ?? '',
        answer: paper.answerKey.answerOf(block.printedNumber),
        options: parsed.ok ? parsed.block.slots.map((slot) => slotSignature(slot.coreAtoms)) : undefined,
      };
    }),
  );
}

beforeEach(async () => {
  workingDir = await fs.mkdtemp(path.join(os.tmpdir(), 'shuffler-test-'));
  sourceFile = path.join(workingDir, 'Sample Paper.docx');
  await fs.writeFile(sourceFile, await buildPaper(defaultSections()));
});

afterEach(async () => {
  await fs.rm(workingDir, { recursive: true, force: true });
});

describe('GenerationService', () => {
  it('reports what it found in the paper', async () => {
    const summary = await service.inspect(sourceFile);
    expect(summary.questionCount).toBe(7);
    expect(summary.subjects.map((s) => s.subject)).toEqual(['PHYSICS', 'CHEMISTRY']);
    // Question 4 is auto-lettered by Word; that layout is supported, so nothing is skipped.
    expect(summary.unshufflableOptions).toEqual([]);
    // Question 5 has a "None of these" option.
    expect(summary.advisories.map((item) => item.questionNumber)).toContain(5);
  });

  it('writes N files into question-sets-01 and passes its own verification', async () => {
    const result = await service.generate(request({ setCount: 3 }));

    expect(path.basename(result.outputFolder)).toBe('question-sets-01');
    expect(result.sets.map((set) => set.fileName)).toEqual([
      'Sample Paper - Set 01.docx',
      'Sample Paper - Set 02.docx',
      'Sample Paper - Set 03.docx',
    ]);
    for (const set of result.sets) {
      expect(set.verification.ok, JSON.stringify(set.verification.checks)).toBe(true);
      await expect(fs.stat(set.filePath)).resolves.toBeTruthy();
    }
    await expect(fs.stat(result.reportFile)).resolves.toBeTruthy();
  });

  it('creates question-sets-02 on the next run', async () => {
    const first = await service.generate(request({ setCount: 1 }));
    const second = await service.generate(request({ setCount: 1 }));
    expect(path.basename(first.outputFolder)).toBe('question-sets-01');
    expect(path.basename(second.outputFolder)).toBe('question-sets-02');
  });

  it('keeps every question inside its own subject', async () => {
    const original = await readPaper(sourceFile);
    const result = await service.generate(request({ setCount: 1 }));
    const generated = await readPaper(result.sets[0]!.filePath);

    const subjectOf = new Map(original.map((q) => [q.stem, q.subject]));
    for (const question of generated) {
      expect(subjectOf.get(question.stem)).toBe(question.subject);
    }
  });

  it('renumbers questions 1..N without gaps', async () => {
    const result = await service.generate(request({ setCount: 1 }));
    const generated = await readPaper(result.sets[0]!.filePath);
    expect(generated.map((q) => q.number)).toEqual([1, 2, 3, 4, 5, 6, 7]);
  });

  it('produces an answer key that points at the originally correct option', async () => {
    const original = await readPaper(sourceFile);
    const byStem = new Map(original.map((q) => [q.stem, q]));

    const result = await service.generate(request({ setCount: 2 }));
    for (const set of result.sets) {
      const generated = await readPaper(set.filePath);
      for (const question of generated) {
        const source = byStem.get(question.stem);
        expect(source, question.stem).toBeTruthy();
        if (!source?.options || !question.options || !source.answer || !question.answer) continue;

        const correctContent = source.options[OPTION_LETTERS.indexOf(source.answer)];
        const keyedContent = question.options[OPTION_LETTERS.indexOf(question.answer)];
        expect(keyedContent, `Q${question.number} of ${set.fileName}`).toBe(correctContent);
      }
    }
  });

  it('offers exactly the same four options, only in another order', async () => {
    const original = await readPaper(sourceFile);
    const byStem = new Map(original.map((q) => [q.stem, q]));
    const result = await service.generate(request({ setCount: 1 }));
    const generated = await readPaper(result.sets[0]!.filePath);

    for (const question of generated) {
      const source = byStem.get(question.stem);
      if (!source?.options || !question.options) continue;
      expect([...question.options].sort()).toEqual([...source.options].sort());
    }
  });

  it('keeps excluded questions at their position and excluded options in order', async () => {
    const original = await readPaper(sourceFile);
    const result = await service.generate(
      request({ setCount: 1, questionExclusions: [2, 6], optionExclusions: [1] }),
    );
    const generated = await readPaper(result.sets[0]!.filePath);

    const originalByNumber = new Map(original.map((q) => [q.number, q]));
    expect(generated.find((q) => q.number === 2)!.stem).toBe(originalByNumber.get(2)!.stem);
    expect(generated.find((q) => q.number === 6)!.stem).toBe(originalByNumber.get(6)!.stem);

    const question1 = generated.find((q) => q.stem === originalByNumber.get(1)!.stem)!;
    expect(question1.options).toEqual(originalByNumber.get(1)!.options);
  });

  it('does not change question order when only options are shuffled', async () => {
    const original = await readPaper(sourceFile);
    const result = await service.generate(request({ setCount: 1, shuffleQuestions: false }));
    const generated = await readPaper(result.sets[0]!.filePath);

    expect(generated.map((q) => q.stem)).toEqual(original.map((q) => q.stem));
    expect(generated.map((q) => q.options)).not.toEqual(original.map((q) => q.options));
  });

  it('does not change options when only questions are shuffled', async () => {
    const original = await readPaper(sourceFile);
    const byStem = new Map(original.map((q) => [q.stem, q]));
    const result = await service.generate(request({ setCount: 1, shuffleOptions: false }));
    const generated = await readPaper(result.sets[0]!.filePath);

    for (const question of generated) {
      expect(question.options).toEqual(byStem.get(question.stem)!.options);
      expect(question.answer).toBe(byStem.get(question.stem)!.answer);
    }
  });

  it('reports the seed it used, and that seed reproduces the same sets', async () => {
    // No seed given: the service invents one and must report it back.
    const first = await service.generate(request({ setCount: 2, seed: undefined }));
    expect(first.seed).toBeTruthy();

    const second = await service.generate(request({ setCount: 2, seed: first.seed }));

    for (let i = 0; i < 2; i++) {
      const before = await DocxPackage.fromBuffer(await fs.readFile(first.sets[i]!.filePath));
      const after = await DocxPackage.fromBuffer(await fs.readFile(second.sets[i]!.filePath));
      expect(after.getPartText('word/document.xml'), `set ${i + 1}`).toBe(
        before.getPartText('word/document.xml'),
      );
    }
  });

  it('produces different sets when no seed is given', async () => {
    const first = await service.generate(request({ setCount: 1, seed: undefined }));
    const second = await service.generate(request({ setCount: 1, seed: undefined }));
    expect(second.seed).not.toBe(first.seed);

    const a = await DocxPackage.fromBuffer(await fs.readFile(first.sets[0]!.filePath));
    const b = await DocxPackage.fromBuffer(await fs.readFile(second.sets[0]!.filePath));
    expect(b.getPartText('word/document.xml')).not.toBe(a.getPartText('word/document.xml'));
  });

  it('records the run seed in the report', async () => {
    const result = await service.generate(request({ setCount: 1, seed: 'march-batch' }));
    const report = await fs.readFile(result.reportFile, 'utf8');
    expect(report).toContain('- Seed: `march-batch`');
  });

  it('stamps the set number onto the answer key page', async () => {
    const result = await service.generate(request({ setCount: 1 }));
    const pkg = await DocxPackage.fromBuffer(await fs.readFile(result.sets[0]!.filePath));
    expect(pkg.getPartText('word/document.xml')).toContain('SET 01');
  });

  it('leaves every other part of the document untouched', async () => {
    const result = await service.generate(request({ setCount: 1 }));
    const before = await DocxPackage.fromBuffer(await fs.readFile(sourceFile));
    const after = await DocxPackage.fromBuffer(await fs.readFile(result.sets[0]!.filePath));

    for (const part of ['[Content_Types].xml', '_rels/.rels', 'word/_rels/document.xml.rels', 'word/numbering.xml']) {
      expect(after.getPartText(part), part).toBe(before.getPartText(part));
    }
  });

  it('rejects invalid input', async () => {
    await expect(service.generate(request({ setCount: 0 }))).rejects.toThrow(/between 1 and 100/);
    await expect(
      service.generate(request({ shuffleQuestions: false, shuffleOptions: false })),
    ).rejects.toThrow(/at least one/);
    await expect(service.generate(request({ sourceFile: 'paper.doc' }))).rejects.toThrow(/\.docx/);
  });
});
