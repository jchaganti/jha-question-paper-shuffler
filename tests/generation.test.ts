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
import { pdfLines, pdfPageCount, pdfText } from './support/pdfText';

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
    expect(summary.groups.map((g) => g.group)).toEqual(['PHYSICS', 'CHEMISTRY']);
    // Question 4 is auto-lettered by Word; that layout is supported, so nothing is skipped.
    expect(summary.unshufflableOptions).toEqual([]);
    // Question 5 has a "None of these" option.
    expect(summary.advisories.map((item) => item.questionNumber)).toContain(5);
  });

  it('writes N files into a dated folder and passes its own verification', async () => {
    const result = await service.generate(request({ setCount: 3 }));

    expect(path.basename(result.outputFolder)).toMatch(
      /^question-sets - \d{2}-\d{2}-\d{4}-\d{2}-\d{2}$/,
    );
    // The run's date and time sit in the name, so match the shape rather than a literal.
    // `setFileName.test.ts` pins the exact format.
    expect(result.sets.map((set) => set.fileName)).toEqual([
      expect.stringMatching(/^Sample Paper - \d{2}-\d{2}-\d{4}-\d{2}-\d{2}-Set-01\.docx$/),
      expect.stringMatching(/^Sample Paper - \d{2}-\d{2}-\d{4}-\d{2}-\d{2}-Set-02\.docx$/),
      expect.stringMatching(/^Sample Paper - \d{2}-\d{2}-\d{4}-\d{2}-\d{2}-Set-03\.docx$/),
    ]);
    for (const set of result.sets) {
      expect(set.verification.ok, JSON.stringify(set.verification.checks)).toBe(true);
      await expect(fs.stat(set.filePath)).resolves.toBeTruthy();
    }
    await expect(fs.stat(result.reportFile)).resolves.toBeTruthy();
  });

  it('never overwrites an earlier run, even one started in the same minute', async () => {
    const first = await service.generate(request({ setCount: 1 }));
    const second = await service.generate(request({ setCount: 1 }));

    expect(second.outputFolder).not.toBe(first.outputFolder);
    // Two runs a second apart share a timestamp, so the second folder is marked "-02".
    expect(path.basename(second.outputFolder)).toMatch(
      /^question-sets - \d{2}-\d{2}-\d{4}-\d{2}-\d{2}(-02)?$/,
    );
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
    expect(pdfText(await fs.readFile(result.reportFile))).toContain('Seed march-batch');
  });

  it('writes the report as a PDF a reader can open', async () => {
    const result = await service.generate(request({ setCount: 1 }));
    const pdf = await fs.readFile(result.reportFile);

    expect(result.reportFile.endsWith('_generation-report.pdf')).toBe(true);
    expect(pdf.subarray(0, 8).toString('latin1')).toBe('%PDF-1.4');
    expect(pdf.subarray(-7).toString('latin1').trim()).toBe('%%EOF');
    expect(pdfPageCount(pdf)).toBeGreaterThan(0);
    expect(pdfText(pdf)).toContain('Question set generation report');
  });

  describe('the report names options the way the paper names them', () => {
    /** The mapping rows of the report, which is where answers are quoted. */
    const mappingRows = (pdf: Buffer): string[] =>
      pdfLines(pdf).filter((line) => /^\d+ \d+ \S/.test(line));

    it('quotes letters for a paper whose key is written A, B, C, D', async () => {
      const result = await service.generate(request({ setCount: 1 }));
      const rows = mappingRows(await fs.readFile(result.reportFile));

      expect(rows.length).toBeGreaterThan(0);
      expect(rows.join('\n')).toMatch(/^\d+ \d+ [ABCD] [ABCD] /m);
      expect(rows.join('\n')).not.toMatch(/\([1-4]\)/);
    });

    it('quotes digits for a paper whose key is written (1), (2), (3), (4)', async () => {
      // The shape of FST-1-Rep-2024: the paper never writes A-D anywhere, so neither
      // should the report that describes it.
      await fs.writeFile(
        sourceFile,
        await buildPaper(defaultSections(), { answerKeyAnswerFormat: 'digit-bracketed' }),
      );
      const result = await service.generate(request({ setCount: 1 }));
      const rows = mappingRows(await fs.readFile(result.reportFile));

      expect(rows.length).toBeGreaterThan(0);
      // Original answer and New answer columns.
      expect(rows.join('\n')).toMatch(/^\d+ \d+ \([1-4]\) \([1-4]\) /m);
      // The option mapping column. WinAnsi has no arrow, so the PDF spells it "->".
      expect(rows.some((row) => /\([1-4]\)->\([1-4]\)/.test(row))).toBe(true);
      // No stray A-D anywhere in a mapping row.
      expect(rows.join('\n')).not.toMatch(/^\d+ \d+ [ABCD] |[ABCD]->/m);
    });

    it('quotes roman numerals for a paper whose key is written (i)...(iv)', async () => {
      await fs.writeFile(
        sourceFile,
        await buildPaper(defaultSections(), { answerKeyAnswerFormat: 'roman-bracketed' }),
      );
      const result = await service.generate(request({ setCount: 1 }));
      const rows = mappingRows(await fs.readFile(result.reportFile));

      expect(rows.length).toBeGreaterThan(0);
      expect(rows.join('\n')).toMatch(/^\d+ \d+ \((?:i|ii|iii|iv)\) \((?:i|ii|iii|iv)\) /m);
      expect(rows.join('\n')).not.toMatch(/^\d+ \d+ [ABCD] |[ABCD]->/m);
    });
  });

  describe('set file names carry the run date and time', () => {
    const NAME_RE = /^Sample Paper - (\d{2}-\d{2}-\d{4}-\d{2}-\d{2})-Set-(\d{2})\.docx$/;

    it('names every file "<paper> - DD-MM-YYYY-HH-MM-Set-NN.docx"', async () => {
      const result = await service.generate(request({ setCount: 3 }));

      for (const set of result.sets) {
        expect(set.fileName).toMatch(NAME_RE);
        expect(NAME_RE.exec(set.fileName)![2]).toBe(String(set.setNumber).padStart(2, '0'));
      }
    });

    it('gives every set of one run the same timestamp', async () => {
      // Stamped once when the run starts, so a run crossing a minute boundary still
      // produces one consistently named batch.
      const result = await service.generate(request({ setCount: 3 }));

      const stamps = new Set(result.sets.map((set) => NAME_RE.exec(set.fileName)![1]));
      expect(stamps.size).toBe(1);
    });

    it('uses that same instant for the report and for GenerationResult', async () => {
      const result = await service.generate(request({ setCount: 1 }));
      const report = pdfText(await fs.readFile(result.reportFile));

      const at = new Date(result.generatedAt);
      const two = (n: number) => String(n).padStart(2, '0');
      expect(report).toContain(
        `generated ${two(at.getDate())}-${two(at.getMonth() + 1)}-${at.getFullYear()} ` +
          `at ${two(at.getHours())}:${two(at.getMinutes())}`,
      );
      const expected = new Date(result.generatedAt);
      const pad = (n: number) => String(n).padStart(2, '0');
      expect(NAME_RE.exec(result.sets[0]!.fileName)![1]).toBe(
        `${pad(expected.getDate())}-${pad(expected.getMonth() + 1)}-${expected.getFullYear()}` +
          `-${pad(expected.getHours())}-${pad(expected.getMinutes())}`,
      );
    });

    it('writes the file it reports, under the name it reports', async () => {
      const result = await service.generate(request({ setCount: 2 }));

      for (const set of result.sets) {
        expect(path.basename(set.filePath)).toBe(set.fileName);
        await expect(fs.stat(set.filePath)).resolves.toBeTruthy();
      }
    });
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
