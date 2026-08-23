/**
 * These tests pin down how the tool behaves on papers that differ from the sample.
 * They are the executable version of the "Assumptions" and "Known limitations"
 * sections of the README: a variation is either
 *   - supported,
 *   - degraded (that question keeps its option order and is reported), or
 *   - a hard, explained failure (never a silent wrong answer key).
 */
import { describe, expect, it } from 'vitest';
import { DocxPackage } from '../src/core/docx/DocxPackage';
import { XmlPart } from '../src/core/docx/xml';
import { GenerationService } from '../src/core/generate/GenerationService';
import { OptionBlockParser } from '../src/core/options/OptionBlockParser';
import { NumberingIndex } from '../src/core/parse/NumberingIndex';
import { PaperParseError, PaperParser } from '../src/core/parse/PaperParser';
import type { QuestionBlock } from '../src/core/parse/PaperModel';
import { buildPaper, defaultSections, type FixtureSection } from './support/PaperFixture';

async function parse(buffer: Buffer) {
  const pkg = await DocxPackage.fromBuffer(buffer);
  return new PaperParser().parsePart(pkg.documentPart(), new NumberingIndex(pkg.numberingPart()));
}

/** Rebuilds the fixture with `word/document.xml` rewritten by `edit`. */
async function patched(edit: (xml: string) => string, sections: readonly FixtureSection[] = defaultSections()) {
  const pkg = await DocxPackage.fromBuffer(await buildPaper(sections));
  const xml = edit(pkg.getPartText('word/document.xml')!);
  return () => new PaperParser().parsePart(XmlPart.parse(xml), new NumberingIndex(pkg.numberingPart()));
}

/**
 * A paper whose *first* question uses the given option paragraphs, padded with plain
 * questions so that the answer key is large enough to be recognised as one.
 */
function oneQuestion(optionParagraphs: readonly string[]): FixtureSection[] {
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
      questions: [{ stem: 'A question', optionParagraphs, answer: 'B' }, ...filler],
    },
  ];
}

async function firstBlock(sections: readonly FixtureSection[]): Promise<QuestionBlock> {
  const paper = await parse(await buildPaper(sections));
  return paper.sections[0]!.blocks[0]!;
}

describe('supported variations', () => {
  it('accepts "A)" labels without brackets', async () => {
    const parsed = new OptionBlockParser().parse(await firstBlock(oneQuestion(['A)\tone\tB)\ttwo\tC)\tthree\tD)\tfour'])));
    expect(parsed.ok).toBe(true);
  });

  it('accepts lower-case "(a)" labels', async () => {
    const parsed = new OptionBlockParser().parse(await firstBlock(oneQuestion(['(a)\tone\t(b)\ttwo', '(c)\tthree\t(d)\tfour'])));
    expect(parsed.ok).toBe(true);
    if (!parsed.ok) return;
    expect(parsed.block.slots.map((slot) => slot.letter)).toEqual(['A', 'B', 'C', 'D']);
  });

  it('accepts a stem that spills over several paragraphs before the options', async () => {
    const parsed = new OptionBlockParser().parse(
      await firstBlock(oneQuestion(['Statement I: something', 'Statement II: something else', '(A)\tone\t(B)\ttwo\t(C)\tthree\t(D)\tfour'])),
    );
    expect(parsed.ok).toBe(true);
  });

  it('treats a paper with no subject headings as a single subject', async () => {
    const [physics, chemistry] = defaultSections();
    const paper = await parse(
      await buildPaper([
        { ...physics!, subject: 'Part I - untitled' },
        { ...chemistry!, subject: 'Part II - untitled' },
      ]),
    );
    expect(paper.sections).toHaveLength(1);
    expect(paper.sections[0]!.subject).toBe('ALL');
    expect(paper.sections[0]!.blocks).toHaveLength(7);
    // Visible in the UI summary, so the user can see that shuffling is paper-wide.
  });
});

describe('degraded variations (question is skipped and reported)', () => {
  const skipReasonFor = async (optionParagraphs: readonly string[]): Promise<string> => {
    const parsed = new OptionBlockParser().parse(await firstBlock(oneQuestion(optionParagraphs)));
    return parsed.ok ? 'accepted' : parsed.reason;
  };

  it('refuses a five-option question instead of moving "(E)" around with (D)', async () => {
    expect(await skipReasonFor(['(A)\tone\t(B)\ttwo', '(C)\tthree\t(D)\tfour', '(E)\tfive'])).toBe(
      'unexpected-label-sequence',
    );
  });

  it('refuses a question with only three options', async () => {
    expect(await skipReasonFor(['(A)\tone\t(B)\ttwo\t(C)\tthree'])).toBe('unexpected-label-sequence');
  });

  it('refuses options that continue onto the next paragraph', async () => {
    expect(
      await skipReasonFor(['(A)\tone', 'continued text for option A', '(B)\ttwo', '(C)\tthree', '(D)\tfour']),
    ).toBe('option-spans-paragraphs');
  });

  it('refuses labels that appear out of order', async () => {
    expect(await skipReasonFor(['(A)\tone\t(C)\tthree', '(B)\ttwo\t(D)\tfour'])).toBe('unexpected-label-sequence');
  });

  it('a skipped question is listed in the summary the UI shows', async () => {
    // The first question has three options, so it cannot be shuffled either way.
    const file = await writeTemp(await buildPaper(oneQuestion(['(A)\tone\t(B)\ttwo\t(C)\tthree'])));
    const summary = await new GenerationService().inspect(file);

    expect(summary.unshufflableOptions.map((item) => item.questionNumber)).toEqual([1]);
    expect(summary.unshufflableOptions[0]!.reason).toBe('unexpected-label-sequence');
  });
});

describe('hard failures (explained, never silent)', () => {
  it('rejects question numbers that were typed by hand', async () => {
    const run = await patched((xml) =>
      xml.replace(/<w:numPr><w:ilvl w:val="0"\/><w:numId w:val="[12]"\/><\/w:numPr>/g, ''),
    );
    expect(run).toThrow(PaperParseError);
    expect(run).toThrow(/Could not identify the question numbering|Found 0 numbered questions/);
  });

  it('rejects an answer key whose letters are bracketed, e.g. "(A)"', async () => {
    const run = await patched((xml) =>
      xml.replace(/<w:t>([A-D])<\/w:t>/g, (_match, letter: string) => `<w:t>(${letter})</w:t>`),
    );
    expect(run).toThrow(/No answer key found/);
  });

  it('rejects a paper whose subjects restart numbering at 1', async () => {
    // Both subjects number from 1, so a key entry no longer identifies one question.
    const question = (n: number) => ({
      stem: `Question ${n}`,
      optionParagraphs: ['(A)\tone\t(B)\ttwo\t(C)\tthree\t(D)\tfour'],
      answer: 'A' as const,
    });
    const restarted: FixtureSection[] = [
      { subject: 'PHYSICS', startNumber: 1, numId: '1', questions: Array.from({ length: 12 }, (_u, i) => question(i + 1)) },
      { subject: 'CHEMISTRY', startNumber: 1, numId: '2', questions: Array.from({ length: 12 }, (_u, i) => question(i + 1)) },
    ];
    await expect(parse(await buildPaper(restarted))).rejects.toThrow(PaperParseError);
    await expect(parse(await buildPaper(restarted))).rejects.toThrow(
      /more than once|do not continue from one another/,
    );
  });

  it('rejects a .doc (binary Word 97-2003) file', async () => {
    await expect(
      new GenerationService().generate({
        sourceFile: 'C:/papers/paper.doc',
        shuffleQuestions: true,
        shuffleOptions: true,
        questionExclusions: [],
        optionExclusions: [],
        setCount: 1,
      }),
    ).rejects.toThrow(/must be a \.docx file/);
  });

  it('rejects a file that is not a Word package at all', async () => {
    await expect(DocxPackage.fromBuffer(Buffer.from('PK\u0003\u0004 not really a zip'))).rejects.toThrow();
  });
});

/** Writes a fixture to a temp file so the service (which reads from disk) can open it. */
async function writeTemp(buffer: Buffer): Promise<string> {
  const { promises: fs } = await import('node:fs');
  const os = await import('node:os');
  const path = await import('node:path');
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'shuffler-var-'));
  const file = path.join(dir, 'Paper.docx');
  await fs.writeFile(file, buffer);
  return file;
}
