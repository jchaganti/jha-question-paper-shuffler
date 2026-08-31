/**
 * Which questions are worth keeping in their original option order.
 *
 * The test is **position dependence**, not question type: an option is flagged only when
 * its meaning points at something outside itself. "None of these" means *none of the
 * options above it*, and "Both (A) and (B)" names two of them, so moving either changes
 * what it asserts.
 *
 * An Assertion-Reason question is not flagged for being one. Its standard four options -
 * "Only statement I is true", "Assertion and Reason are true and Reason is the correct
 * explanation of Assertion" - each state their own meaning in full, so they move safely
 * and the key is remapped for them like any other question. Only the legend layout is
 * position-dependent: some papers print the wording once as directions above a run of
 * questions and leave the options as bare tokens, and there the meaning lives outside the
 * option.
 */
import { describe, expect, it } from 'vitest';
import { DocxPackage } from '../src/core/docx/DocxPackage';
import { OptionAdvisor } from '../src/core/generate/OptionAdvisor';
import { OptionSetParser } from '../src/core/options/OptionSetParser';
import { NumberingIndex } from '../src/core/parse/NumberingIndex';
import { PaperParser } from '../src/core/parse/PaperParser';
import { buildPaper, type FixtureQuestion, type FixtureSection } from './support/PaperFixture';

function paperWith(question: FixtureQuestion): FixtureSection[] {
  const filler = Array.from({ length: 6 }, (_unused, index) => ({
    stem: `Filler question ${index + 1}`,
    optionParagraphs: ['(A)\tone\t(B)\ttwo\t(C)\tthree\t(D)\tfour'],
    answer: 'A' as const,
  }));
  return [{ subject: 'PHYSICS', startNumber: 1, numId: '1', questions: [question, ...filler] }];
}

/** Advisories for the first question of a paper built around it. */
async function adviseFirst(question: FixtureQuestion) {
  const pkg = await DocxPackage.fromBuffer(await buildPaper(paperWith(question)));
  const numbering = new NumberingIndex(pkg.numberingPart());
  const paper = new PaperParser().parsePart(pkg.documentPart(), numbering);
  const parser = new OptionSetParser(numbering, paper.answerKey.style.scheme);
  return new OptionAdvisor().advise(paper, parser).filter((item) => item.questionNumber === 1);
}

/** The four standard Assertion-Reason options, each stating its own meaning. */
const SELF_DESCRIBING_OPTIONS = [
  '(A)\tAssertion and Reason are true and Reason is the correct explanation of Assertion',
  '(B)\tAssertion and Reason are true but Reason is not the correct explanation of Assertion',
  '(C)\tAssertion is true and Reason is false',
  '(D)\tBoth Assertion and Reason are false',
];

describe('an Assertion-Reason question whose options say what they mean', () => {
  it('is not flagged - the options carry their meaning wherever they sit', async () => {
    expect(
      await adviseFirst({
        stem: 'Assertion: Petunia belongs to Solanaceae. Reason: Datura belongs to Brassicaceae.',
        optionParagraphs: SELF_DESCRIBING_OPTIONS,
        answer: 'C',
      }),
    ).toEqual([]);
  });

  it('is not flagged in the Statement-I / Statement-II wording either', async () => {
    expect(
      await adviseFirst({
        stem: 'Statement I: iron rusts. Statement II: rusting is reduction.',
        optionParagraphs: [
          '(A)\tOnly statement I is true',
          '(B)\tOnly statement II is true',
          '(C)\tBoth statements are true',
          // "wrong" rather than "false" - papers use both, and either states its meaning.
          '(D)\tBoth statements are wrong',
        ],
        answer: 'A',
      }),
    ).toEqual([]);
  });
});

describe('an Assertion-Reason question printed with its wording as directions', () => {
  it('is flagged, because each option is a bare token whose meaning is elsewhere', async () => {
    const advisories = await adviseFirst({
      stem: 'Assertion: Petunia belongs to Solanaceae. Reason: Datura belongs to Brassicaceae.',
      // The legend layout: "If both A and R are true ... mark (1)" is printed above, so the
      // options themselves say nothing.
      optionParagraphs: ['(A)\t1', '(B)\t2', '(C)\t3', '(D)\t4'],
      answer: 'C',
    });

    expect(advisories.map((item) => item.kind)).toEqual(['assertion-reason']);
    expect(advisories[0]!.detail).toMatch(/do not say what they mean/);
  });
});

describe('options that genuinely depend on their position', () => {
  it('flags a catch-all option', async () => {
    const advisories = await adviseFirst({
      stem: 'Which of these is a vector?',
      optionParagraphs: ['(A)\tspeed\t(B)\tmass\t(C)\tvelocity\t(D)\tNone of these'],
      answer: 'C',
    });

    expect(advisories.map((item) => item.kind)).toEqual(['catch-all-option']);
  });

  it('flags an option that names two other options', async () => {
    const advisories = await adviseFirst({
      stem: 'Which reactant is limiting?',
      optionParagraphs: ['(A)\tCaCO3\t(B)\tHCl', '(C)\tBoth (A) and (B)\t(D)\tneither one'],
      answer: 'C',
    });

    expect(advisories.map((item) => item.kind)).toEqual(['references-other-option']);
  });

  it('reads the options, not the stem - a stem saying "all of the above" is not a flag', async () => {
    // The stem describes the question; only an *option* can depend on where it sits.
    expect(
      await adviseFirst({
        stem: 'Considering all of the above reactions, which rate is highest?',
        optionParagraphs: ['(A)\tfirst\t(B)\tsecond\t(C)\tthird\t(D)\tfourth'],
        answer: 'A',
      }),
    ).toEqual([]);
  });

  it('leaves an ordinary question unflagged', async () => {
    expect(
      await adviseFirst({
        stem: 'What is the SI unit of force?',
        optionParagraphs: ['(A)\tnewton\t(B)\tjoule\t(C)\twatt\t(D)\tpascal'],
        answer: 'A',
      }),
    ).toEqual([]);
  });
});

describe('when the options cannot be read at all', () => {
  it('falls back to searching the whole question, so nothing is missed', async () => {
    // Three options, so the parser refuses. Such a question is already reported as
    // unshufflable, and the dry run drops it from the suggestions - but the advisor itself
    // errs towards flagging rather than staying silent.
    const advisories = await adviseFirst({
      stem: 'Which of these is a vector?',
      optionParagraphs: ['(A)\tspeed\t(B)\tmass\t(C)\tNone of these'],
      answer: 'A',
    });

    expect(advisories.map((item) => item.kind)).toEqual(['catch-all-option']);
  });
});
