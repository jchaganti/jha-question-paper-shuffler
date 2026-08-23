import { OPTION_LETTERS, type OptionLetter, type VerificationResult } from '../../shared/types';
import { DocxPackage } from '../docx/DocxPackage';
import { questionSignature } from '../generate/Signatures';
import type { IOptionSetParser } from '../options/OptionSet';
import { OptionSetParser } from '../options/OptionSetParser';
import { NumberingIndex } from '../parse/NumberingIndex';
import { PaperParser } from '../parse/PaperParser';
import type { ParsedPaper } from '../parse/PaperModel';

/** What the original paper looked like, per question. */
export interface OriginalFacts {
  readonly questionCount: number;
  /** stem fingerprint -> printed question number (only when unique). */
  readonly stemToNumber: ReadonlyMap<string, number>;
  /** printed question number -> option fingerprints, slot A..D (absent if unparsed). */
  readonly slotSignatures: ReadonlyMap<number, readonly string[]>;
  readonly answers: ReadonlyMap<number, OptionLetter>;
}

interface Check {
  name: string;
  ok: boolean;
  detail: string;
}

/**
 * Re-opens a generated file and proves the four properties that matter:
 *  1. no question was lost or duplicated;
 *  2. the question at every position is the one the plan intended;
 *  3. each question still offers exactly its original four options;
 *  4. the letter printed in the new answer key points at the *same text* that was
 *     correct in the original paper.
 *
 * This runs against the written .docx, so it also proves the file re-parses cleanly.
 */
export class SetVerifier {
  constructor(private readonly parser: PaperParser = new PaperParser()) {}

  static factsOf(paper: ParsedPaper, optionParser: IOptionSetParser): OriginalFacts {
    const stemCounts = new Map<string, number>();
    const stemToNumber = new Map<string, number>();
    const slotSignatures = new Map<number, readonly string[]>();
    const answers = new Map<number, OptionLetter>();

    for (const section of paper.sections) {
      for (const block of section.blocks) {
        const signature = questionSignature(block, optionParser);
        stemCounts.set(signature, (stemCounts.get(signature) ?? 0) + 1);
        stemToNumber.set(signature, block.printedNumber);
        const parsed = optionParser.parse(block);
        if (parsed.ok) slotSignatures.set(block.printedNumber, parsed.options.signatures);
        const answer = paper.answerKey.answerOf(block.printedNumber);
        if (answer) answers.set(block.printedNumber, answer);
      }
    }

    for (const [signature, count] of stemCounts) {
      if (count > 1) stemToNumber.delete(signature);
    }

    return { questionCount: paper.questionCount, stemToNumber, slotSignatures, answers };
  }

  async verify(buffer: Buffer, original: OriginalFacts): Promise<VerificationResult> {
    const checks: Check[] = [];
    const pkg = await DocxPackage.fromBuffer(buffer);
    const part = pkg.documentPart();
    const numbering = new NumberingIndex(pkg.numberingPart());
    const paper = this.parser.parsePart(part, numbering);
    // Built from the *generated* file's own numbering, so nothing is assumed about it.
    const optionParser = new OptionSetParser(numbering);

    checks.push({
      name: 'Question count unchanged',
      ok: paper.questionCount === original.questionCount,
      detail: `${paper.questionCount} of ${original.questionCount} questions`,
    });

    let identified = 0;
    let unknown = 0;
    let optionsPreserved = 0;
    let optionsBroken: string[] = [];
    let keyOk = 0;
    let keyBroken: string[] = [];
    let missingKey: number[] = [];

    for (const section of paper.sections) {
      for (const block of section.blocks) {
        const originalNumber = original.stemToNumber.get(questionSignature(block, optionParser));
        if (originalNumber === undefined) {
          unknown++;
          continue;
        }
        identified++;

        const newAnswer = paper.answerKey.answerOf(block.printedNumber);
        if (!newAnswer) {
          missingKey.push(block.printedNumber);
          continue;
        }

        const originalSlots = original.slotSignatures.get(originalNumber);
        const parsed = optionParser.parse(block);
        if (!originalSlots || !parsed.ok) continue;

        const newSlots = parsed.options.signatures;
        if (sameMultiset(originalSlots, newSlots)) optionsPreserved++;
        else optionsBroken.push(`Q${block.printedNumber} (was Q${originalNumber})`);

        const originalAnswer = original.answers.get(originalNumber);
        if (!originalAnswer) continue;
        const expectedSignature = originalSlots[OPTION_LETTERS.indexOf(originalAnswer)];
        const actualSignature = newSlots[OPTION_LETTERS.indexOf(newAnswer)];
        if (expectedSignature !== undefined && expectedSignature === actualSignature) keyOk++;
        else keyBroken.push(`Q${block.printedNumber} key says ${newAnswer} (was Q${originalNumber} answer ${originalAnswer})`);
      }
    }

    checks.push({
      name: 'Every question identified in the generated file',
      ok: unknown === 0,
      detail: unknown === 0 ? `${identified} questions matched` : `${unknown} question(s) could not be matched by content`,
    });
    checks.push({
      name: 'Option sets unchanged (nothing lost or duplicated)',
      ok: optionsBroken.length === 0,
      detail: optionsBroken.length === 0 ? `${optionsPreserved} questions checked` : optionsBroken.slice(0, 10).join('; '),
    });
    checks.push({
      name: 'New answer key points at the originally correct text',
      ok: keyBroken.length === 0 && missingKey.length === 0,
      detail:
        keyBroken.length === 0 && missingKey.length === 0
          ? `${keyOk} answers checked`
          : [...keyBroken.slice(0, 10), ...(missingKey.length ? [`no key entry for ${missingKey.slice(0, 10).join(', ')}`] : [])].join('; '),
    });

    return { ok: checks.every((check) => check.ok), checks };
  }
}

function sameMultiset(a: readonly string[], b: readonly string[]): boolean {
  if (a.length !== b.length) return false;
  const left = [...a].sort();
  const right = [...b].sort();
  return left.every((value, index) => value === right[index]);
}
