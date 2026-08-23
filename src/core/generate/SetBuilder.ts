import { OPTION_LETTERS, type OptionLetter, type QuestionMapping, type SkippedOptionShuffle } from '../../shared/types';
import { DocxPackage } from '../docx/DocxPackage';
import type { Element } from '../docx/dom';
import {
  NS,
  childElements,
  createElement,
  firstChild,
  ownerDocumentOf,
  preserveSpace,
  replaceChildren,
  visibleText,
} from '../docx/xml';
import { OptionSetParser } from '../options/OptionSetParser';
import { NumberingIndex } from '../parse/NumberingIndex';
import { PaperParser } from '../parse/PaperParser';
import type { ParsedPaper } from '../parse/PaperModel';
import type { SetPlan } from '../shuffle/ShufflePlanner';
import { DEFAULT_ANSWER_KEY_HEADINGS } from '../parse/PaperModel';
import { isHeadingParagraph } from '../parse/AnswerKeyTable';

/** Reports build progress as a 0..1 fraction of this set's work, with a step name. */
export type BuildStepListener = (fraction: number, message: string) => void;

export interface BuildSetInput {
  readonly sourceBuffer: Buffer;
  readonly plan: SetPlan;
  /** Answer letters of the *original* paper, by printed question number. */
  readonly originalAnswers: ReadonlyMap<number, OptionLetter>;
  /** Text appended to the answer-key title, e.g. "SET 01". Empty to disable. */
  readonly setLabel: string;
  readonly onStep?: BuildStepListener;
}

/**
 * Share of one set's work per step. Packaging (zipping ~1.3 MB of equations and images)
 * dominates, so the weights are deliberately not uniform.
 */
const STEP = {
  parsed: 0.1,
  optionsFrom: 0.1,
  optionsTo: 0.5,
  ordered: 0.6,
  packaged: 1,
} as const;

export interface BuiltSet {
  readonly buffer: Buffer;
  readonly mappings: QuestionMapping[];
  readonly questionsMoved: number;
  readonly optionsShuffled: number;
  readonly skipped: SkippedOptionShuffle[];
}

/**
 * Produces one shuffled set.
 *
 * The document is *edited*, never rebuilt: question blocks are re-ordered as whole
 * slices of body XML (so stems, images, tables and equations travel together), option
 * contents are swapped between labels, and the existing answer-key cells are rewritten.
 */
export class SetBuilder {
  constructor(private readonly parser: PaperParser = new PaperParser()) {}

  async build(input: BuildSetInput): Promise<BuiltSet> {
    const report = input.onStep ?? ((): void => {});

    const pkg = await DocxPackage.fromBuffer(input.sourceBuffer);
    const part = pkg.documentPart();
    const numbering = new NumberingIndex(pkg.numberingPart());
    const paper = this.parser.parsePart(part, numbering);
    const optionParser = new OptionSetParser(numbering);
    report(STEP.parsed, 'Reading the paper');

    const skipped: SkippedOptionShuffle[] = [];
    let optionsShuffled = 0;

    // 1. Options: swap content between the (A)-(D) labels.
    const allBlocks = paper.sections.flatMap((section) =>
      section.blocks.map((block) => ({ section, block })),
    );
    allBlocks.forEach(({ section, block }, index) => {
      if (index % 20 === 0) {
        const done = allBlocks.length === 0 ? 1 : index / allBlocks.length;
        report(STEP.optionsFrom + (STEP.optionsTo - STEP.optionsFrom) * done, 'Shuffling options');
      }
      const permutation = input.plan.optionPermutations.get(block.printedNumber);
      if (!permutation) return;
      const parsed = optionParser.parse(block);
      if (!parsed.ok) {
        skipped.push({
          questionNumber: block.printedNumber,
          subject: section.subject,
          reason: parsed.reason,
          detail: parsed.detail,
        });
        return;
      }
      parsed.options.apply(permutation);
      optionsShuffled++;
    });
    report(STEP.optionsTo, 'Shuffling options');

    // 2. Question order + answer key.
    const mappings: QuestionMapping[] = [];
    let questionsMoved = 0;

    const newBodyNodes: Element[] = [...paper.preambleNodes];
    paper.sections.forEach((section, sectionIndex) => {
      const order = input.plan.sectionOrders[sectionIndex] ?? section.blocks.map((_, i) => i);
      newBodyNodes.push(...section.headerNodes);

      order.forEach((sourceIndex, newIndex) => {
        const block = section.blocks[sourceIndex];
        if (!block) throw new Error(`Plan references question ${sourceIndex} of ${section.subject}, which does not exist`);
        newBodyNodes.push(...block.nodes);

        const newNumber = section.blocks[newIndex]!.printedNumber;
        const originalNumber = block.printedNumber;
        if (newNumber !== originalNumber) questionsMoved++;

        const originalAnswer = input.originalAnswers.get(originalNumber);
        if (!originalAnswer) {
          throw new Error(`The original answer key has no answer for question ${originalNumber}`);
        }
        const permutation = input.plan.optionPermutations.get(originalNumber);
        const wasShuffled = !!permutation && !skipped.some((s) => s.questionNumber === originalNumber);
        const newAnswer = wasShuffled ? mapAnswer(originalAnswer, permutation!) : originalAnswer;

        paper.answerKey.setAnswer(newNumber, newAnswer);
        mappings.push({
          newNumber,
          originalNumber,
          originalAnswer,
          newAnswer,
          optionMapping: wasShuffled ? describePermutation(permutation!) : '',
        });
      });

      newBodyNodes.push(...section.tailNodes);
    });

    newBodyNodes.push(...paper.answerKeyNodes);
    replaceChildren(paper.body, newBodyNodes);
    report(STEP.ordered, 'Re-ordering questions and updating the answer key');

    // 3. Stamp the set label onto the answer-key title line.
    if (input.setLabel) stampSetLabel(paper, input.setLabel);

    pkg.setPartText('word/document.xml', part.serialize());
    const buffer = await pkg.toBuffer();
    report(STEP.packaged, 'Packaging the document');
    return { buffer, mappings, questionsMoved, optionsShuffled, skipped };
  }
}

/** `permutation[newSlot] = originalSlot`, so the new letter is where the old slot went. */
export function mapAnswer(originalAnswer: OptionLetter, permutation: readonly number[]): OptionLetter {
  const originalSlot = OPTION_LETTERS.indexOf(originalAnswer);
  const newSlot = permutation.findIndex((source) => source === originalSlot);
  if (newSlot < 0) throw new Error(`Permutation ${permutation.join(',')} does not contain slot ${originalAnswer}`);
  return OPTION_LETTERS[newSlot]!;
}

export function describePermutation(permutation: readonly number[]): string {
  return permutation
    .map((source, target) => `${OPTION_LETTERS[source]}→${OPTION_LETTERS[target]}`)
    .sort()
    .join(', ');
}

/**
 * Appends e.g. " - SET 01" to the title under the ANSWER KEY heading so that a printed
 * key can never be matched with the wrong paper. Falls back to the heading itself.
 */
function stampSetLabel(paper: ParsedPaper, label: string): void {
  const nodes = paper.answerKeyNodes;
  const headingIndex = nodes.findIndex((node) => isHeadingParagraph(node, DEFAULT_ANSWER_KEY_HEADINGS));
  let target: Element | undefined;
  if (headingIndex >= 0) {
    for (let i = headingIndex + 1; i < nodes.length; i++) {
      const node = nodes[i]!;
      if (node.namespaceURI !== NS.w || node.localName !== 'p') break;
      if (visibleText(node).trim() !== '') {
        target = node;
        break;
      }
    }
    target ??= nodes[headingIndex];
  }
  if (!target) return;

  const doc = ownerDocumentOf(target);
  const runs = childElements(target, NS.w, 'r');
  const template = runs[runs.length - 1];
  const run = createElement(doc, 'w:r', NS.w);
  const rPr = template ? firstChild(template, NS.w, 'rPr') : undefined;
  if (rPr) run.appendChild(rPr.cloneNode(true));
  const text = createElement(doc, 'w:t', NS.w);
  preserveSpace(text);
  text.appendChild(doc.createTextNode(`  –  ${label}`));
  run.appendChild(text);
  target.appendChild(run);
}
