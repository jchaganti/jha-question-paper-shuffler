import type { ShuffleAccounting } from './types';

/** How to name each of the three outcomes for a particular shuffle. */
export interface AccountingLabels {
  /** "free to move", "with their options shuffled". */
  readonly shuffled: string;
  /** "held in place by a picture", "the tool cannot read". */
  readonly keptByTool: string;
  /** "you asked to keep in place", "you asked to keep in order". */
  readonly keptByUser: string;
}

/** Labels for the question shuffle. */
export const QUESTION_LABELS: AccountingLabels = {
  shuffled: 'free to move',
  keptByTool: 'held in place by a picture',
  keptByUser: 'you asked to keep in place',
};

/** Labels for the option shuffle. */
export const OPTION_LABELS: AccountingLabels = {
  shuffled: 'with their options shuffled',
  keptByTool: 'the tool cannot read',
  keptByUser: 'you asked to keep in order',
};

/**
 * Splits a paper into "shuffled", "kept by the tool" and "kept because you asked".
 *
 * Every question of the paper lands in exactly one of the three, which is the whole point:
 * the dry run can then show its own arithmetic instead of leaving the reader to work out
 * where the questions it did not mention went. See `ShuffleAccounting`.
 *
 * Numbers the user typed that this paper does not have are dropped rather than silently
 * counted - they belong to some other paper, and `unknownNumbers` reports them separately.
 */
export function accountFor(input: {
  /** Every question number printed in the paper. */
  readonly questionNumbers: readonly number[];
  /** Whether this shuffle is switched on at all. */
  readonly active: boolean;
  /** Numbers the tool refuses to shuffle, whatever the user asks. */
  readonly keptByTool: readonly number[];
  /** Numbers the user asked to keep, as typed - including any this paper does not have. */
  readonly keptByUser: readonly number[];
  /** How to name the three outcomes in `summary`. */
  readonly labels: AccountingLabels;
}): ShuffleAccounting {
  const total = input.questionNumbers.length;
  if (!input.active) return { total, active: false, shuffled: 0, keptByTool: [], keptByUser: [] };

  const inPaper = new Set(input.questionNumbers);
  const tool = ascending(input.keptByTool.filter((n) => inPaper.has(n)));
  const toolSet = new Set(tool);
  // The tool's list wins where the two overlap, so a question is never counted twice.
  const user = ascending(input.keptByUser.filter((n) => inPaper.has(n) && !toolSet.has(n)));
  const accounting: ShuffleAccounting = {
    total,
    active: true,
    shuffled: total - tool.length - user.length,
    keptByTool: tool,
    keptByUser: user,
  };
  return { ...accounting, summary: accountingLine(accounting, input.labels) };
}

/**
 * Says that some of the numbers in an exclusion list do not exist in this paper.
 *
 * Almost always the sign of a list left over from a different paper - the numbers mean
 * nothing here - so the message points at the list rather than at the paper. Kept beside
 * the accounting because it is the other half of the same idea: which numbers are real.
 */
export function staleListWarning(missing: readonly number[], field: string): string {
  const one = missing.length === 1;
  return (
    `"${field}" lists ${one ? 'question' : 'questions'} ${missing.join(', ')}, but this paper ` +
    `has no such question ${one ? 'number' : 'numbers'}, so ${one ? 'it does' : 'they do'} ` +
    'nothing. Question numbers only mean something in the paper they came from, so this list ' +
    'was most likely typed for a different paper. Clear out anything that does not belong to ' +
    'this one.'
  );
}

/** The numbers the user typed that this paper does not have, ascending. */
export function unknownNumbers(
  questionNumbers: readonly number[],
  typed: readonly number[],
): number[] {
  const inPaper = new Set(questionNumbers);
  return ascending(typed.filter((n) => !inPaper.has(n)));
}

/**
 * One line of arithmetic: "77 with their options shuffled + 3 the tool cannot read +
 * 20 you asked to keep in order = 100".
 *
 * Written here rather than at each of the three places that show it, so the UI, the CLI and
 * the PDF report cannot word the same sum differently. Returns undefined when the shuffle
 * is switched off, because there is then no sum to show.
 */
export function accountingLine(
  accounting: ShuffleAccounting,
  labels: AccountingLabels,
): string | undefined {
  if (!accounting.active) return undefined;
  const parts = [`${accounting.shuffled} ${labels.shuffled}`];
  if (accounting.keptByTool.length > 0) {
    parts.push(`${accounting.keptByTool.length} ${labels.keptByTool}`);
  }
  if (accounting.keptByUser.length > 0) {
    parts.push(`${accounting.keptByUser.length} ${labels.keptByUser}`);
  }
  return `${parts.join(' + ')} = ${accounting.total}`;
}

function ascending(numbers: readonly number[]): number[] {
  return [...new Set(numbers)].sort((a, b) => a - b);
}
