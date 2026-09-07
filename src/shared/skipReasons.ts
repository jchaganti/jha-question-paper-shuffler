import type { SkipReason, SkippedOptionGroup, SkippedOptionShuffle } from './types';

/**
 * How a question that could not be read is described to the person who typed the paper.
 *
 * The `SkipReason` itself is a name for the code to switch on - `option-spans-paragraphs`
 * says nothing to someone whose job is typing question papers. Every place that shows a
 * skipped question shows this heading instead, and the question's own `detail` underneath
 * it, so the same words reach the dry run, the report and the UI.
 *
 * The renderer deliberately does not import this module: `renderer.js` is loaded straight
 * by the browser with no bundler, so a runtime import would have to resolve on its own.
 * Only *types* cross that boundary; the values it displays arrive over IPC.
 */
export const SKIP_REASON_LABEL: Record<SkipReason, string> = {
  'excluded-by-user': 'You asked for this question to keep its option order',
  'options-not-found': 'The four options could not be found',
  'unexpected-label-sequence': 'The four option labels are not all there',
  'option-spans-paragraphs': 'An option carries on to the next line',
  'option-contains-floating-graphic': 'A picture inside an option floats instead of sitting in the line',
  'options-inside-table': 'The options are in a table',
  'label-bracket-is-a-symbol': 'An option label was put in with Insert, Symbol',
  'label-after-spaces-not-tab': 'Options on one line are spaced apart instead of tabbed',
  'unexpected-option-count': 'Word is lettering this question, but not four options',
  'ambiguous-option-list': 'Two lists here look like the options',
};

/**
 * Skipped questions gathered by problem, most-affected first.
 *
 * A paper typed one way tends to go wrong the same way many times over: 24 questions of one
 * corpus paper lay their options out in a table. Printing the same paragraph 24 times buries
 * the one question that went wrong for its own reason, so the fix is stated once per problem
 * and the questions are listed under it - the same shape the layout notes already use.
 */
export function groupSkippedOptions(items: readonly SkippedOptionShuffle[]): SkippedOptionGroup[] {
  const byReason = new Map<SkipReason, { fix: string; items: SkippedOptionShuffle[] }>();
  for (const item of items) {
    const entry = byReason.get(item.reason) ?? { fix: item.fix, items: [] };
    entry.items.push(item);
    byReason.set(item.reason, entry);
  }

  return [...byReason.entries()]
    .map(([reason, entry]) => {
      const details = new Set(entry.items.map((item) => item.detail));
      return {
        reason,
        label: SKIP_REASON_LABEL[reason],
        fix: entry.fix,
        questionNumbers: entry.items.map((item) => item.questionNumber).sort((a, b) => a - b),
        questions: [...entry.items].sort((a, b) => a.questionNumber - b.questionNumber),
        // 24 copies of one sentence is the noise this grouping exists to remove.
        ...(details.size === 1 ? { sharedDetail: [...details][0]! } : {}),
      };
    })
    .sort((a, b) => b.questionNumbers.length - a.questionNumbers.length);
}
