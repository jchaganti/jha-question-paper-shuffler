import type { LayoutNoteGroup, OptionLayoutIssue, QuestionLayoutNote } from './types';

/**
 * Grouping of option-layout notes, used by the CLI and the report writer, and applied in
 * the main process so that the grouped result reaches the UI ready to display.
 *
 * The renderer deliberately does not import this module: `renderer.js` is loaded straight
 * by the browser with no bundler, so a runtime import would have to resolve on its own.
 * Only *types* cross that boundary; the values it displays arrive over IPC.
 */

/** Plain-language heading for each issue, for someone who only types the paper. */
export const LAYOUT_ISSUE_LABEL: Record<OptionLayoutIssue, string> = {
  'mixed-auto-and-typed-labels': 'Some options are lettered by Word and the rest typed by hand',
  'label-not-after-tab': 'An option label has no tab in front of it',
  'mixed-label-case': 'The four option labels mix capital and small letters',
  'several-lettered-lists': 'More than one lettered list could have been the options',
  'floating-picture-in-option-area': 'A floating picture sits among the options',
};

export function groupLayoutNotes(notes: readonly QuestionLayoutNote[]): LayoutNoteGroup[] {
  const byIssue = new Map<OptionLayoutIssue, { fix: string; numbers: Set<number> }>();
  for (const note of notes) {
    const entry = byIssue.get(note.issue) ?? { fix: note.fix, numbers: new Set<number>() };
    entry.numbers.add(note.questionNumber);
    byIssue.set(note.issue, entry);
  }

  return [...byIssue.entries()]
    .map(([issue, entry]) => ({
      issue,
      label: LAYOUT_ISSUE_LABEL[issue],
      fix: entry.fix,
      questionNumbers: [...entry.numbers].sort((a, b) => a - b),
    }))
    // Most-affected first: that is the one worth fixing before the next paper.
    .sort((a, b) => b.questionNumbers.length - a.questionNumbers.length);
}

/** Distinct questions covered by a set of notes (a question can carry more than one). */
export function questionsWithLayoutNotes(notes: readonly QuestionLayoutNote[]): number[] {
  return [...new Set(notes.map((note) => note.questionNumber))].sort((a, b) => a - b);
}
