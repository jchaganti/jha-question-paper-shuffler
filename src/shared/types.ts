/**
 * Contracts shared between the Electron main process, the renderer (UI) and the CLI.
 * This module is intentionally free of runtime dependencies so that it can be
 * imported from any of the three environments.
 */

import type { AnswerStyle } from './answerStyle';

/**
 * The four answer slots of a multiple-choice question.
 *
 * This is the tool's *internal* vocabulary: a slot is always A-D whatever the paper prints
 * on the page. Anything shown to the user is rendered back into the paper's own style with
 * `renderAnswer` - see `./answerStyle`.
 */
export const OPTION_LETTERS = ['A', 'B', 'C', 'D'] as const;
export type OptionLetter = (typeof OPTION_LETTERS)[number];

/** Everything the user supplies on the UI (or on the command line). */
export interface GenerationRequest {
  /** Absolute path of the original question paper (.docx). */
  readonly sourceFile: string;
  /** Shuffle the position of questions inside each section (a subject, or a part of one). */
  readonly shuffleQuestions: boolean;
  /** Shuffle the order of the four options of each question. */
  readonly shuffleOptions: boolean;
  /**
   * Printed question numbers (as they appear in the original paper, 1..N) whose
   * position must not change when questions are shuffled.
   */
  readonly questionExclusions: readonly number[];
  /**
   * Printed question numbers whose options must keep their original order.
   * Use this for questions such as "Both (A) and (B)" / "None of these" /
   * assertion-reason sets where the option order carries meaning.
   */
  readonly optionExclusions: readonly number[];
  /** Number of sets to generate (>= 1). */
  readonly setCount: number;
  /**
   * Keep each question whole on one page: a question that would not fit at the bottom of
   * a page starts on the next page instead of being split across two.
   *
   * Defaults to `true` when omitted. Turning it off leaves the page flow of the original
   * paper untouched, which keeps the page count down at the cost of split questions.
   */
  readonly keepQuestionsWhole?: boolean;
  /**
   * Optional seed. When supplied, the same seed + same inputs always produce
   * byte-identical sets, which makes a run reproducible and auditable.
   */
  readonly seed?: string;
}

/** Why the options of a particular question were left untouched. */
export type SkipReason =
  | 'excluded-by-user'
  | 'options-not-found'
  | 'unexpected-label-sequence'
  | 'option-spans-paragraphs'
  | 'option-contains-floating-graphic'
  | 'options-inside-table'
  /** A label's opening bracket is a symbol-font character, not ordinary text. */
  | 'label-bracket-is-a-symbol'
  /** Options on one line are separated by a run of spaces instead of a tab. */
  | 'label-after-spaces-not-tab'
  /** An auto-lettered list was found, but not with exactly four items. */
  | 'unexpected-option-count'
  /** Several lettered lists could be the options; the tool will not guess. */
  | 'ambiguous-option-list';

export interface SkippedOptionShuffle {
  readonly questionNumber: number;
  readonly subject: string;
  readonly reason: SkipReason;
  /** What is wrong with this question, in the author's terms, naming the option. */
  readonly detail: string;
  /** What to change in the Word document so that this question can be shuffled. */
  readonly fix: string;
}

/**
 * Skipped questions gathered by problem, ready to display - the same shape as
 * `LayoutNoteGroup`, and for the same reason: a paper typed one way goes wrong the same way
 * many times over, and the fix is worth stating once. See `shared/skipReasons.ts`.
 */
export interface SkippedOptionGroup {
  readonly reason: SkipReason;
  /** Plain-language heading for someone who only types the paper. */
  readonly label: string;
  /** What to change in the Word document. */
  readonly fix: string;
  /** Affected question numbers, ascending. */
  readonly questionNumbers: readonly number[];
  /** The questions themselves, in number order, each with its own detail. */
  readonly questions: readonly SkippedOptionShuffle[];
  /**
   * The one thing they all say, when every question in the group says the same thing -
   * which is the common case for a problem that is about the layout rather than about a
   * particular option. Shown once instead of once per question; undefined when the details
   * differ and each question has something of its own to say.
   */
  readonly sharedDetail?: string;
}

/**
 * Something about the way a question's options are labelled that the tool could work out,
 * but that the Word document should not have contained.
 *
 * These questions **are** shuffled - the parser has a fallback for each of them. They are
 * reported so that whoever types the paper can correct it, because a fallback is a reading
 * of an ambiguous document rather than a certainty.
 */
export type OptionLayoutIssue =
  /** Option (A) is lettered by Word while (B), (C) and (D) are typed by hand. */
  | 'mixed-auto-and-typed-labels'
  /** A label had no tab in front of it, so it was read from the punctuation before it. */
  | 'label-not-after-tab'
  /** The four labels do not all use the same case, e.g. "(A) (b) (C) (d)". */
  | 'mixed-label-case'
  /** Several lettered lists could have been the options; the likeliest one was used. */
  | 'several-lettered-lists'
  /** A floating picture is anchored among the options; it was left exactly where it is. */
  | 'floating-picture-in-option-area';

export interface OptionLayoutNote {
  readonly issue: OptionLayoutIssue;
  readonly detail: string;
  /** What to change in the Word document, in the author's terms. */
  readonly fix: string;
}

/** An `OptionLayoutNote` together with the question it came from. */
export interface QuestionLayoutNote extends OptionLayoutNote {
  readonly questionNumber: number;
  readonly subject: string;
}

/**
 * Layout notes gathered by problem, ready to display: an author fixing a paper wants
 * "these eight questions have the same problem, here is the fix", not the same sentence
 * repeated eight times. Grouped in the main process so that the UI, the CLI and the
 * report all show the same thing (see `shared/layoutNotes.ts`).
 */
export interface LayoutNoteGroup {
  readonly issue: OptionLayoutIssue;
  /** Plain-language heading for someone who only types the paper. */
  readonly label: string;
  /** What to change in the Word document. */
  readonly fix: string;
  /** Affected question numbers, ascending and without repeats. */
  readonly questionNumbers: readonly number[];
}

/** An advisory: the option text looks position-dependent, so shuffling may change meaning. */
export interface OptionAdvisory {
  readonly questionNumber: number;
  readonly subject: string;
  readonly kind: 'catch-all-option' | 'references-other-option' | 'assertion-reason';
  readonly detail: string;
}

/**
 * One run of questions that shuffle among themselves. Usually a subject, but a subject
 * divided into "SECTION A" and "SECTION B" - or into parts above those - contributes one
 * of these per division, because a question must never cross such a boundary.
 */
export interface GroupSummary {
  /** "PHYSICS", or "PHYSICS - SECTION B" when the subject is divided. */
  readonly group: string;
  readonly firstQuestionNumber: number;
  readonly lastQuestionNumber: number;
  readonly questionCount: number;
}

/**
 * Questions the tool keeps at their original positions, whatever the user asked, because
 * moving them apart would take a floating picture away from the question it illustrates.
 *
 * One of these per picture, naming every question it holds in place - usually a pair.
 */
export interface PinnedQuestionGroup {
  readonly questionNumbers: readonly number[];
  readonly subject: string;
  readonly detail: string;
  /** What to change in the Word document to lift the restriction. */
  readonly fix: string;
}

/** What the parser understood about the source paper. Shown in the UI before generating. */
export interface PaperSummary {
  readonly sourceFile: string;
  readonly questionCount: number;
  /** Every run of questions that shuffles on its own - one per subject, or per section. */
  readonly groups: readonly GroupSummary[];
  /**
   * How this paper names an option - letters `(A)`, digits `(1)`, roman numerals `(i)` -
   * read from its own answer key. Answers shown back to the user are rendered through it.
   */
  readonly answerStyle: AnswerStyle;
  /**
   * Every question number printed in this paper, ascending.
   *
   * The one place that says which numbers exist. Every count the report shows is derived
   * from this list, and so is every check on the numbers the user typed - so a number that
   * belongs to a different paper is always caught, and the report's arithmetic can never
   * disagree with its own section table.
   */
  readonly questionNumbers: readonly number[];
  /** Questions whose options cannot be shuffled safely, whatever the user asks. */
  readonly unshufflableOptions: readonly SkippedOptionShuffle[];
  /** The same questions gathered by problem, for display. Empty when the list above is. */
  readonly unshufflableGroups: readonly SkippedOptionGroup[];
  /** Questions that cannot be moved safely, whatever the user asks. */
  readonly pinnedQuestions: readonly PinnedQuestionGroup[];
  /** Questions worth adding to the "options not shuffled" exclusion list. */
  readonly advisories: readonly OptionAdvisory[];
  /**
   * Questions whose options *were* read and will be shuffled, but only because the parser
   * fell back on a heuristic. Reported so the Word document can be corrected.
   */
  readonly layoutNotes: readonly QuestionLayoutNote[];
  /** The same notes gathered by problem, for display. Empty when `layoutNotes` is. */
  readonly layoutNoteGroups: readonly LayoutNoteGroup[];
}

/** Per subject, what a dry run expects to happen with the current settings. */
export interface DryRunGroup {
  /** "PHYSICS", or "PHYSICS - SECTION B" when the subject is divided. */
  readonly group: string;
  readonly questionCount: number;
  /** Questions free to move (not pinned by the exclusion list). */
  readonly movable: number;
  /** Questions whose options will be permuted. */
  readonly optionsShuffled: number;
}

/**
 * Where every question of the paper ends up, for one of the two shuffles.
 *
 * The three outcomes are mutually exclusive and together cover the whole paper, so what the
 * dry run says can be read as arithmetic - shuffled + kept by the tool + kept by you = the
 * number of questions - rather than as three numbers that happen to appear near each other.
 * Without that, a headline such as "options would be shuffled for 77 of 100" leaves 23
 * questions unaccounted for and the report looks wrong even when it is right.
 *
 * A question the tool cannot touch counts as the tool's even when the user also listed it:
 * taking it off the list would change nothing, so saying it is kept "because you asked"
 * would send the user to the wrong place.
 */
export interface ShuffleAccounting {
  /** Questions in the paper. Equals `shuffled + keptByTool.length + keptByUser.length`. */
  readonly total: number;
  /** False when this shuffle is switched off - then nothing is shuffled and both lists are empty. */
  readonly active: boolean;
  readonly shuffled: number;
  /** Kept because the tool will not do it: unreadable options, or a picture holding a question. */
  readonly keptByTool: readonly number[];
  /** Kept only because the user listed the number, ascending. */
  readonly keptByUser: readonly number[];
  /**
   * The same arithmetic as one sentence - "77 with their options shuffled + 3 the tool
   * cannot read + 20 you asked to keep in order = 100".
   *
   * Worded once, in the main process, so the UI, the CLI and the PDF report cannot word it
   * differently - and so the renderer needs no runtime import to show it. Undefined when
   * this shuffle is switched off.
   */
  readonly summary?: string;
}

/**
 * The result of "Dry run": what *would* happen with the settings as they stand, without
 * writing a single file.
 */
export interface DryRunReport {
  readonly paper: PaperSummary;
  readonly setCount: number;
  readonly shuffleQuestions: boolean;
  readonly shuffleOptions: boolean;
  readonly groups: readonly DryRunGroup[];
  /** Questions free to move, in total. Same number as `questionAccounting.shuffled`. */
  readonly questionsEligibleToMove: number;
  /** Questions whose options will be permuted, in total. Same as `optionAccounting.shuffled`. */
  readonly optionsToShuffle: number;
  /** Every question of the paper, split by why it will or will not move. */
  readonly questionAccounting: ShuffleAccounting;
  /** Every question of the paper, split by why its options will or will not be permuted. */
  readonly optionAccounting: ShuffleAccounting;
  /** Questions whose options stay put because they cannot be parsed with certainty. */
  readonly optionsKeptByTool: readonly SkippedOptionShuffle[];
  /** Questions whose position stays put because a picture is anchored across the boundary. */
  readonly questionsKeptByTool: readonly PinnedQuestionGroup[];
  /** Position-dependent options not yet in the exclusion list - worth adding. */
  readonly suggestedForExclusion: readonly OptionAdvisory[];
  /** Anything the user should read before generating (bad exclusion numbers, etc.). */
  readonly warnings: readonly string[];
}

export interface QuestionMapping {
  /** Question number in the generated set. */
  readonly newNumber: number;
  /** Question number this content had in the original paper. */
  readonly originalNumber: number;
  /** Answer letter in the original paper. */
  readonly originalAnswer: OptionLetter;
  /** Answer letter in the generated set. */
  readonly newAnswer: OptionLetter;
  /** Option letter mapping, e.g. "A→C, B→A, C→D, D→B"; empty when options were kept. */
  readonly optionMapping: string;
}

export interface GeneratedSet {
  readonly setNumber: number;
  readonly fileName: string;
  readonly filePath: string;
  readonly seed: string;
  readonly questionsMoved: number;
  readonly optionsShuffled: number;
  /** Questions marked to stay whole on one page; 0 when that option was switched off. */
  readonly questionsKeptWhole: number;
  readonly skipped: readonly SkippedOptionShuffle[];
  readonly mappings: readonly QuestionMapping[];
  readonly verification: VerificationResult;
}

export interface VerificationResult {
  readonly ok: boolean;
  readonly checks: readonly { readonly name: string; readonly ok: boolean; readonly detail: string }[];
}

export interface GenerationResult {
  readonly outputFolder: string;
  readonly reportFile: string;
  /**
   * The seed this run actually used - what the user typed, or the random one invented for
   * them. Passing it back as `GenerationRequest.seed` reproduces these exact sets.
   */
  readonly seed: string;
  /**
   * When this run started, ISO 8601. The same instant is stamped into every set's file
   * name, so a file can always be traced back to the run that produced it.
   */
  readonly generatedAt: string;
  readonly paper: PaperSummary;
  readonly sets: readonly GeneratedSet[];
}

export type ProgressStage =
  | 'reading'
  | 'parsing'
  | 'planning'
  | 'options'
  | 'ordering'
  | 'packaging'
  | 'writing'
  | 'verifying'
  | 'done';

export interface ProgressEvent {
  readonly stage: ProgressStage;
  /** Human readable step, e.g. "Shuffling options". */
  readonly message: string;
  /** 1-based index of the set being built, when a set is being built. */
  readonly setNumber?: number;
  readonly setCount?: number;
  /** 0..1 progress inside the current set. */
  readonly setFraction?: number;
  /** 0..1 progress over the whole run. */
  readonly fraction?: number;
}

/** Contract exposed to the renderer through the preload bridge. */
export interface ShufflerApi {
  pickSourceFile(): Promise<string | null>;
  inspect(sourceFile: string): Promise<Result<PaperSummary>>;
  dryRun(request: GenerationRequest): Promise<Result<DryRunReport>>;
  generate(request: GenerationRequest): Promise<Result<GenerationResult>>;
  revealFolder(folder: string): Promise<void>;
  onProgress(listener: (event: ProgressEvent) => void): void;
}

export type Result<T> = { readonly ok: true; readonly value: T } | { readonly ok: false; readonly error: string };
