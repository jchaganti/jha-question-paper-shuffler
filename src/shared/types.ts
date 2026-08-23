/**
 * Contracts shared between the Electron main process, the renderer (UI) and the CLI.
 * This module is intentionally free of runtime dependencies so that it can be
 * imported from any of the three environments.
 */

/** The four answer slots of a multiple-choice question. */
export const OPTION_LETTERS = ['A', 'B', 'C', 'D'] as const;
export type OptionLetter = (typeof OPTION_LETTERS)[number];

/** Everything the user supplies on the UI (or on the command line). */
export interface GenerationRequest {
  /** Absolute path of the original question paper (.docx). */
  readonly sourceFile: string;
  /** Shuffle the position of questions inside each subject. */
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
  /** An auto-lettered list was found, but not with exactly four items. */
  | 'unexpected-option-count'
  /** Several lettered lists could be the options; the tool will not guess. */
  | 'ambiguous-option-list';

export interface SkippedOptionShuffle {
  readonly questionNumber: number;
  readonly subject: string;
  readonly reason: SkipReason;
  readonly detail: string;
}

/** An advisory: the option text looks position-dependent, so shuffling may change meaning. */
export interface OptionAdvisory {
  readonly questionNumber: number;
  readonly subject: string;
  readonly kind: 'catch-all-option' | 'references-other-option' | 'assertion-reason';
  readonly detail: string;
}

export interface SubjectSummary {
  readonly subject: string;
  readonly firstQuestionNumber: number;
  readonly lastQuestionNumber: number;
  readonly questionCount: number;
}

/** What the parser understood about the source paper. Shown in the UI before generating. */
export interface PaperSummary {
  readonly sourceFile: string;
  readonly questionCount: number;
  readonly subjects: readonly SubjectSummary[];
  /** Questions whose options cannot be shuffled safely, whatever the user asks. */
  readonly unshufflableOptions: readonly SkippedOptionShuffle[];
  /** Questions worth adding to the "options not shuffled" exclusion list. */
  readonly advisories: readonly OptionAdvisory[];
}

/** Per subject, what a dry run expects to happen with the current settings. */
export interface DryRunSubject {
  readonly subject: string;
  readonly questionCount: number;
  /** Questions free to move (not pinned by the exclusion list). */
  readonly movable: number;
  /** Questions whose options will be permuted. */
  readonly optionsShuffled: number;
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
  readonly subjects: readonly DryRunSubject[];
  /** Questions free to move, in total. */
  readonly questionsEligibleToMove: number;
  /** Questions whose options will be permuted, in total. */
  readonly optionsToShuffle: number;
  /** Questions whose options stay put because the user asked. */
  readonly optionsKeptByUser: readonly number[];
  /** Questions whose options stay put because they cannot be parsed with certainty. */
  readonly optionsKeptByTool: readonly SkippedOptionShuffle[];
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
