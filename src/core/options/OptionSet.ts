import type { OptionLayoutNote, SkipReason } from '../../shared/types';
import type { QuestionBlock } from '../parse/PaperModel';

/**
 * The four options of a question, ready to be permuted.
 *
 * Papers label options in two quite different ways, so each layout gets its own
 * implementation of this interface and the rest of the application does not care which
 * one it is holding:
 *
 *  - `typed-labels`  - "(A) ... (B) ..." typed into the paragraph text;
 *  - `auto-lettered` - one option per paragraph, the "(A)" produced by Word numbering.
 */
export interface OptionSet {
  readonly layout: 'typed-labels' | 'auto-lettered';
  /** Content fingerprint per slot, in A, B, C, D order. */
  readonly signatures: readonly string[];
  /** Rewrites the document so that slot *i* shows what slot `permutation[i]` had. */
  apply(permutation: readonly number[]): void;
}

export type OptionSetResult =
  | {
      readonly ok: true;
      readonly options: OptionSet;
      /**
       * Empty when the options were read with certainty. Otherwise, what had to be worked
       * out - the options are still shuffled, but the document should be corrected.
       */
      readonly notes: readonly OptionLayoutNote[];
    }
  | {
      readonly ok: false;
      readonly reason: SkipReason;
      /** What is wrong, naming the option it is wrong with. */
      readonly detail: string;
      /** What to change in Word so that this question can be shuffled. */
      readonly fix: string;
    };

export interface IOptionSetParser {
  parse(block: QuestionBlock): OptionSetResult;
}
