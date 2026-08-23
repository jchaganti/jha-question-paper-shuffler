import type { QuestionBlock } from '../parse/PaperModel';
import type { NumberingIndex } from '../parse/NumberingIndex';
import { AutoLetteredOptionParser } from './AutoLetteredOptionParser';
import { OptionBlockParser, slotSignature } from './OptionBlockParser';
import { OptionShuffleApplier } from './OptionShuffleApplier';
import type { IOptionSetParser, OptionSet, OptionSetResult } from './OptionSet';

/**
 * Reads the options of a question, whichever way the paper labels them.
 *
 * Typed "(A) ... (B) ..." labels are tried first because they are unambiguous; auto-lettered
 * lists are tried next. When both fail, the more specific of the two complaints is reported,
 * so the user sees "found labels A,B,D" rather than "no labels found".
 */
export class OptionSetParser implements IOptionSetParser {
  private readonly autoLettered: AutoLetteredOptionParser;

  private readonly typedLabels: OptionBlockParser;

  constructor(
    numbering: NumberingIndex,
    private readonly applier: OptionShuffleApplier = new OptionShuffleApplier(),
  ) {
    // The typed-label reader needs the numbering too: some questions type only (B), (C)
    // and (D) and let Word letter the first option.
    this.typedLabels = new OptionBlockParser(numbering);
    this.autoLettered = new AutoLetteredOptionParser(numbering);
  }

  parse(block: QuestionBlock): OptionSetResult {
    const typed = this.typedLabels.parse(block);
    if (typed.ok) {
      return { ok: true, options: new TypedLabelOptionSet(typed.block, this.applier) };
    }

    const auto = this.autoLettered.parse(block);
    if (auto.ok) return auto;

    const vague = 'options-not-found';
    if (typed.reason !== vague) return typed;
    if (auto.reason !== vague) return auto;
    return typed;
  }
}

type ParsedTypedLabels = Extract<ReturnType<OptionBlockParser['parse']>, { ok: true }>['block'];

class TypedLabelOptionSet implements OptionSet {
  readonly layout = 'typed-labels' as const;

  constructor(
    private readonly block: ParsedTypedLabels,
    private readonly applier: OptionShuffleApplier,
  ) {}

  get signatures(): string[] {
    return this.block.slots.map((slot) => slotSignature(slot.coreAtoms));
  }

  apply(permutation: readonly number[]): void {
    this.applier.apply(this.block, permutation);
  }
}
