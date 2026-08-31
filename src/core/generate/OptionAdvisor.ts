import type { OptionAdvisory } from '../../shared/types';
import { visibleText } from '../docx/xml';
import type { IOptionSetParser } from '../options/OptionSet';
import type { ParsedPaper, QuestionBlock } from '../parse/PaperModel';

const CATCH_ALL = /\b(none|all)\s+of\s+(these|the\s+above|them)\b/i;
const REFERENCES_OPTION = /\b(both|either|neither)\s*\(?[A-D]\)?\s*(?:and|or|&|,)\s*\(?[A-D]\)?/i;
const ASSERTION_REASON = /\bassertion\b[\s\S]*\breason\b|\bstatement[- ]?(i{1,3}|1|2)\b[\s\S]*\bstatement[- ]?(i{1,3}|1|2)\b/i;

/**
 * The words that make an assertion-reason option mean something on its own: "Both
 * statements are true", "Assertion is true and Reason is false".
 *
 * An option carrying any of them says what it means wherever it sits. One that carries
 * none is a bare token - "1", "(a)" - whose meaning lives in a legend printed elsewhere.
 */
const SELF_DESCRIBING =
  /\b(assertions?|reasons?|statements?|true|false|correct|incorrect|wrong|right|explanation)\b/i;

/**
 * Flags questions whose **option order carries meaning**, so the user can decide whether to
 * add them to the "options not shuffled" exclusion list.
 *
 * The test is position dependence, not question type. An option is position-dependent when
 * its meaning points at something outside itself:
 *
 *  - another option - "Both (A) and (B)";
 *  - the options above it - "None of these";
 *  - a legend printed elsewhere - a bare "(1)" under a "If both A and R are true, mark (1)"
 *    directions block.
 *
 * Assertion-reason questions are *not* flagged merely for being assertion-reason. The
 * standard four options - "Only statement I is true", "Assertion and Reason are true and
 * Reason is the correct explanation of Assertion" - each state their own meaning in full,
 * so they move safely and the key is remapped for them like any other question. Only the
 * legend layout, where the options are bare tokens, is position-dependent.
 *
 * Detection reads the **options**, not the whole question: a stem that happens to say "all
 * of the above" says nothing about whether the options can move. When the options cannot be
 * read at all the whole block is searched instead, which errs towards flagging - but such a
 * question is already reported as unshufflable, so the advice is moot.
 *
 * This is advice only - it never changes what the generator does.
 */
export class OptionAdvisor {
  advise(paper: ParsedPaper, optionParser?: IOptionSetParser): OptionAdvisory[] {
    const out: OptionAdvisory[] = [];

    for (const section of paper.sections) {
      for (const block of section.blocks) {
        const options = readOptions(block, optionParser);
        const optionText = options ? options.join('\n') : blockText(block);
        const add = (kind: OptionAdvisory['kind'], detail: string): void => {
          out.push({ questionNumber: block.printedNumber, subject: section.subject, kind, detail });
        };

        const referenceMatch = REFERENCES_OPTION.exec(optionText);
        if (referenceMatch) {
          add('references-other-option', `Option text refers to another option: "${referenceMatch[0].trim()}"`);
          continue;
        }
        const catchAllMatch = CATCH_ALL.exec(optionText);
        if (catchAllMatch) {
          add('catch-all-option', `Contains a catch-all option: "${catchAllMatch[0].trim()}"`);
          continue;
        }
        // The question type is read from the whole block, because the "Assertion:" and
        // "Reason:" lines are in the stem; whether it is *position-dependent* is then read
        // from the options alone.
        if (ASSERTION_REASON.test(blockText(block)) && !selfDescribing(options)) {
          add(
            'assertion-reason',
            'Assertion-Reason / Statement-I-II question whose options do not say what they mean - ' +
              'their wording is printed once as directions above. Moving them would change which ' +
              'direction each option points at.',
          );
        }
      }
    }

    return out;
  }
}

/** True when every option states its own meaning, so none depends on where it sits. */
function selfDescribing(options: readonly string[] | undefined): boolean {
  return options !== undefined && options.length > 0 && options.every((option) => SELF_DESCRIBING.test(option));
}

function blockText(block: QuestionBlock): string {
  return block.nodes.map((node) => visibleText(node)).join(' ');
}

/** The four option texts, or undefined when this question's options cannot be read. */
function readOptions(block: QuestionBlock, optionParser?: IOptionSetParser): string[] | undefined {
  if (!optionParser) return undefined;
  const parsed = optionParser.parse(block);
  // A signature is "text|distinguishing tokens"; only the text matters here.
  return parsed.ok ? parsed.options.signatures.map((signature) => signature.split('|')[0]!) : undefined;
}
