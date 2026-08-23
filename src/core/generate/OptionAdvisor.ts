import type { OptionAdvisory } from '../../shared/types';
import { visibleText } from '../docx/xml';
import type { ParsedPaper } from '../parse/PaperModel';

const CATCH_ALL = /\b(none|all)\s+of\s+(these|the\s+above|them)\b/i;
const REFERENCES_OPTION = /\b(both|either|neither)\s*\(?[A-D]\)?\s*(?:and|or|&|,)\s*\(?[A-D]\)?/i;
const ASSERTION_REASON = /\bassertion\b[\s\S]*\breason\b|\bstatement[- ]?(i{1,3}|1|2)\b[\s\S]*\bstatement[- ]?(i{1,3}|1|2)\b/i;

/**
 * Flags questions whose option order carries meaning, so the user can decide whether
 * to add them to the "options not shuffled" exclusion list.
 *
 * This is advice only - it never changes what the generator does.
 */
export class OptionAdvisor {
  advise(paper: ParsedPaper): OptionAdvisory[] {
    const out: OptionAdvisory[] = [];

    for (const section of paper.sections) {
      for (const block of section.blocks) {
        const text = block.nodes.map((node) => visibleText(node)).join(' ');
        const add = (kind: OptionAdvisory['kind'], detail: string): void => {
          out.push({ questionNumber: block.printedNumber, subject: section.subject, kind, detail });
        };

        const referenceMatch = REFERENCES_OPTION.exec(text);
        if (referenceMatch) {
          add('references-other-option', `Option text refers to another option: "${referenceMatch[0].trim()}"`);
          continue;
        }
        const catchAllMatch = CATCH_ALL.exec(text);
        if (catchAllMatch) {
          add('catch-all-option', `Contains a catch-all option: "${catchAllMatch[0].trim()}"`);
          continue;
        }
        if (ASSERTION_REASON.test(text)) {
          add('assertion-reason', 'Assertion-Reason / Statement-I-II style question.');
        }
      }
    }

    return out;
  }
}
