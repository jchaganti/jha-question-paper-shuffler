import type { DocxPackage } from '../docx/DocxPackage';
import type { Element } from '../docx/dom';
import { NS, XmlPart, childElements, firstChild, hasDescendant, visibleText, wVal } from '../docx/xml';
import { AnswerKeyTable, isHeadingParagraph } from './AnswerKeyTable';
import { NumberingIndex } from './NumberingIndex';
import {
  DEFAULT_ANSWER_KEY_HEADINGS,
  DEFAULT_SUBJECT_HEADINGS,
  FALLBACK_SUBJECT,
  type ParsedPaper,
  type PaperSection,
  type ParserOptions,
  type QuestionBlock,
} from './PaperModel';

export class PaperParseError extends Error {}

/**
 * Turns `word/document.xml` into subjects -> questions -> body nodes.
 *
 * Assumptions (validated, and reported as errors when they do not hold):
 *  1. Question numbers are produced by Word's automatic numbering, not typed in.
 *     Reordering question blocks therefore renumbers the paper automatically.
 *  2. Subjects are introduced by a paragraph whose entire text is the subject name.
 *  3. The paper ends with an answer key whose cells pair a question number with a
 *     letter A-D, and the key covers every question exactly once.
 */
export class PaperParser {
  constructor(private readonly options: ParserOptions = {}) {}

  parse(pkg: DocxPackage): ParsedPaper {
    const part = pkg.documentPart();
    return this.parsePart(part, new NumberingIndex(pkg.numberingPart()));
  }

  parsePart(part: XmlPart, numbering: NumberingIndex): ParsedPaper {
    const body = firstChild(part.root, NS.w, 'body');
    if (!body) throw new PaperParseError('word/document.xml has no <w:body>.');

    const children = childElements(body);
    const subjectHeadings = this.options.subjectHeadings ?? DEFAULT_SUBJECT_HEADINGS;
    const keyHeadings = this.options.answerKeyHeadings ?? DEFAULT_ANSWER_KEY_HEADINGS;

    const answerKeyStart = findAnswerKeyStart(children, keyHeadings);
    const questionRegion = children.slice(0, answerKeyStart);
    const answerKeyNodes = children.slice(answerKeyStart);

    const answerKey = AnswerKeyTable.detect(answerKeyNodes) ?? AnswerKeyTable.detect(children);
    if (!answerKey) {
      throw new PaperParseError(
        'No answer key found. The tool expects a table at the end of the paper whose cells pair ' +
          'a question number with an answer - a letter (A-D), a digit (1-4) or a roman numeral (i-iv).',
      );
    }
    if (answerKey.duplicateNumbers.length > 0) {
      throw new PaperParseError(
        `The answer key lists question ${answerKey.duplicateNumbers.slice(0, 5).join(', ')} more than once. ` +
          'Papers where each subject restarts its numbering at 1 are not supported: the tool needs one ' +
          'continuous set of question numbers so that a key entry identifies exactly one question.',
      );
    }
    // A hole inside the key's own run of numbers is always a defect, and it makes the key
    // an undercount of the paper - so it is reported here, before that wrong total is used
    // to work out the numbering and produces an error about the wrong thing.
    const keyDefect = describeKeyGaps(answerKey);
    if (keyDefect) throw new PaperParseError(`The answer key is incomplete. ${keyDefect}`);
    const keyNumbers = answerKey.questionNumbers;

    const questionNumIds = resolveQuestionNumIds(questionRegion, numbering, keyNumbers.length);
    const isQuestionStart = (node: Element): boolean => {
      const numId = paragraphNumId(node);
      return numId !== undefined && questionNumIds.includes(numId);
    };

    const totalQuestions = questionRegion.filter(isQuestionStart).length;
    if (totalQuestions !== keyNumbers.length) {
      throw new PaperParseError(
        `Found ${totalQuestions} numbered questions but the answer key lists ${keyNumbers.length}. ` +
          'The paper structure is not recognised - check that every question uses the same automatic ' +
          'numbering list and that the answer key covers all questions.',
      );
    }

    // Subject boundaries.
    const headingIndexes: number[] = [];
    questionRegion.forEach((node, index) => {
      if (isHeadingParagraph(node, subjectHeadings)) headingIndexes.push(index);
    });

    const preambleNodes = questionRegion.slice(0, headingIndexes[0] ?? 0);
    const subjectRanges: { subject: string; start: number; end: number }[] = [];
    if (headingIndexes.length === 0) {
      subjectRanges.push({ subject: FALLBACK_SUBJECT, start: 0, end: questionRegion.length });
    } else {
      headingIndexes.forEach((start, i) => {
        const end = headingIndexes[i + 1] ?? questionRegion.length;
        subjectRanges.push({
          subject: visibleText(questionRegion[start]!).replace(/\s+/g, ' ').trim().toUpperCase(),
          start,
          end,
        });
      });
    }

    // A subject is divided further by its own "SECTION A" / "PART 2" headings, and each
    // division shuffles separately - see `splitIntoGroups`.
    const sectionRanges = subjectRanges.flatMap((range) =>
      splitIntoGroups(questionRegion, range, isQuestionStart),
    );

    let printedIndex = 0;
    const sections: PaperSection[] = sectionRanges.map((range, sectionIndex) => {
      const nodes = questionRegion.slice(range.start, range.end);
      const starts: number[] = [];
      nodes.forEach((node, i) => {
        if (isQuestionStart(node)) starts.push(i);
      });

      const headerNodes = nodes.slice(0, starts[0] ?? nodes.length);
      const blocks: QuestionBlock[] = starts.map((start, i) => {
        const end = starts[i + 1] ?? nodes.length;
        return {
          nodes: nodes.slice(start, end),
          questionParagraph: nodes[start]!,
          printedNumber: keyNumbers[printedIndex + i]!,
          sectionIndex,
          localIndex: i,
        };
      });
      printedIndex += blocks.length;

      // Blank paragraphs that close a subject are page spacing, not part of the last
      // question: they stay put so that the next subject still starts on a fresh page.
      let tailNodes: Element[] = [];
      const last = blocks[blocks.length - 1];
      if (last) {
        const trailing = countTrailingEmptyParagraphs(last.nodes);
        if (trailing > 0) {
          tailNodes = last.nodes.splice(last.nodes.length - trailing, trailing);
        }
      }

      // Only a real heading counts, and only on the division that opens the subject: the
      // "ALL" fallback section starts at the top of the paper, where the first paragraph is
      // the paper's title rather than a subject, and "SECTION B" does not start a new page.
      const headingNode = headingIndexes.length > 0 && range.opensSubject ? nodes[0] : undefined;
      return {
        subject: range.subject,
        group: range.group,
        label: groupLabel(range.subject, range.group),
        headingNode,
        headerNodes,
        blocks,
        tailNodes,
      };
    });

    return {
      part,
      body,
      preambleNodes,
      sections,
      answerKeyNodes,
      answerKey,
      questionCount: totalQuestions,
      questionNumIds,
    };
  }
}

/**
 * A paragraph that divides a subject: "SECTION A (All questions are compulsory)",
 * "SECTION - B", "PART II", "BIOLOGY PART 1".
 *
 * Read rather than assumed, so a paper that writes its divisions differently still works:
 * an optional leading word (papers repeat the subject - "BIOLOGY PART 1"), the word PART or
 * SECTION, an optional dash or colon, and the division's name - a letter, a number or a
 * roman numeral. What follows may only be a parenthesised note, which is where papers put
 * "(Attempt any 10 questions)".
 *
 * Requiring the whole paragraph to be that shape is what keeps prose out: a question
 * mentioning "cross-section" has words either side, and a question stem is auto-numbered,
 * which is checked separately.
 */
const DIVIDER_RE =
  /^(?:[A-Z]+\s+)?(PART|SECTION)\s*[-‐-―:.]?\s*([A-Z]|\d{1,2}|[IVX]{1,4})\s*(?:\([^)]*\))?$/;

/**
 * Splits one subject at its own division headings, so that questions shuffle within a
 * division and never across one.
 *
 * The divisions are not merely labels. "SECTION A (All questions are compulsory)" and
 * "SECTION B (Attempt any 10 questions)" ask different things of the candidate, and moving
 * a question between them would change the paper. A subject with no such headings comes
 * back as one group, exactly as before.
 *
 * PART headings nest above SECTION headings, so the two are tracked separately and the
 * label carries both: "PART 2 SECTION A". A heading with no questions before the next one -
 * "BIOLOGY PART 1" immediately above "SECTION - A" - does not open a group of its own; it
 * joins the heading below it, which is the group the questions are actually in.
 */
function splitIntoGroups(
  questionRegion: readonly Element[],
  range: { subject: string; start: number; end: number },
  isQuestionStart: (node: Element) => boolean,
): { subject: string; group?: string; start: number; end: number; opensSubject: boolean }[] {
  const cuts: { start: number; group?: string }[] = [{ start: range.start }];
  let part: string | undefined;
  let section: string | undefined;

  for (let index = range.start + 1; index < range.end; index++) {
    const node = questionRegion[index]!;
    const heading = dividerHeading(node);
    if (!heading) continue;

    if (heading.keyword === 'PART') {
      part = `PART ${heading.name}`;
      // A new part restarts its sections, so "PART 2" alone must not keep Part 1's section.
      section = undefined;
    } else {
      section = `SECTION ${heading.name}`;
    }
    const group = [part, section].filter((value) => value !== undefined).join(' ');

    const open = cuts[cuts.length - 1]!;
    const hasQuestions = questionRegion.slice(open.start, index).some(isQuestionStart);
    if (hasQuestions) cuts.push({ start: index, group });
    else open.group = group;
  }

  return cuts.map((cut, i) => ({
    subject: range.subject,
    group: cut.group,
    start: cut.start,
    end: cuts[i + 1]?.start ?? range.end,
    opensSubject: i === 0,
  }));
}

/**
 * How one run of questions is named to the user.
 *
 * "ALL" is the stand-in for a paper whose subject headings were not recognised, so it is
 * left out of the name: a single-subject paper divided into sections reads "SECTION A",
 * not "ALL - SECTION A".
 */
function groupLabel(subject: string, group: string | undefined): string {
  if (group === undefined) return subject;
  return subject === FALLBACK_SUBJECT ? group : `${subject} - ${group}`;
}

/** The keyword and name of a division heading, or undefined when the paragraph is not one. */
function dividerHeading(node: Element): { keyword: string; name: string } | undefined {
  if (node.namespaceURI !== NS.w || node.localName !== 'p') return undefined;
  // A question stem is auto-numbered; a heading never is.
  if (paragraphNumId(node) !== undefined) return undefined;
  const text = visibleText(node).replace(/\s+/g, ' ').trim().toUpperCase();
  const match = DIVIDER_RE.exec(text);
  return match ? { keyword: match[1]!, name: match[2]! } : undefined;
}

/**
 * Guards against a paper whose questions were split over several numbering lists that do
 * not chain (for example each subject restarting at 1). Short decimal lists are ordinary
 * numbered statements inside a question, so only a long unused list is treated as a sign
 * that the paper is not laid out the way this tool expects.
 */
const SUSPICIOUS_UNUSED_LIST_SIZE = 10;

function assertNoUnusedQuestionList(
  chosen: readonly string[],
  counts: ReadonlyMap<string, number>,
  numbering: NumberingIndex,
): void {
  const suspicious = [...counts.entries()].filter(
    ([numId, count]) =>
      !chosen.includes(numId) &&
      count >= SUSPICIOUS_UNUSED_LIST_SIZE &&
      numbering.isDecimalList(numId) &&
      numbering.get(numId)?.start === 1,
  );
  if (suspicious.length === 0) return;

  throw new PaperParseError(
    'This paper numbers its questions with more than one list, and the lists do not continue ' +
      'from one another' +
      suspicious.map(([numId, count]) => ` (list ${numId} restarts at 1 with ${count} items)`).join(',') +
      '. The tool needs one continuous numbering across the paper so that a question number ' +
      'identifies exactly one question. Nothing has been generated.',
  );
}

export function paragraphNumId(node: Element): string | undefined {
  if (node.namespaceURI !== NS.w || node.localName !== 'p') return undefined;
  const pPr = firstChild(node, NS.w, 'pPr');
  if (!pPr) return undefined;
  const numPr = firstChild(pPr, NS.w, 'numPr');
  if (!numPr) return undefined;
  const numId = firstChild(numPr, NS.w, 'numId');
  return numId ? wVal(numId) : undefined;
}

export function isEmptyParagraph(node: Element): boolean {
  if (node.namespaceURI !== NS.w || node.localName !== 'p') return false;
  if (visibleText(node).trim() !== '') return false;
  return (
    !hasDescendant(node, NS.w, 'drawing') &&
    !hasDescendant(node, NS.w, 'object') &&
    !hasDescendant(node, NS.w, 'pict')
  );
}

function countTrailingEmptyParagraphs(nodes: readonly Element[]): number {
  let count = 0;
  for (let i = nodes.length - 1; i > 0; i--) {
    if (!isEmptyParagraph(nodes[i]!)) break;
    count++;
  }
  return count;
}

const list = (items: readonly (string | number)[]): string =>
  items.length <= 2 ? items.join(' and ') : `${items.slice(0, -1).join(', ')} and ${items[items.length - 1]}`;

/**
 * Turns "the key is one entry short" into the box the author has to correct. A gap in the
 * key's own run of numbers names the question; a spoilt number box next to an answer
 * letter - "57," for "57." - names the mistake itself.
 */
function describeKeyGaps(answerKey: AnswerKeyTable): string | undefined {
  const gaps = answerKey.gaps;
  const spoilt = answerKey.spoiltNumbers;
  if (gaps.length === 0 && spoilt.length === 0) return undefined;

  const parts: string[] = [];
  if (gaps.length > 0) {
    const shown = gaps.slice(0, 5);
    parts.push(
      `The key has no entry for question ${list(shown)}` +
        (gaps.length > shown.length ? ` (and ${gaps.length - shown.length} more)` : '') +
        '.',
    );
  }
  if (spoilt.length > 0) {
    const shown = spoilt.slice(0, 5).map((text) => `"${text}"`);
    const one = spoilt.length === 1;
    parts.push(
      (one ? 'A number box in the key reads ' : 'Number boxes in the key read ') +
        list(shown) +
        (spoilt.length > shown.length ? ` (and ${spoilt.length - shown.length} more)` : '') +
        ' - a stray character was typed after the number. Retype ' +
        (one ? 'that box' : 'those boxes') +
        ' as the plain number.',
    );
  } else {
    parts.push('Add the missing entry to the key, or check that box for a stray character.');
  }
  return parts.join(' ');
}

function findAnswerKeyStart(children: readonly Element[], headings: readonly string[]): number {
  for (let i = 0; i < children.length; i++) {
    if (isHeadingParagraph(children[i]!, headings)) return i;
  }
  // Fall back to the first table that looks like an answer key.
  for (let i = 0; i < children.length; i++) {
    const node = children[i]!;
    if (node.namespaceURI === NS.w && node.localName === 'tbl' && AnswerKeyTable.detect([node])) return i;
  }
  return children.length;
}

/**
 * Works out which automatic numbering lists carry question numbers.
 *
 * A paper may use several lists (this sample uses one for Physics starting at 1 and
 * another for Chemistry + Biology starting at 46), while option lists and statement
 * lists use their own numbering. Lists are therefore accepted in document order only
 * when they are decimal *and* start exactly where the previous list stopped.
 */
export function resolveQuestionNumIds(
  questionRegion: readonly Element[],
  numbering: NumberingIndex,
  expectedTotal: number,
): string[] {
  const numbered = questionRegion
    .map((node) => paragraphNumId(node))
    .filter((numId): numId is string => numId !== undefined);

  const counts = new Map<string, number>();
  for (const numId of numbered) counts.set(numId, (counts.get(numId) ?? 0) + 1);

  const chosen: string[] = [];
  let expectedNext = 1;
  for (const numId of numbered) {
    if (chosen.includes(numId)) continue;
    const def = numbering.get(numId);
    if (!def || def.format !== 'decimal') continue;
    if (def.start !== expectedNext) continue;
    chosen.push(numId);
    expectedNext += counts.get(numId) ?? 0;
  }

  const total = chosen.reduce((sum, numId) => sum + (counts.get(numId) ?? 0), 0);
  if (total === expectedTotal) {
    assertNoUnusedQuestionList(chosen, counts, numbering);
    return chosen;
  }

  // Fallback: accept decimal lists in document order until the counts add up.
  const fallback: string[] = [];
  let running = 0;
  for (const numId of numbered) {
    if (fallback.includes(numId)) continue;
    if (!numbering.isDecimalList(numId)) continue;
    fallback.push(numId);
    running += counts.get(numId) ?? 0;
    if (running === expectedTotal) {
      assertNoUnusedQuestionList(fallback, counts, numbering);
      return fallback;
    }
    if (running > expectedTotal) break;
  }

  throw new PaperParseError(
    `Could not identify the question numbering. This usually means the question numbers were ` +
      `typed by hand instead of using Word's automatic numbering. Decimal numbering lists found: ` +
      [...counts.entries()]
        .filter(([numId]) => numbering.isDecimalList(numId))
        .map(([numId, count]) => `numId ${numId} (start ${numbering.get(numId)?.start ?? '?'}, ${count} items)`)
        .join('; ') +
      `. Expected ${expectedTotal} questions in total.`,
  );
}
