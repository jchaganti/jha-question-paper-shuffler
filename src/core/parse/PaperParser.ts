import type { DocxPackage } from '../docx/DocxPackage';
import type { Element } from '../docx/dom';
import { NS, XmlPart, childElements, firstChild, hasDescendant, visibleText, wVal } from '../docx/xml';
import { AnswerKeyTable, isHeadingParagraph } from './AnswerKeyTable';
import { NumberingIndex } from './NumberingIndex';
import {
  DEFAULT_ANSWER_KEY_HEADINGS,
  DEFAULT_SUBJECT_HEADINGS,
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
          'a question number with an answer letter (A-D).',
      );
    }
    if (answerKey.duplicateNumbers.length > 0) {
      throw new PaperParseError(
        `The answer key lists question ${answerKey.duplicateNumbers.slice(0, 5).join(', ')} more than once. ` +
          'Papers where each subject restarts its numbering at 1 are not supported: the tool needs one ' +
          'continuous set of question numbers so that a key entry identifies exactly one question.',
      );
    }
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
    const sectionRanges: { subject: string; start: number; end: number }[] = [];
    if (headingIndexes.length === 0) {
      sectionRanges.push({ subject: 'ALL', start: 0, end: questionRegion.length });
    } else {
      headingIndexes.forEach((start, i) => {
        const end = headingIndexes[i + 1] ?? questionRegion.length;
        sectionRanges.push({
          subject: visibleText(questionRegion[start]!).replace(/\s+/g, ' ').trim().toUpperCase(),
          start,
          end,
        });
      });
    }

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

      // Only a real heading counts: the "ALL" fallback section starts at the top of the
      // paper, where the first paragraph is the paper's title, not a subject.
      const headingNode = headingIndexes.length > 0 ? nodes[0] : undefined;
      return { subject: range.subject, headingNode, headerNodes, blocks, tailNodes };
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
