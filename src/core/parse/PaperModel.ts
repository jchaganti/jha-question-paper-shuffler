import type { Element } from '../docx/dom';
import type { XmlPart } from '../docx/xml';
import type { AnswerKeyTable } from './AnswerKeyTable';

/** One question with everything that belongs to it, as a slice of body nodes. */
export interface QuestionBlock {
  /** Body-level nodes of this question: the numbered paragraph, its options, tables, images. */
  readonly nodes: Element[];
  /** The auto-numbered paragraph that carries the question stem. */
  readonly questionParagraph: Element;
  /** Question number as printed in the *original* paper. */
  readonly printedNumber: number;
  /** Index of the `PaperSection` - the subject division - this question sits in. */
  readonly sectionIndex: number;
  /** 0-based position inside that division. */
  readonly localIndex: number;
}

/**
 * One run of questions that may be re-ordered among themselves, and the smallest unit the
 * shuffler works in.
 *
 * Usually that is a subject. But a paper often divides a subject further - "SECTION A (All
 * questions are compulsory)" and "SECTION B (Attempt any 10 questions)", sometimes with
 * "PART 1" / "PART 2" above those. Those divisions carry rules of their own, so a question
 * must never cross one: a compulsory question moved into the attempt-any-10 section changes
 * what the candidate is asked to do. Each division is therefore its own `PaperSection`.
 */
export interface PaperSection {
  /** The subject these questions belong to: "PHYSICS". Shared by all its divisions. */
  readonly subject: string;
  /**
   * The division within the subject, when the paper has one: "SECTION A",
   * "PART 2 SECTION A". Absent when the subject is undivided.
   */
  readonly group?: string;
  /** How this run of questions is named to the user: "PHYSICS" or "PHYSICS - SECTION B". */
  readonly label: string;
  /**
   * The paragraph that announces the *subject*, when the paper has one - carried by the
   * first division only, since that is the one that starts a new page. Absent when no
   * subject heading was recognised and the whole paper is treated as a single section -
   * in that case `headerNodes` opens with the paper's own title, which is not a subject.
   */
  readonly headingNode?: Element;
  /** Subject and division headings plus anything before the first question (kept in place). */
  readonly headerNodes: Element[];
  readonly blocks: QuestionBlock[];
  /** Blank spacer paragraphs that close the subject (kept in place). */
  readonly tailNodes: Element[];
}

export interface ParsedPaper {
  readonly part: XmlPart;
  readonly body: Element;
  /** Nodes before the first subject heading (kept in place). */
  readonly preambleNodes: Element[];
  readonly sections: PaperSection[];
  /** "ANSWER KEY" heading, the key table(s) and whatever follows (kept in place). */
  readonly answerKeyNodes: Element[];
  readonly answerKey: AnswerKeyTable;
  readonly questionCount: number;
  /** numIds recognised as question numbering. */
  readonly questionNumIds: readonly string[];
}

/**
 * Stand-in subject for a paper whose subject headings were not recognised - a single-subject
 * paper, usually. Its sections still shuffle separately; only the subject name is missing.
 */
export const FALLBACK_SUBJECT = 'ALL';

export const DEFAULT_SUBJECT_HEADINGS = [
  'PHYSICS',
  'CHEMISTRY',
  'BIOLOGY',
  'BOTANY',
  'ZOOLOGY',
  'MATHEMATICS',
  'MATHS',
] as const;

export const DEFAULT_ANSWER_KEY_HEADINGS = ['ANSWER KEY', 'ANSWER-KEY', 'ANSWERS', 'ANSWER KEYS'] as const;

export interface ParserOptions {
  readonly subjectHeadings?: readonly string[];
  readonly answerKeyHeadings?: readonly string[];
}
