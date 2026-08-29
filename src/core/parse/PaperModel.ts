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
  readonly sectionIndex: number;
  /** 0-based position inside the subject. */
  readonly localIndex: number;
}

export interface PaperSection {
  readonly subject: string;
  /**
   * The paragraph that announces the subject, when the paper has one. Absent when no
   * subject heading was recognised and the whole paper is treated as a single section -
   * in that case `headerNodes` opens with the paper's own title, which is not a subject.
   */
  readonly headingNode?: Element;
  /** Subject heading plus anything before the first question (kept in place). */
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
