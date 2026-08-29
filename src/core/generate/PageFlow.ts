import type { Element } from '../docx/dom';
import { NS, childElements, createElement, descendants, firstChild, ownerDocumentOf } from '../docx/xml';
import type { ParsedPaper, QuestionBlock } from '../parse/PaperModel';
import { isEmptyParagraph } from '../parse/PaperParser';

/**
 * Keeps every question whole on one page.
 *
 * Where the page breaks fall is decided by Word at layout time, from font metrics and
 * image sizes this tool cannot measure. So nothing is measured: instead each question is
 * marked with the two flags Word itself uses for exactly this purpose.
 *
 *  - `w:keepNext`  - "do not put a page break between me and the paragraph after me".
 *    Set on every paragraph of the question except its last one, which chains the whole
 *    question together.
 *  - `w:keepLines` - "do not split my own lines over two pages". Needed as well, because
 *    keepNext only governs the gaps *between* paragraphs; a single long option or stem
 *    could still be cut in half.
 *
 * The effect is the one asked for: a question that no longer fits at the bottom of a page
 * moves down as a block and starts at the top of the next page.
 *
 * Assumptions and limits (deliberate, and safe):
 *  1. A question taller than one page cannot be kept together. Word drops the request in
 *     that case and breaks the question as before - it never loses content.
 *  2. Blank paragraphs that trail a question are page spacing, not content, so they are
 *     left out of the chain. The page break is then free to fall in that gap, which is
 *     what keeps the paper from growing more than it must.
 *  3. Flags already present in the source document are left exactly as they are.
 *  4. If the paper is laid out in columns, "page" means "column" - the question moves to
 *     the top of the next column, which is the same guarantee for a reader.
 */
export class PageFlowGuard {
  /** Marks every question in the paper. Returns the number of questions touched. */
  keepQuestionsWhole(paper: ParsedPaper): number {
    let count = 0;
    for (const section of paper.sections) {
      for (const block of section.blocks) {
        if (this.keepBlockWhole(block)) count++;
      }
    }
    return count;
  }

  /**
   * Starts the answer key on a fresh page.
   *
   * Papers normally push the key onto its own page with a run of blank paragraphs, which
   * stops working as soon as the questions are re-ordered and the text reflows - one
   * sample paper has no blank paragraphs at all, so its key ran on straight after the last
   * option. `w:pageBreakBefore` states the intent instead of approximating it.
   *
   * Left alone when a page break is already there (an explicit break run, a next-page
   * section break, or the property itself), so no paper ever gains a blank page.
   *
   * Returns true when the property was added.
   */
  answerKeyOnNewPage(paper: ParsedPaper): boolean {
    const body = childElements(paper.body);
    const index = body.length - paper.answerKeyNodes.length;
    const first = paper.answerKeyNodes[0];
    if (!first) return false;

    // The key normally opens with its "ANSWER KEY" heading; when a paper has no heading
    // it opens with the table, and the break goes on the first paragraph inside it.
    const target = isWordElement(first, 'p') ? first : descendants(first, NS.w, 'p')[0];
    return target ? startOnNewPage(body, index, target) : false;
  }

  /**
   * Starts each subject on a fresh page, by the same rule as the answer key: papers
   * separate their subjects with blank paragraphs, which stops working once the questions
   * are re-ordered and the text reflows.
   *
   * Skipped for a subject that already opens a page, and for the first subject when it is
   * the very start of the document - there is nothing to break away from, and asking for a
   * break there is how a document gains a leading blank page.
   *
   * A paper whose subjects the tool did not recognise has no subject headings to mark, so
   * nothing happens (its one section covers the whole paper).
   *
   * Returns how many subjects were marked.
   */
  subjectsOnNewPage(paper: ParsedPaper): number {
    const body = childElements(paper.body);
    let count = 0;
    for (const section of paper.sections) {
      const heading = section.headingNode;
      if (!heading) continue;
      const index = body.indexOf(heading);
      if (index <= 0) continue;
      if (startOnNewPage(body, index, heading)) count++;
    }
    return count;
  }

  /**
   * Marks one question. Returns false only for a question with nothing to mark, which
   * cannot normally happen (a question always has at least its stem paragraph).
   */
  keepBlockWhole(block: QuestionBlock): boolean {
    const paragraphs: Element[] = [];
    for (const node of block.nodes) {
      collectParagraphs(node, paragraphs);
      // A table row splits over a page break on its own terms, regardless of keepNext.
      if (isWordElement(node, 'tbl')) {
        for (const row of descendants(node, NS.w, 'tr')) {
          ensureFlag(ensureTrPr(row), 'cantSplit', TRPR_ORDER);
        }
      }
    }

    // Walk back over trailing blank paragraphs (assumption 2 above).
    let last = paragraphs.length - 1;
    while (last > 0 && isEmptyParagraph(paragraphs[last]!)) last--;
    if (last < 0) return false;

    for (let i = 0; i <= last; i++) {
      const pPr = ensurePPr(paragraphs[i]!);
      ensureFlag(pPr, 'keepLines', PPR_ORDER);
      if (i < last) ensureFlag(pPr, 'keepNext', PPR_ORDER);
    }
    return true;
  }
}

/**
 * `w:pPr` children in the order the OOXML schema (CT_PPr) demands. Word rejects a
 * document whose paragraph properties are out of order, so a new flag has to be spliced
 * into the right place rather than appended.
 */
const PPR_ORDER: readonly string[] = [
  'pStyle',
  'keepNext',
  'keepLines',
  'pageBreakBefore',
  'framePr',
  'widowControl',
  'numPr',
  'suppressLineNumbers',
  'pBdr',
  'shd',
  'tabs',
  'suppressAutoHyphens',
  'kinsoku',
  'wordWrap',
  'overflowPunct',
  'topLinePunct',
  'autoSpaceDE',
  'autoSpaceDN',
  'bidi',
  'adjustRightInd',
  'snapToGrid',
  'spacing',
  'ind',
  'contextualSpacing',
  'mirrorIndents',
  'suppressOverlap',
  'jc',
  'textDirection',
  'textAlignment',
  'textboxTightWrap',
  'outlineLvl',
  'divId',
  'cnfStyle',
  'rPr',
  'sectPr',
  'pPrChange',
];

/** `w:trPr` children in schema order (CT_TrPr). */
const TRPR_ORDER: readonly string[] = [
  'cnfStyle',
  'divId',
  'gridBefore',
  'gridAfter',
  'wBefore',
  'wAfter',
  'cantSplit',
  'trHeight',
  'tblHeader',
  'tblCellSpacing',
  'jc',
  'hidden',
  'ins',
  'del',
  'trPrChange',
];

function isWordElement(node: Element, localName: string): boolean {
  return node.namespaceURI === NS.w && node.localName === localName;
}

/**
 * Asks Word to start `target` at the top of a page, unless a page already ends in front of
 * it. Returns true when the property was added.
 */
function startOnNewPage(body: readonly Element[], index: number, target: Element): boolean {
  const existing = firstChild(target, NS.w, 'pPr');
  if (existing && firstChild(existing, NS.w, 'pageBreakBefore')) return false;
  if (breakAlreadyPrecedes(body, index)) return false;

  ensureFlag(ensurePPr(target), 'pageBreakBefore', PPR_ORDER);
  return true;
}

/**
 * True when a page already ends between the last real content before `index` and `index`
 * itself. Blank paragraphs are walked over - one of them may be the paragraph that *holds*
 * the break run, and a blank paragraph never ends a page by itself.
 */
function breakAlreadyPrecedes(body: readonly Element[], index: number): boolean {
  for (let i = index - 1; i >= 0; i--) {
    const node = body[i]!;
    if (hasExplicitPageBreak(node)) return true;
    const pPr = isWordElement(node, 'p') ? firstChild(node, NS.w, 'pPr') : undefined;
    // A section break ends the page too (the default section type is "next page").
    if (pPr && firstChild(pPr, NS.w, 'sectPr')) return true;
    if (!isWordElement(node, 'p') || !isEmptyParagraph(node)) return false;
  }
  return false;
}

/** A typed page break (Ctrl+Enter). `w:lastRenderedPageBreak` is a hint, not a break. */
function hasExplicitPageBreak(node: Element): boolean {
  return descendants(node, NS.w, 'br').some((br) => br.getAttributeNS(NS.w, 'type') === 'page');
}

/** Every paragraph of a body node, in reading order - including those inside tables. */
function collectParagraphs(node: Element, out: Element[]): void {
  if (isWordElement(node, 'p')) {
    out.push(node);
    return;
  }
  out.push(...descendants(node, NS.w, 'p'));
}

/** `w:pPr` must be the first child of `w:p`. */
function ensurePPr(paragraph: Element): Element {
  const existing = firstChild(paragraph, NS.w, 'pPr');
  if (existing) return existing;
  const pPr = createElement(ownerDocumentOf(paragraph), 'w:pPr', NS.w);
  paragraph.insertBefore(pPr, paragraph.firstChild);
  return pPr;
}

/** `w:trPr` follows an optional `w:tblPrEx` and precedes the cells. */
function ensureTrPr(row: Element): Element {
  const existing = firstChild(row, NS.w, 'trPr');
  if (existing) return existing;
  const trPr = createElement(ownerDocumentOf(row), 'w:trPr', NS.w);
  const successor = childElements(row).find((child) => !isWordElement(child, 'tblPrEx'));
  row.insertBefore(trPr, successor ?? null);
  return trPr;
}

/**
 * Adds a boolean property (an empty element means "on") if it is not there already,
 * inserted before the first sibling that the schema orders after it.
 */
function ensureFlag(props: Element, localName: string, order: readonly string[]): void {
  if (firstChild(props, NS.w, localName)) return;
  const rank = order.indexOf(localName);
  const successor = childElements(props).find((child) => {
    if (child.namespaceURI !== NS.w) return false;
    return order.indexOf(child.localName ?? '') > rank;
  });
  props.insertBefore(createElement(ownerDocumentOf(props), `w:${localName}`, NS.w), successor ?? null);
}
