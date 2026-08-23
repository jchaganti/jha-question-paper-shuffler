import type { Element } from '../docx/dom';
import { NS, isElement, visibleText } from '../docx/xml';
import type { IOptionSetParser } from '../options/OptionSet';
import type { QuestionBlock } from '../parse/PaperModel';

/**
 * Stable fingerprint of a question, used after generation to prove that the question
 * which landed at a position is the one the plan intended.
 *
 * The fingerprint must be *invariant under option shuffling* (otherwise a shuffled
 * question would look like a different question), yet still distinguish two questions
 * that share a stem - this paper has two questions reading "Select the correct
 * statement." So it combines the stem with the *sorted* set of option contents.
 */
export function questionSignature(block: QuestionBlock, parser: IOptionSetParser): string {
  const stem = normalise(visibleText(block.questionParagraph));
  const ids = relationshipIds(block.questionParagraph);
  const tableText = block.nodes
    .filter((node) => node.namespaceURI === NS.w && node.localName === 'tbl')
    .map((table) => normalise(visibleText(table)))
    .join('~');

  const parsed = parser.parse(block);
  const body = parsed.ok
    ? [...parsed.options.signatures].sort()
    : block.nodes
        .filter((node) => node !== block.questionParagraph)
        .map((node) => normalise(visibleText(node)))
        .filter((text) => text !== '')
        .sort();

  return `${stem}|${ids.join(',')}|${tableText}|${body.join('~')}`;
}

export function normalise(text: string): string {
  return text.replace(/\s+/g, ' ').trim().toLowerCase();
}

export function relationshipIds(root: Element): string[] {
  const out: string[] = [];
  const walk = (el: Element): void => {
    for (let i = 0; i < el.attributes.length; i++) {
      const a = el.attributes.item(i);
      if (a && a.namespaceURI === NS.r) out.push(`${a.localName}=${a.value}`);
    }
    for (let i = 0; i < el.childNodes.length; i++) {
      const node = el.childNodes.item(i);
      if (isElement(node)) walk(node);
    }
  };
  walk(root);
  return out.sort();
}
