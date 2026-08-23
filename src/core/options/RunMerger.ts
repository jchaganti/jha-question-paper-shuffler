import type { Element, Node } from '../docx/dom';
import { NS, childElements, firstChild, isElement, serializeElement } from '../docx/xml';

const MERGEABLE_CHILDREN = ['t', 'tab', 'br', 'sym', 'noBreakHyphen'];

/**
 * Merges neighbouring runs that share identical run properties.
 *
 * Shuffling works one inline element at a time, which would otherwise leave a paragraph
 * with one run per character group. Word tolerates that, but merging keeps the generated
 * XML close in size and shape to the original.
 */
export function mergeAdjacentRuns(nodes: readonly Node[]): Node[] {
  const out: Node[] = [];

  for (const node of nodes) {
    const previous = out[out.length - 1];
    if (
      isElement(node) &&
      isSimpleRun(node) &&
      previous &&
      isElement(previous) &&
      isSimpleRun(previous) &&
      runSignature(previous) === runSignature(node)
    ) {
      for (const child of childElements(node)) {
        if (child.namespaceURI === NS.w && child.localName === 'rPr') continue;
        previous.appendChild(child);
      }
      continue;
    }
    out.push(node);
  }

  return out;
}

function isSimpleRun(el: Element): boolean {
  if (el.namespaceURI !== NS.w || el.localName !== 'r') return false;
  return childElements(el).every(
    (child) =>
      child.namespaceURI === NS.w &&
      (child.localName === 'rPr' || MERGEABLE_CHILDREN.includes(child.localName ?? '')),
  );
}

function runSignature(run: Element): string {
  const attributes: string[] = [];
  for (let i = 0; i < run.attributes.length; i++) {
    const a = run.attributes.item(i);
    if (a) attributes.push(`${a.name}=${a.value}`);
  }
  const rPr = firstChild(run, NS.w, 'rPr');
  return `${attributes.sort().join('|')}#${rPr ? serializeElement(rPr) : ''}`;
}
