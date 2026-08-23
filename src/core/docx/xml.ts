import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import type { Document, Element, Node } from './dom';

/** OOXML namespaces used by this application. */
export const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  xml: 'http://www.w3.org/XML/1998/namespace',
} as const;

const XML_PROLOG_RE = /^\s*<\?xml[^>]*\?>\s*/;
const DEFAULT_PROLOG = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

/**
 * A parsed OOXML part. `@xmldom/xmldom` drops the XML declaration when parsing, so the
 * original prolog is captured and re-emitted on serialisation to keep the part as close
 * to what Word writes as possible.
 */
export class XmlPart {
  private constructor(
    readonly doc: Document,
    private readonly prolog: string,
  ) {}

  static parse(xml: string): XmlPart {
    const match = XML_PROLOG_RE.exec(xml);
    const prolog = match ? match[0] : DEFAULT_PROLOG;
    const doc = new DOMParser().parseFromString(xml, 'text/xml');
    return new XmlPart(doc, prolog);
  }

  get root(): Element {
    const el = this.doc.documentElement;
    if (!el) throw new Error('XML part has no document element');
    return el;
  }

  serialize(): string {
    const xml = new XMLSerializer().serializeToString(this.doc);
    // Some parser versions keep the declaration as a processing instruction; only add
    // the captured prolog back when it was really dropped.
    return XML_PROLOG_RE.test(xml) ? xml : this.prolog + xml;
  }
}

export function isElement(node: Node | null | undefined): node is Element {
  return !!node && node.nodeType === 1;
}

/** Direct element children of `parent`, optionally filtered by namespace + local name. */
export function childElements(parent: Element, ns?: string, localName?: string): Element[] {
  const out: Element[] = [];
  for (let i = 0; i < parent.childNodes.length; i++) {
    const node = parent.childNodes.item(i);
    if (!isElement(node)) continue;
    if (ns !== undefined && node.namespaceURI !== ns) continue;
    if (localName !== undefined && node.localName !== localName) continue;
    out.push(node);
  }
  return out;
}

/** First matching direct child, or undefined. */
export function firstChild(parent: Element, ns: string, localName: string): Element | undefined {
  return childElements(parent, ns, localName)[0];
}

/** All descendants (excluding `root` itself) matching namespace + local name. */
export function descendants(root: Element, ns: string, localName: string): Element[] {
  const out: Element[] = [];
  const walk = (el: Element): void => {
    for (let i = 0; i < el.childNodes.length; i++) {
      const node = el.childNodes.item(i);
      if (!isElement(node)) continue;
      if (node.namespaceURI === ns && node.localName === localName) out.push(node);
      walk(node);
    }
  };
  walk(root);
  return out;
}

/** True when `root` has at least one descendant with the given namespace + local name. */
export function hasDescendant(root: Element, ns: string, localName: string): boolean {
  for (let i = 0; i < root.childNodes.length; i++) {
    const node = root.childNodes.item(i);
    if (!isElement(node)) continue;
    if (node.namespaceURI === ns && node.localName === localName) return true;
    if (hasDescendant(node, ns, localName)) return true;
  }
  return false;
}

export function wVal(el: Element): string | undefined {
  const value = el.getAttributeNS(NS.w, 'val');
  return value === null || value === '' ? undefined : value;
}

/** Visible text of a paragraph/run subtree: `w:t` text plus tabs and line breaks. */
export function visibleText(root: Element): string {
  let out = '';
  const walk = (el: Element): void => {
    for (let i = 0; i < el.childNodes.length; i++) {
      const node = el.childNodes.item(i);
      if (!isElement(node)) continue;
      if (node.namespaceURI === NS.w) {
        switch (node.localName) {
          case 't':
          case 'delText':
            out += node.textContent ?? '';
            continue;
          case 'tab':
            out += '\t';
            continue;
          case 'br':
            out += '\n';
            continue;
          case 'noBreakHyphen':
            out += '-';
            continue;
          default:
            break;
        }
      }
      walk(node);
    }
  };
  walk(root);
  return out;
}

export function createElement(doc: Document, qualifiedName: string, ns: string): Element {
  return doc.createElementNS(ns, qualifiedName);
}

/** Sets `xml:space="preserve"` so that leading/trailing spaces survive a round trip. */
export function preserveSpace(el: Element): void {
  el.setAttributeNS(NS.xml, 'xml:space', 'preserve');
}

export function replaceChildren(el: Element, nodes: readonly Node[]): void {
  while (el.firstChild) el.removeChild(el.firstChild);
  for (const node of nodes) el.appendChild(node);
}

export function serializeElement(el: Element): string {
  return new XMLSerializer().serializeToString(el);
}

export function ownerDocumentOf(el: Element): Document {
  const doc = el.ownerDocument;
  if (!doc) throw new Error('Element is not attached to a document');
  return doc;
}
