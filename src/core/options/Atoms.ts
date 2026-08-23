import type { Element } from '../docx/dom';
import {
  NS,
  childElements,
  createElement,
  descendants,
  firstChild,
  hasDescendant,
  isElement,
  ownerDocumentOf,
  preserveSpace,
  visibleText,
} from '../docx/xml';

/**
 * The smallest movable piece of a paragraph.
 *
 * Option text is not stored one-option-per-run: a single run can hold a tab plus the
 * next option's label, and an option's value can be a MathType/OLE object. Shuffling
 * options therefore works on "atoms" - one inline element wrapped in its own run,
 * carrying the original run properties - so that fonts, superscripts, symbols and
 * equations survive being moved.
 */
export interface Atom {
  /** Element to place inside the paragraph. Always safe to move as a whole. */
  readonly node: Element;
  /** Index of the source paragraph inside the option region. */
  readonly paragraphIndex: number;
  /** Visible text contributed by this atom ('\t' for a tab). */
  readonly text: string;
  /** True when the atom carries no visible ink (spaces, tabs, bookmarks). */
  readonly blank: boolean;
  /** True when the atom holds a floating (anchored) image, which must not be moved. */
  readonly floatingGraphic: boolean;
}

export interface ParagraphAtoms {
  readonly paragraph: Element;
  readonly paragraphIndex: number;
  readonly atoms: Atom[];
}

/** Splits every paragraph of the option region into atoms. */
export function buildParagraphAtoms(paragraphs: readonly Element[]): ParagraphAtoms[] {
  return paragraphs.map((paragraph, paragraphIndex) => ({
    paragraph,
    paragraphIndex,
    atoms: atomsOfParagraph(paragraph, paragraphIndex),
  }));
}

function atomsOfParagraph(paragraph: Element, paragraphIndex: number): Atom[] {
  const doc = ownerDocumentOf(paragraph);
  const atoms: Atom[] = [];

  for (const child of childElements(paragraph)) {
    if (child.namespaceURI === NS.w && child.localName === 'pPr') continue;

    if (child.namespaceURI === NS.w && child.localName === 'r') {
      const rPr = firstChild(child, NS.w, 'rPr');
      for (const inline of childElements(child)) {
        if (inline.namespaceURI === NS.w && inline.localName === 'rPr') continue;
        const run = createElement(doc, 'w:r', NS.w);
        copyAttributes(child, run);
        if (rPr) run.appendChild(rPr.cloneNode(true));
        run.appendChild(inline.cloneNode(true));
        atoms.push(makeAtom(run, inline, paragraphIndex));
      }
      continue;
    }

    atoms.push(makeAtom(child.cloneNode(true) as Element, child, paragraphIndex));
  }

  return atoms;
}

function makeAtom(node: Element, source: Element, paragraphIndex: number): Atom {
  const text = atomText(source);
  const ink = hasInk(source) || text.trim() !== '';
  return {
    node,
    paragraphIndex,
    text,
    blank: !ink,
    floatingGraphic: hasFloatingGraphic(source),
  };
}

function copyAttributes(from: Element, to: Element): void {
  for (let i = 0; i < from.attributes.length; i++) {
    const a = from.attributes.item(i);
    if (!a) continue;
    if (a.namespaceURI) to.setAttributeNS(a.namespaceURI, a.name, a.value);
    else to.setAttribute(a.name, a.value);
  }
}

export function atomText(el: Element): string {
  if (el.namespaceURI === NS.w) {
    switch (el.localName) {
      case 't':
      case 'delText':
        return el.textContent ?? '';
      case 'tab':
        return '\t';
      case 'br':
        return '\n';
      case 'noBreakHyphen':
        return '-';
      case 'sym':
        return '';
      default:
        break;
    }
  }
  return visibleText(el);
}

const INK_ELEMENTS = ['object', 'drawing', 'pict', 'sym', 'noBreakHyphen'];

function hasInk(el: Element): boolean {
  if (el.namespaceURI === NS.w && INK_ELEMENTS.includes(el.localName ?? '')) return true;
  if (el.namespaceURI === NS.m) return true;
  for (const localName of INK_ELEMENTS) {
    if (hasDescendant(el, NS.w, localName)) return true;
  }
  return hasDescendant(el, NS.m, 'oMath');
}

/**
 * Floating pictures are positioned relative to the page, so moving the run that anchors
 * them would leave the picture behind. Questions with such options are reported and left
 * untouched. Inline drawings and OLE equations, by contrast, travel with their run.
 */
function hasFloatingGraphic(el: Element): boolean {
  for (const drawing of withSelf(el, NS.w, 'drawing')) {
    if (descendants(drawing, NS.wp, 'anchor').length > 0) return true;
  }
  for (const pict of withSelf(el, NS.w, 'pict')) {
    for (const shape of allElements(pict)) {
      const style = shape.getAttribute('style');
      if (style && /position\s*:\s*absolute/i.test(style)) return true;
    }
  }
  return false;
}

function withSelf(el: Element, ns: string, localName: string): Element[] {
  const out = descendants(el, ns, localName);
  if (el.namespaceURI === ns && el.localName === localName) out.unshift(el);
  return out;
}

function allElements(root: Element): Element[] {
  const out: Element[] = [root];
  for (let i = 0; i < root.childNodes.length; i++) {
    const node = root.childNodes.item(i);
    if (isElement(node)) out.push(...allElements(node));
  }
  return out;
}

/**
 * Ensures an atom boundary at each of the given text offsets inside one paragraph's atom
 * list, splitting `w:t` atoms as needed. Offsets are relative to the paragraph text.
 */
export function splitAtomsAtOffsets(atoms: readonly Atom[], offsets: readonly number[]): Atom[] {
  const wanted = [...new Set(offsets)].sort((a, b) => a - b);
  const out: Atom[] = [];
  let cursor = 0;

  for (const atom of atoms) {
    const start = cursor;
    const end = cursor + atom.text.length;
    cursor = end;

    const cuts = wanted.filter((offset) => offset > start && offset < end);
    if (cuts.length === 0 || !isSplittableText(atom)) {
      out.push(atom);
      continue;
    }

    let localStart = 0;
    for (const cut of [...cuts, end]) {
      const localEnd = cut - start;
      out.push(sliceTextAtom(atom, localStart, localEnd));
      localStart = localEnd;
    }
  }

  return out;
}

function isSplittableText(atom: Atom): boolean {
  const run = atom.node;
  if (run.namespaceURI !== NS.w || run.localName !== 'r') return false;
  const textEl = firstChild(run, NS.w, 't');
  return !!textEl && (textEl.textContent ?? '') === atom.text;
}

function sliceTextAtom(atom: Atom, from: number, to: number): Atom {
  const run = atom.node.cloneNode(true) as Element;
  const textEl = firstChild(run, NS.w, 't');
  const slice = atom.text.slice(from, to);
  if (textEl) {
    textEl.textContent = slice;
    preserveSpace(textEl);
  }
  return {
    node: run,
    paragraphIndex: atom.paragraphIndex,
    text: slice,
    blank: slice.trim() === '',
    floatingGraphic: false,
  };
}
