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
  symbolTextOf,
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
  /**
   * True when the atom is a `w:sym` whose font this tool cannot read - a picture font such
   * as Wingdings. It carries ink but no readable text, so it cannot be told apart from its
   * neighbours by reading the paragraph. A `w:sym` in the Symbol font is *not* flagged:
   * its character is decoded and contributes to `text` like any other. See `SymbolFont.ts`.
   */
  readonly unreadableSymbol: boolean;
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
  const floatingGraphic = hasFloatingGraphic(source);
  // A floating picture is positioned from the page, not from the text around it, so any
  // words in its text box are not part of the sentence the reader sees. Counting them
  // would put "O A B Cl" in front of the next option label and hide it. It contributes no
  // text to the reading flow; the picture itself is untouched.
  const text = floatingGraphic ? '' : atomText(source);
  const ink = hasInk(source) || text.trim() !== '';
  return {
    node,
    paragraphIndex,
    text,
    blank: !ink,
    floatingGraphic,
    unreadableSymbol: source.namespaceURI === NS.w && source.localName === 'sym' && text === '',
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
        return symbolTextOf(el);
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
export function hasFloatingGraphic(el: Element): boolean {
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

/**
 * Splits one option's atoms into the layout that belongs to the *slot* and the answer text
 * that moves with the *content*.
 *
 * Whole blank atoms at either end are layout - the tab after a label, a spacer paragraph.
 * So is whitespace at the edges *inside* an atom, and that is the subtle case: when four
 * options share a line, the run holding an option often carries the space that separates
 * it from the next label ("(i), (ii) and (iii) " then "(D)"). Letting that space travel
 * with the content glues the next label onto whatever lands there - "all of the above(D)".
 *
 * Trimming it into the padding keeps every separator exactly where the author put it, and
 * loses nothing: the padding is re-emitted for the slot it came from.
 *
 * A floating picture anchored at either end is layout for the same reason: the page, not
 * the option, decides where it is drawn, so it stays behind while the words move. One left
 * *inside* the content is a different matter and is refused by the caller.
 */
export function splitLeadCorePad(content: readonly Atom[]): {
  lead: Atom[];
  core: Atom[];
  pad: Atom[];
} {
  const isLayout = (atom: Atom): boolean => atom.blank || atom.floatingGraphic;
  let start = 0;
  while (start < content.length && isLayout(content[start]!)) start++;
  let end = content.length;
  while (end > start && isLayout(content[end - 1]!)) end--;

  const lead = content.slice(0, start);
  const pad = content.slice(end);
  let core = content.slice(start, end);

  const first = core[0];
  if (first) {
    const spaces = first.text.length - first.text.replace(/^\s+/, '').length;
    if (spaces > 0 && isSplittableText(first)) {
      lead.push(sliceTextAtom(first, 0, spaces));
      core = [sliceTextAtom(first, spaces, first.text.length), ...core.slice(1)];
    }
  }

  // Read the last atom *after* the leading split, in case the core is a single atom.
  const last = core[core.length - 1];
  if (last) {
    const kept = last.text.replace(/\s+$/, '').length;
    if (kept > 0 && kept < last.text.length && isSplittableText(last)) {
      pad.unshift(sliceTextAtom(last, kept, last.text.length));
      core = [...core.slice(0, -1), sliceTextAtom(last, 0, kept)];
    }
  }

  return { lead, core, pad };
}

/**
 * Splits the *last* option's content where the option list ends.
 *
 * The last option is the only one with no following label to bound it, so it runs to the
 * end of the question block and swallows whatever trails behind: the spacer paragraphs
 * that separate one question from the next, a "SECTION B (Attempt any 10 questions)"
 * instruction, the artwork of the question below. Read as content, that made the option
 * "continue on another paragraph" and the question was refused.
 *
 * A blank paragraph ends the list. That is the same reading the parser already applies to
 * a subject - blank paragraphs at its end are page spacing, not part of the last question -
 * and it keeps a genuine continuation refused, because a continuation follows immediately,
 * with no blank paragraph in between.
 */
export function splitTrailingBlock(
  content: readonly Atom[],
  /** Indexes of the option region's paragraphs that hold nothing but whitespace. */
  blankParagraphs: ReadonlySet<number>,
): { kept: Atom[]; trailing: Atom[] } {
  const start = content[0]?.paragraphIndex;
  if (start === undefined) return { kept: [...content], trailing: [] };

  // The first blank paragraph after the one the option starts on ends the list.
  const end = [...blankParagraphs].filter((index) => index > start).sort((a, b) => a - b)[0];
  if (end === undefined) return { kept: [...content], trailing: [] };

  const cut = content.findIndex((atom) => atom.paragraphIndex >= end);
  return cut < 0 ? { kept: [...content], trailing: [] } : { kept: content.slice(0, cut), trailing: content.slice(cut) };
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
    unreadableSymbol: false,
  };
}
