import type { IDocxPackage } from '../docx/DocxPackage';
import type { Document, Element } from '../docx/dom';
import { PPR_ORDER, RPR_ORDER, SECTPR_ORDER, ensureChild, ensurePPr, ensureRPr } from '../docx/props';
import {
  NS,
  XmlPart,
  childElements,
  createElement,
  descendants,
  firstChild,
  hasDescendant,
  isElement,
  ownerDocumentOf,
  preserveSpace,
  visibleText,
  wVal,
} from '../docx/xml';

const RELS_PART = 'word/_rels/document.xml.rels';
const CONTENT_TYPES_PART = '[Content_Types].xml';
const SETTINGS_PART = 'word/settings.xml';
const STYLES_PART = 'word/styles.xml';

const HEADER_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header';
const HEADER_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';

/** Font sizes are half-points: 22 is 11pt, Word's own default. */
const FALLBACK_SIZE = 22;
/**
 * Sizes, relative to the size the paper writes its own headers in: page one is set 2pt
 * above it (4 half-points), the pages after it at it. The set name has to be spotted on a
 * printed page from across a desk, so it is not written smaller than the paper's own text.
 */
const FIRST_PAGE_LARGER_BY = 4;
/** Never below 8pt, however small the paper's own text is. */
const SMALLEST = 16;

/** A header part for a section that has none, with one paragraph for the stamp to go in. */
const EMPTY_HEADER =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  `<w:hdr xmlns:w="${NS.w}" xmlns:r="${NS.r}"><w:p/></w:hdr>`;

/** Which pages a header covers. `first` needs `w:titlePg`; `even` needs `w:evenAndOddHeaders`. */
type HeaderType = 'default' | 'even' | 'first';

/**
 * Prints the set name - "SET A" - at the top right of every page.
 *
 * A set leaves this tool as a .docx and reaches a candidate as a stack of loose printed
 * pages. By then the file name is gone and the answer key is somewhere else, so the paper
 * has to say which set it is on the page itself. The page header is the one place that
 * repeats without touching a word of the paper.
 *
 * Page one is set in bold and 2pt above the paper's own size, the rest unbolded and at it -
 * the set is announced once and then confirmed. Word draws a different header on page one only when the section
 * asks for one (`w:titlePg`), which these papers do not, so the first-page header is built
 * here as **a copy of the paper's own default header** plus the stamp. That is what keeps
 * page one looking exactly as it did: whatever the header held - in every sample paper, a
 * WordArt watermark - is still there, and the copy brings the header's own relationships
 * with it, so an image in the header is not left pointing at nothing.
 *
 * The stamp joins a header paragraph that prints nothing rather than opening a line of its
 * own, so the header does not grow and the paper does not gain a page. A header with text
 * already in it gets a line above that text, which is where "top right" puts it.
 */
export class HeaderStamp {
  /** Writes `setName` into the headers of every section. Returns how many were stamped. */
  apply(pkg: IDocxPackage, document: XmlPart, setName: string): number {
    if (setName === '') return 0;
    const sections = sectionProperties(document.root);
    if (sections.length === 0) return 0;

    const parts = new PartCache(pkg);
    const rels = new DocumentRelationships(parts);
    const size = headerFontSize(parts);
    const evenPagesDiffer = isOn(firstChildOf(parts.optional(SETTINGS_PART)?.root, 'evenAndOddHeaders'));

    let stamped = 0;
    const laterPages = new Set<string>();

    sections.forEach((sectPr, index) => {
      // Page one of the paper - the only place the set name is set in bold.
      if (index === 0) {
        const part = firstPageHeader(sectPr, rels, parts);
        const bold = { bold: true, size: size + FIRST_PAGE_LARGER_BY };
        if (stampHeader(parts.part(part), setName, bold)) stamped++;
      }

      for (const type of ['default', 'even', 'first'] as const) {
        // The first-page header of section one is the bold one, handled above. Later
        // sections start mid-paper, so their first page is one of the "other" pages.
        if (type === 'first' && index === 0) continue;
        if (type === 'even' && !evenPagesDiffer) continue;
        if (type === 'first' && !isOn(firstChildOf(sectPr, 'titlePg'))) continue;
        const part = headerPartOf(sectPr, type, rels);
        if (part) laterPages.add(part);
      }

      // A section with no header of its own shows nothing on its pages. Give it one, or
      // the set name would stop partway through the paper.
      if (!headerPartOf(sectPr, 'default', rels)) {
        const added = rels.newHeaderPart(parts, EMPTY_HEADER);
        setHeaderReference(sectPr, 'default', added.relationshipId);
        laterPages.add(added.partName);
      }
    });

    const smaller = { bold: false, size: Math.max(SMALLEST, size) };
    for (const part of laterPages) {
      if (stampHeader(parts.part(part), setName, smaller)) stamped++;
    }

    parts.flush();
    return stamped;
  }
}

/**
 * The part holding the header Word draws on page one, made ready to carry the stamp.
 *
 * A paper that already asks for a different first page keeps the header it wrote. Any
 * other gets a fresh part copied from its default header, so that turning `w:titlePg` on
 * cannot replace page one's header with a blank one.
 */
function firstPageHeader(sectPr: Element, rels: DocumentRelationships, parts: PartCache): string {
  const existing = headerPartOf(sectPr, 'first', rels);
  if (existing && isOn(firstChildOf(sectPr, 'titlePg'))) return existing;

  const source = headerPartOf(sectPr, 'default', rels);
  const created = rels.newHeaderPart(parts, source ? parts.text(source) : EMPTY_HEADER, source);
  setHeaderReference(sectPr, 'first', created.relationshipId);
  turnOn(ensureChild(sectPr, 'titlePg', SECTPR_ORDER));
  return created.partName;
}

/**
 * Writes the set name at the right of the header. Returns false when it is already there,
 * which is what stops a paper built from an already-stamped set gaining a second name.
 */
function stampHeader(part: XmlPart, setName: string, style: { bold: boolean; size: number }): boolean {
  const hdr = part.root;
  if (visibleText(hdr).includes(setName)) return false;

  const doc = ownerDocumentOf(hdr);
  const paragraphs = childElements(hdr, NS.w, 'p');
  const first = paragraphs[0];

  // A paragraph that prints nothing on its own line - the watermark line these papers
  // have, or an empty header - has room for the stamp. Anything else gets a line above it.
  let target: Element;
  if (first && printsNothing(first)) {
    target = first;
  } else {
    target = createElement(doc, 'w:p', NS.w);
    hdr.insertBefore(target, first ?? null);
  }

  ensureChild(ensurePPr(target), 'jc', PPR_ORDER).setAttributeNS(NS.w, 'w:val', 'right');
  target.appendChild(stampRun(doc, setName, style));
  return true;
}

/**
 * True when a paragraph leaves its line empty - everything in it floats.
 *
 * Every sample paper puts a watermark in its page header, and a watermark is drawn from
 * the page rather than from the line, so the line it sits on is blank. Whether that reads
 * as blank cannot be decided on text alone: one paper's watermark is WordArt, whose words
 * live in an attribute, and another's is a text box, whose words are ordinary `w:t` - the
 * same thing on the page, and they have to be treated the same way.
 *
 * Alignment is the reason it matters. The stamp right-aligns the paragraph it joins, which
 * moves anything that sits *in* the line - a logo placed in line with the text would slide
 * across the page - but not a shape placed from the page, which ignores alignment.
 */
function printsNothing(paragraph: Element): boolean {
  const copy = paragraph.cloneNode(true) as Element;
  // A fallback is drawn only when Word cannot draw the choice beside it, so leaving both
  // in would count one shape twice.
  for (const fallback of descendants(copy, NS.mc, 'Fallback')) remove(fallback);
  for (const shape of shapesIn(copy)) {
    if (isFloating(shape)) remove(shape);
  }
  return visibleText(copy).trim() === '' && shapesIn(copy).length === 0;
}

function shapesIn(root: Element): Element[] {
  return ['drawing', 'pict', 'object'].flatMap((name) => descendants(root, NS.w, name));
}

/** A shape placed from the page rather than set in the line of text. */
function isFloating(shape: Element): boolean {
  // DrawingML says so outright: `wp:anchor` floats, `wp:inline` does not.
  if (shape.localName === 'drawing') return hasDescendant(shape, NS.wp, 'anchor');
  // VML says it in CSS, on the shape - `v:shape`, `v:rect`, `v:group` and the rest, so the
  // attribute is looked for rather than the element names enumerated.
  return elementsIn(shape).some((el) =>
    (el.getAttribute('style') ?? '').replace(/\s/g, '').includes('position:absolute'),
  );
}

function elementsIn(root: Element): Element[] {
  const out: Element[] = [];
  const walk = (el: Element): void => {
    for (let i = 0; i < el.childNodes.length; i++) {
      const node = el.childNodes.item(i);
      if (!isElement(node)) continue;
      out.push(node);
      walk(node);
    }
  };
  walk(root);
  return out;
}

function remove(node: Element): void {
  node.parentNode?.removeChild(node);
}

function stampRun(doc: Document, text: string, style: { bold: boolean; size: number }): Element {
  const run = createElement(doc, 'w:r', NS.w);
  const rPr = ensureRPr(run);

  // Bold is stated either way. The header's own style may already be bold, and "normal"
  // has to mean normal on a paper that made that choice.
  for (const name of ['b', 'bCs']) {
    const flag = ensureChild(rPr, name, RPR_ORDER);
    if (style.bold) turnOn(flag);
    else flag.setAttributeNS(NS.w, 'w:val', '0');
  }
  for (const name of ['sz', 'szCs']) {
    ensureChild(rPr, name, RPR_ORDER).setAttributeNS(NS.w, 'w:val', String(style.size));
  }
  // No font is named, so the stamp is set in whatever the paper writes its headers in.
  const printed = createElement(doc, 'w:t', NS.w);
  preserveSpace(printed);
  printed.appendChild(doc.createTextNode(text));
  run.appendChild(printed);
  return run;
}

/**
 * The size the paper writes its headers in, in half-points: the `Header` style's own size
 * if it names one, else the document's default size. Read rather than assumed, so the
 * stamp is the size of the paper's own text and not of some house style.
 */
function headerFontSize(parts: PartCache): number {
  const styles = parts.optional(STYLES_PART);
  if (!styles) return FALLBACK_SIZE;

  for (const style of childElements(styles.root, NS.w, 'style')) {
    if (style.getAttributeNS(NS.w, 'styleId') !== 'Header') continue;
    const size = sizeIn(firstChild(style, NS.w, 'rPr'));
    if (size !== undefined) return size;
  }
  const defaults = firstChild(styles.root, NS.w, 'docDefaults');
  const size = sizeIn(firstChildOf(firstChildOf(defaults, 'rPrDefault'), 'rPr'));
  return size ?? FALLBACK_SIZE;
}

function sizeIn(rPr: Element | undefined): number | undefined {
  const sz = rPr && firstChild(rPr, NS.w, 'sz');
  const value = sz ? Number(wVal(sz)) : Number.NaN;
  return Number.isFinite(value) && value > 0 ? value : undefined;
}

/**
 * Every `w:sectPr` of the document, in reading order - the one closing each section that
 * ends mid-document, then the body's own, which closes the last. Section one covers page
 * one, which is the only reason the order matters here.
 */
function sectionProperties(root: Element): Element[] {
  return descendants(root, NS.w, 'sectPr').filter((sectPr) => {
    const parent = sectPr.parentNode;
    if (!parent || parent.nodeType !== 1) return false;
    const element = parent as Element;
    return element.namespaceURI === NS.w && (element.localName === 'body' || element.localName === 'pPr');
  });
}

function headerReferenceOf(sectPr: Element, type: HeaderType): Element | undefined {
  return childElements(sectPr, NS.w, 'headerReference').find(
    (ref) => ref.getAttributeNS(NS.w, 'type') === type,
  );
}

function headerPartOf(sectPr: Element, type: HeaderType, rels: DocumentRelationships): string | undefined {
  const reference = headerReferenceOf(sectPr, type);
  return reference ? rels.target(reference.getAttributeNS(NS.r, 'id')) : undefined;
}

function setHeaderReference(sectPr: Element, type: HeaderType, relationshipId: string): void {
  let reference = headerReferenceOf(sectPr, type);
  if (!reference) {
    reference = createElement(ownerDocumentOf(sectPr), 'w:headerReference', NS.w);
    reference.setAttributeNS(NS.w, 'w:type', type);
    // The references share one schema group at the front of `w:sectPr`, in which they may
    // appear in any order, so anywhere among them is valid.
    const rank = SECTPR_ORDER.indexOf('headerReference');
    const successor = childElements(sectPr).find(
      (child) => child.namespaceURI === NS.w && SECTPR_ORDER.indexOf(child.localName ?? '') > rank,
    );
    sectPr.insertBefore(reference, successor ?? null);
  }
  reference.setAttributeNS(NS.r, 'r:id', relationshipId);
}

/** True when a boolean OOXML property is present and not explicitly switched off. */
function isOn(flag: Element | undefined): boolean {
  if (!flag) return false;
  const value = wVal(flag);
  return value !== '0' && value !== 'false' && value !== 'off';
}

function turnOn(flag: Element): void {
  if (flag.hasAttributeNS(NS.w, 'val')) flag.removeAttributeNS(NS.w, 'val');
}

/** `firstChild` for a parent that may be absent, matching the `w` namespace. */
function firstChildOf(parent: Element | undefined, localName: string): Element | undefined {
  return parent ? firstChild(parent, NS.w, localName) : undefined;
}

/** The `word/_rels/document.xml.rels` part: which relationship id points at which part. */
class DocumentRelationships {
  constructor(private readonly parts: PartCache) {}

  private get root(): Element {
    return this.parts.part(RELS_PART).root;
  }

  /** The part a relationship id points at, as a package path: `word/header2.xml`. */
  target(relationshipId: string | null | undefined): string | undefined {
    if (!relationshipId) return undefined;
    const relationship = childElements(this.root).find((el) => el.getAttribute('Id') === relationshipId);
    const target = relationship?.getAttribute('Target');
    if (!target) return undefined;
    // Targets are written relative to the part that owns them - `word/document.xml`.
    return target.startsWith('/') ? target.slice(1) : `word/${target}`;
  }

  /**
   * Adds a header part holding `xml`, with its own relationship and content type.
   *
   * `copiedFrom`, when given, is the part `xml` was taken from: its relationships are
   * copied alongside, so a picture or a text box in the header still resolves.
   */
  newHeaderPart(
    parts: PartCache,
    xml: string,
    copiedFrom?: string,
  ): { partName: string; relationshipId: string } {
    let number = 1;
    while (parts.exists(`word/header${number}.xml`)) number++;
    const fileName = `header${number}.xml`;
    const partName = `word/${fileName}`;

    parts.create(partName, xml);
    const sourceRels = copiedFrom === undefined ? undefined : parts.raw(relsOf(copiedFrom));
    if (sourceRels !== undefined) parts.create(relsOf(partName), sourceRels);

    contentTypeOverride(parts, partName);
    return { partName, relationshipId: this.add(HEADER_RELATIONSHIP, fileName) };
  }

  private add(type: string, target: string): string {
    const taken = new Set(childElements(this.root).map((el) => el.getAttribute('Id')));
    let number = 1;
    while (taken.has(`rId${number}`)) number++;
    const id = `rId${number}`;

    const relationship = createElement(ownerDocumentOf(this.root), 'Relationship', NS.packageRels);
    relationship.setAttribute('Id', id);
    relationship.setAttribute('Type', type);
    relationship.setAttribute('Target', target);
    this.root.appendChild(relationship);
    return id;
  }
}

/** `word/header2.xml` -> `word/_rels/header2.xml.rels`. */
function relsOf(partName: string): string {
  const cut = partName.lastIndexOf('/');
  return `${partName.slice(0, cut)}/_rels/${partName.slice(cut + 1)}.rels`;
}

/** Declares what kind of XML a new part holds, without which Word will not open the file. */
function contentTypeOverride(parts: PartCache, partName: string): void {
  const types = parts.part(CONTENT_TYPES_PART).root;
  const override = createElement(ownerDocumentOf(types), 'Override', NS.contentTypes);
  override.setAttribute('PartName', `/${partName}`);
  override.setAttribute('ContentType', HEADER_CONTENT_TYPE);
  types.appendChild(override);
}

/**
 * The parts this stamp reads and writes, parsed once each and written back together.
 *
 * Several sections can share one header, so a part must not be parsed twice - the second
 * copy would be written over the first and the first stamp would be lost.
 */
class PartCache {
  private readonly open = new Map<string, XmlPart>();

  constructor(private readonly pkg: IDocxPackage) {}

  exists(partName: string): boolean {
    return this.pkg.getPartText(partName) !== undefined;
  }

  /** The part's text as it stands on disk. */
  raw(partName: string): string | undefined {
    return this.pkg.getPartText(partName);
  }

  text(partName: string): string {
    const xml = this.raw(partName);
    if (xml === undefined) throw new Error(`${partName} is missing from the document.`);
    return xml;
  }

  part(partName: string): XmlPart {
    const already = this.open.get(partName);
    if (already) return already;
    const parsed = XmlPart.parse(this.text(partName));
    this.open.set(partName, parsed);
    return parsed;
  }

  optional(partName: string): XmlPart | undefined {
    return this.exists(partName) ? this.part(partName) : undefined;
  }

  create(partName: string, xml: string): void {
    this.pkg.setPartText(partName, xml);
  }

  flush(): void {
    for (const [partName, part] of this.open) this.pkg.setPartText(partName, part.serialize());
  }
}
