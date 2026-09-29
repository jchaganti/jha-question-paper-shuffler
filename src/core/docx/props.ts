import type { Element } from './dom';
import { NS, childElements, createElement, firstChild, ownerDocumentOf } from './xml';

/**
 * Adding a property to a `w:pPr`, `w:trPr` or `w:sectPr`.
 *
 * OOXML property containers are *sequences*, not bags: the schema fixes the order of their
 * children, and Word refuses to open a document whose properties are out of order. So a new
 * property is never appended - it is spliced in before the first sibling the schema places
 * after it.
 */

/** `w:pPr` children in the order CT_PPr demands. */
export const PPR_ORDER: readonly string[] = [
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

/** `w:trPr` children in the order CT_TrPr demands. */
export const TRPR_ORDER: readonly string[] = [
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

/**
 * `w:sectPr` children in the order CT_SectPr demands.
 *
 * The header and footer references share one schema group, in which they may interleave
 * freely, so they are listed together at the front: anything inserted among them is valid
 * wherever it lands, and everything else belongs after all of them.
 */
export const SECTPR_ORDER: readonly string[] = [
  'headerReference',
  'footerReference',
  'footnotePr',
  'endnotePr',
  'type',
  'pgSz',
  'pgMar',
  'paperSrc',
  'pgBorders',
  'lnNumType',
  'pgNumType',
  'cols',
  'formProt',
  'vAlign',
  'noEndnote',
  'titlePg',
  'textDirection',
  'bidi',
  'rtlGutter',
  'docGrid',
  'printerSettings',
  'sectPrChange',
];

/** `w:rPr` children in the order CT_RPr demands, as far as this application writes them. */
export const RPR_ORDER: readonly string[] = [
  'rStyle',
  'rFonts',
  'b',
  'bCs',
  'i',
  'iCs',
  'caps',
  'smallCaps',
  'strike',
  'dstrike',
  'outline',
  'shadow',
  'emboss',
  'imprint',
  'noProof',
  'snapToGrid',
  'vanish',
  'webHidden',
  'color',
  'spacing',
  'w',
  'kern',
  'position',
  'sz',
  'szCs',
  'highlight',
  'u',
  'effect',
  'bdr',
  'shd',
  'fitText',
  'vertAlign',
  'rtl',
  'cs',
  'em',
  'lang',
];

/**
 * The named property of `props`, added in schema order when it is not there already.
 * Returns it either way, so the caller can set `w:val` on it.
 */
export function ensureChild(props: Element, localName: string, order: readonly string[]): Element {
  const existing = firstChild(props, NS.w, localName);
  if (existing) return existing;

  const rank = order.indexOf(localName);
  const successor = childElements(props).find((child) => {
    if (child.namespaceURI !== NS.w) return false;
    return order.indexOf(child.localName ?? '') > rank;
  });
  const added = createElement(ownerDocumentOf(props), `w:${localName}`, NS.w);
  props.insertBefore(added, successor ?? null);
  return added;
}

/** Adds a boolean property - an empty element means "on" - if it is not there already. */
export function ensureFlag(props: Element, localName: string, order: readonly string[]): void {
  ensureChild(props, localName, order);
}

/** `w:pPr` must be the first child of `w:p`. */
export function ensurePPr(paragraph: Element): Element {
  const existing = firstChild(paragraph, NS.w, 'pPr');
  if (existing) return existing;
  const pPr = createElement(ownerDocumentOf(paragraph), 'w:pPr', NS.w);
  paragraph.insertBefore(pPr, paragraph.firstChild);
  return pPr;
}

/** `w:rPr` must be the first child of `w:r`. */
export function ensureRPr(run: Element): Element {
  const existing = firstChild(run, NS.w, 'rPr');
  if (existing) return existing;
  const rPr = createElement(ownerDocumentOf(run), 'w:rPr', NS.w);
  run.insertBefore(rPr, run.firstChild);
  return rPr;
}
