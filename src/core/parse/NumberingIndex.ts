import type { Element } from '../docx/dom';
import { NS, XmlPart, childElements, firstChild, wVal } from '../docx/xml';

export interface NumberingDefinition {
  readonly numId: string;
  readonly abstractNumId?: string;
  /** Number format of level 0, e.g. `decimal`, `upperLetter`, `lowerLetter`. */
  readonly format: string;
  /** Effective start value of level 0 (a `w:startOverride` beats the abstract start). */
  readonly start: number;
  /** Level 0 text pattern, e.g. `%1.` or `(%1)`. */
  readonly levelText: string;
}

/**
 * Index of `word/numbering.xml`, used to tell auto-numbered *question* lists
 * (decimal, "1." "2." ...) from auto-numbered *option* lists ("(A)" "(B)" ...).
 */
export class NumberingIndex {
  private readonly byNumId = new Map<string, NumberingDefinition>();

  constructor(part?: XmlPart) {
    if (!part) return;
    const root = part.root;

    const abstracts = new Map<string, { format: string; start: number; levelText: string }>();
    for (const abstractNum of childElements(root, NS.w, 'abstractNum')) {
      const id = abstractNum.getAttributeNS(NS.w, 'abstractNumId');
      if (!id) continue;
      const level0 = childElements(abstractNum, NS.w, 'lvl').find(
        (lvl) => (lvl.getAttributeNS(NS.w, 'ilvl') ?? '0') === '0',
      );
      if (!level0) continue;
      abstracts.set(id, {
        format: readVal(level0, 'numFmt') ?? 'decimal',
        start: Number(readVal(level0, 'start') ?? '1'),
        levelText: readVal(level0, 'lvlText') ?? '%1.',
      });
    }

    for (const num of childElements(root, NS.w, 'num')) {
      const numId = num.getAttributeNS(NS.w, 'numId');
      if (!numId) continue;
      const abstractRef = firstChild(num, NS.w, 'abstractNumId');
      const abstractNumId = abstractRef ? wVal(abstractRef) : undefined;
      const abstract = abstractNumId ? abstracts.get(abstractNumId) : undefined;

      let start = abstract?.start ?? 1;
      const override = childElements(num, NS.w, 'lvlOverride').find(
        (o) => (o.getAttributeNS(NS.w, 'ilvl') ?? '0') === '0',
      );
      if (override) {
        const startOverride = firstChild(override, NS.w, 'startOverride');
        const value = startOverride ? wVal(startOverride) : undefined;
        if (value !== undefined) start = Number(value);
      }

      this.byNumId.set(numId, {
        numId,
        abstractNumId,
        format: abstract?.format ?? 'decimal',
        start,
        levelText: abstract?.levelText ?? '%1.',
      });
    }
  }

  get(numId: string): NumberingDefinition | undefined {
    return this.byNumId.get(numId);
  }

  /** A list that numbers questions: decimal digits such as "12." rather than "(B)". */
  isDecimalList(numId: string): boolean {
    return this.get(numId)?.format === 'decimal';
  }
}

function readVal(parent: Element, localName: string): string | undefined {
  const el = firstChild(parent, NS.w, localName);
  return el ? wVal(el) : undefined;
}
