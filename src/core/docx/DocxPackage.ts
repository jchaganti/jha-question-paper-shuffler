import { promises as fs } from 'node:fs';
import JSZip from 'jszip';
import { XmlPart } from './xml';

export const DOCUMENT_PART = 'word/document.xml';
export const NUMBERING_PART = 'word/numbering.xml';

/**
 * Read/modify/write access to a .docx (a zip of XML parts).
 *
 * Design note: the generator never rebuilds a document from scratch. It loads the
 * original package, edits `word/document.xml` in place and writes every other part
 * back byte-for-byte. That is what keeps MathType/OLE equations, embedded images,
 * styles, numbering, headers and footers intact.
 */
export interface IDocxPackage {
  getPartText(part: string): string | undefined;
  setPartText(part: string, xml: string): void;
  toBuffer(): Promise<Buffer>;
}

export class DocxPackage implements IDocxPackage {
  private constructor(private readonly zip: JSZip) {}

  static async load(filePath: string): Promise<DocxPackage> {
    const buffer = await fs.readFile(filePath);
    return DocxPackage.fromBuffer(buffer);
  }

  static async fromBuffer(buffer: Buffer): Promise<DocxPackage> {
    const zip = await JSZip.loadAsync(buffer);
    if (!zip.file(DOCUMENT_PART)) {
      throw new Error('Not a Word document: word/document.xml is missing.');
    }
    // Cache part text eagerly so that later reads are synchronous.
    const texts = new Map<string, string>();
    const names = Object.keys(zip.files).filter((name) => name.endsWith('.xml') || name.endsWith('.rels'));
    for (const name of names) {
      const file = zip.file(name);
      if (!file || file.dir) continue;
      texts.set(name, await file.async('string'));
    }
    const pkg = new DocxPackage(zip);
    pkg.cache = texts;
    return pkg;
  }

  private cache = new Map<string, string>();

  getPartText(part: string): string | undefined {
    return this.cache.get(part);
  }

  setPartText(part: string, xml: string): void {
    this.cache.set(part, xml);
    this.zip.file(part, xml);
  }

  async toBuffer(): Promise<Buffer> {
    return this.zip.generateAsync({
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    });
  }

  async writeTo(filePath: string): Promise<void> {
    await fs.writeFile(filePath, await this.toBuffer());
  }

  /** Convenience: the main document part, parsed. */
  documentPart(): XmlPart {
    const xml = this.getPartText(DOCUMENT_PART);
    if (xml === undefined) throw new Error('word/document.xml is missing.');
    return XmlPart.parse(xml);
  }

  numberingPart(): XmlPart | undefined {
    const xml = this.getPartText(NUMBERING_PART);
    return xml === undefined ? undefined : XmlPart.parse(xml);
  }
}
