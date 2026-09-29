/**
 * The set name at the top right of every page.
 *
 * A generated set reaches a candidate as a stack of loose printed pages, where the file
 * name is gone and the answer key is elsewhere. The page header is the one place that
 * repeats on every page without touching a word of the paper.
 *
 * Most of this file works on a stand-in package rather than a .docx, because what has to
 * be right is the *package*: Word will not open a document whose section properties are out
 * of schema order, whose new part is not declared in `[Content_Types].xml`, or whose header
 * points at a relationship that is not there.
 */
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import type { IDocxPackage } from '../src/core/docx/DocxPackage';
import { DocxPackage } from '../src/core/docx/DocxPackage';
import { NS, XmlPart, childElements, firstChild, visibleText } from '../src/core/docx/xml';
import { GenerationService } from '../src/core/generate/GenerationService';
import { HeaderStamp } from '../src/core/generate/HeaderStamp';
import { buildPaper, defaultSections } from './support/PaperFixture';

const W = `xmlns:w="${NS.w}" xmlns:r="${NS.r}"`;

/** A .docx that never touches the disk: just the parts this stamp reads and writes. */
class FakePackage implements IDocxPackage {
  constructor(private readonly parts: Map<string, string>) {}

  getPartText(part: string): string | undefined {
    return this.parts.get(part);
  }

  setPartText(part: string, xml: string): void {
    this.parts.set(part, xml);
  }

  toBuffer(): Promise<Buffer> {
    throw new Error('not needed');
  }

  names(): string[] {
    return [...this.parts.keys()].sort();
  }
}

const CONTENT_TYPES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  '</Types>';

const header = (body: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W}>${body}</w:hdr>`;

/**
 * A header the way every sample paper writes one: one paragraph holding a watermark.
 *
 * WordArt, drawn with VML, which says it floats in CSS on the shape. Its words are in an
 * attribute, so nothing reads them as text.
 */
const WATERMARK_HEADER = header(
  '<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r><w:pict>' +
    '<v:shape xmlns:v="urn:schemas-microsoft-com:vml" id="WaterMarkObject" ' +
    'style="position:absolute;margin-left:0;margin-top:0;width:534pt;height:152pt;rotation:315">' +
    '<v:textpath on="t" string="DEV JHA"/></v:shape></w:pict></w:r></w:p>',
);

/**
 * The same watermark drawn the other way: a DrawingML text box, anchored to the page. It
 * looks identical, and its words *are* ordinary text, so only the anchor tells them apart.
 */
const TEXT_BOX_WATERMARK_HEADER = header(
  '<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r><w:drawing>' +
    `<wp:anchor xmlns:wp="${NS.wp}" behindDoc="1"><w:txbxContent><w:p><w:r>` +
    '<w:t>DEV JHA</w:t></w:r></w:p></w:txbxContent></wp:anchor></w:drawing></w:r></w:p>',
);

/** A logo set *in* the line, which right-aligning the paragraph would move across the page. */
const INLINE_LOGO_HEADER = header(
  `<w:p><w:r><w:drawing><wp:inline xmlns:wp="${NS.wp}"/></w:drawing></w:r></w:p>`,
);

interface PaperShape {
  /** `word/headerN.xml` parts, by file name. */
  readonly headers?: Record<string, string>;
  /** `w:headerReference` entries to put in the section, by type. */
  readonly references?: Partial<Record<'default' | 'even' | 'first', string>>;
  readonly titlePg?: boolean;
  readonly evenAndOddHeaders?: boolean;
  /** Half-point size for the `Header` style; omitted means the document default of 24. */
  readonly headerStyleSize?: number;
  readonly documentDefaultSize?: number | false;
}

function paper(shape: PaperShape = {}): FakePackage {
  const references = shape.references ?? {};
  const relationships = Object.values(references).map(
    (file, index) =>
      `<Relationship Id="rId${index + 1}" ` +
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" ' +
      `Target="${file}"/>`,
  );
  const sectionReferences = Object.entries(references)
    .map(([type, file]) => {
      const id = `rId${Object.values(references).indexOf(file) + 1}`;
      return `<w:headerReference w:type="${type}" r:id="${id}"/>`;
    })
    .join('');

  // The order here is the one Word writes, and the one the stamp has to splice into.
  const sectPr =
    `<w:sectPr>${sectionReferences}<w:pgSz w:w="11906" w:h="16838"/>` +
    '<w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080"/>' +
    (shape.titlePg ? '<w:titlePg/>' : '') +
    '<w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>';

  const parts = new Map<string, string>([
    ['[Content_Types].xml', CONTENT_TYPES],
    [
      'word/_rels/document.xml.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        relationships.join('') +
        '</Relationships>',
    ],
    [
      'word/document.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}>` +
        `<w:body><w:p><w:r><w:t>Question 1</w:t></w:r></w:p>${sectPr}</w:body></w:document>`,
    ],
  ]);

  const defaultSize = shape.documentDefaultSize === undefined ? 24 : shape.documentDefaultSize;
  if (defaultSize !== false) {
    const headerStyle =
      shape.headerStyleSize === undefined
        ? ''
        : `<w:style w:type="paragraph" w:styleId="Header"><w:rPr><w:sz w:val="${shape.headerStyleSize}"/></w:rPr></w:style>`;
    parts.set(
      'word/styles.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W}>` +
        `<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="${defaultSize}"/></w:rPr></w:rPrDefault></w:docDefaults>` +
        `${headerStyle}</w:styles>`,
    );
  }
  if (shape.evenAndOddHeaders) {
    parts.set(
      'word/settings.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${W}><w:evenAndOddHeaders/></w:settings>`,
    );
  }
  for (const [file, xml] of Object.entries(shape.headers ?? {})) parts.set(`word/${file}`, xml);
  return new FakePackage(parts);
}

function stamp(pkg: FakePackage, setName = 'SET A'): number {
  const document = XmlPart.parse(pkg.getPartText('word/document.xml')!);
  const count = new HeaderStamp().apply(pkg, document, setName);
  pkg.setPartText('word/document.xml', document.serialize());
  return count;
}

/** The header part each `w:headerReference` of the one section points at. */
function headerParts(pkg: FakePackage): Record<string, string> {
  const document = XmlPart.parse(pkg.getPartText('word/document.xml')!);
  const rels = XmlPart.parse(pkg.getPartText('word/_rels/document.xml.rels')!);
  const target = (id: string): string =>
    childElements(rels.root).find((el) => el.getAttribute('Id') === id)!.getAttribute('Target')!;

  const body = firstChild(document.root, NS.w, 'body')!;
  const sectPr = firstChild(body, NS.w, 'sectPr')!;
  const out: Record<string, string> = {};
  for (const reference of childElements(sectPr, NS.w, 'headerReference')) {
    const type = reference.getAttributeNS(NS.w, 'type')!;
    out[type] = `word/${target(reference.getAttributeNS(NS.r, 'id')!)}`;
  }
  return out;
}

/** What the stamp run in a header says: its text, weight and size. */
function stampIn(pkg: FakePackage, partName: string): { text: string; bold: boolean; size: number } | undefined {
  const xml = pkg.getPartText(partName);
  if (xml === undefined) return undefined;
  const part = XmlPart.parse(xml);
  for (const paragraph of childElements(part.root, NS.w, 'p')) {
    for (const run of childElements(paragraph, NS.w, 'r')) {
      const text = visibleText(run);
      if (!text.startsWith('SET ')) continue;
      const rPr = firstChild(run, NS.w, 'rPr')!;
      const bold = firstChild(rPr, NS.w, 'b')!;
      return {
        text,
        bold: bold.getAttributeNS(NS.w, 'val') !== '0',
        size: Number(firstChild(rPr, NS.w, 'sz')!.getAttributeNS(NS.w, 'val')),
      };
    }
  }
  return undefined;
}

const paragraphCount = (pkg: FakePackage, partName: string): number =>
  childElements(XmlPart.parse(pkg.getPartText(partName)!).root, NS.w, 'p').length;

const alignmentOf = (pkg: FakePackage, partName: string, index = 0): string | undefined => {
  const paragraph = childElements(XmlPart.parse(pkg.getPartText(partName)!).root, NS.w, 'p')[index];
  const pPr = paragraph && firstChild(paragraph, NS.w, 'pPr');
  const jc = pPr && firstChild(pPr, NS.w, 'jc');
  return jc?.getAttributeNS(NS.w, 'val') ?? undefined;
};

describe('the set name in the page header', () => {
  it('is bold and larger on page one, and the paper\'s own size on the rest', () => {
    const pkg = paper({ headers: { 'header1.xml': WATERMARK_HEADER }, references: { default: 'header1.xml' } });
    expect(stamp(pkg)).toBe(2);

    const parts = headerParts(pkg);
    // Page one gets a header of its own, which the paper did not have before.
    expect(parts.first).not.toBe(parts.default);
    // The paper writes its headers at 12pt: page one is 14pt, the rest 12pt.
    expect(stampIn(pkg, parts.first!)).toEqual({ text: 'SET A', bold: true, size: 28 });
    expect(stampIn(pkg, parts.default!)).toEqual({ text: 'SET A', bold: false, size: 24 });
  });

  it('sits at the right of the page', () => {
    const pkg = paper({ headers: { 'header1.xml': WATERMARK_HEADER }, references: { default: 'header1.xml' } });
    stamp(pkg);
    const parts = headerParts(pkg);
    expect(alignmentOf(pkg, parts.default!)).toBe('right');
    expect(alignmentOf(pkg, parts.first!)).toBe('right');
  });

  it('asks Word for a different first page, in the order the schema demands', () => {
    const pkg = paper({ headers: { 'header1.xml': WATERMARK_HEADER }, references: { default: 'header1.xml' } });
    stamp(pkg);

    const document = XmlPart.parse(pkg.getPartText('word/document.xml')!);
    const body = firstChild(document.root, NS.w, 'body')!;
    const sectPr = firstChild(body, NS.w, 'sectPr')!;
    // Word refuses to open a document whose section properties are out of order: every
    // reference first, then pgSz, pgMar, cols, titlePg, docGrid.
    expect(childElements(sectPr).map((child) => child.localName)).toEqual([
      'headerReference',
      'headerReference',
      'pgSz',
      'pgMar',
      'cols',
      'titlePg',
      'docGrid',
    ]);
  });

  it('builds page one from the paper\'s own header, so nothing on it is lost', () => {
    const pkg = paper({ headers: { 'header1.xml': WATERMARK_HEADER }, references: { default: 'header1.xml' } });
    stamp(pkg);

    // The watermark travels into the copy: turning on "different first page" must not
    // replace page one's header with a blank one.
    const parts = headerParts(pkg);
    expect(pkg.getPartText(parts.first!)).toContain('WaterMarkObject');
    expect(pkg.getPartText(parts.default!)).toContain('WaterMarkObject');
  });

  it('brings the header\'s own relationships along with the copy', () => {
    const pkg = paper({
      headers: { 'header1.xml': header('<w:p><w:r><w:drawing r:embed="rId9"/></w:r></w:p>') },
      references: { default: 'header1.xml' },
    });
    pkg.setPartText(
      'word/_rels/header1.xml.rels',
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId9" Type="image" Target="media/logo.png"/></Relationships>',
    );
    stamp(pkg);

    // A picture in the header would otherwise point at a relationship the copy has not got.
    const parts = headerParts(pkg);
    const copyRels = parts.first!.replace('word/', 'word/_rels/') + '.rels';
    expect(pkg.getPartText(copyRels)).toContain('Target="media/logo.png"');
  });

  it('declares the new part, or Word will not open the file', () => {
    const pkg = paper({ headers: { 'header1.xml': WATERMARK_HEADER }, references: { default: 'header1.xml' } });
    stamp(pkg);

    const parts = headerParts(pkg);
    const types = pkg.getPartText('[Content_Types].xml')!;
    expect(types).toContain(`PartName="/${parts.first}"`);
    expect(types).toContain('wordprocessingml.header+xml');
    // The new element must not carry a namespace of its own, or the package is invalid.
    expect(types).not.toContain('xmlns=""');
    expect(pkg.getPartText('word/_rels/document.xml.rels')).not.toContain('xmlns=""');
  });

  it('joins a header line that prints nothing, rather than adding one', () => {
    const pkg = paper({ headers: { 'header1.xml': WATERMARK_HEADER }, references: { default: 'header1.xml' } });
    stamp(pkg);

    // Every sample paper's header is a single watermark line. Adding a line to it would
    // push the body down and could cost the paper a page.
    expect(paragraphCount(pkg, headerParts(pkg).default!)).toBe(1);
  });

  it('treats a watermark whose words are real text as the blank line it prints', () => {
    // One paper's watermark is a text box rather than WordArt. The same thing on the page
    // has to be read the same way, or two papers that look alike come out differently.
    const pkg = paper({
      headers: { 'header1.xml': TEXT_BOX_WATERMARK_HEADER },
      references: { default: 'header1.xml' },
    });
    stamp(pkg);
    expect(paragraphCount(pkg, headerParts(pkg).default!)).toBe(1);
  });

  it('leaves a header logo where the author put it', () => {
    // A picture set in the line moves when the line is right-aligned, so that line is not
    // the stamp's to take.
    const pkg = paper({ headers: { 'header1.xml': INLINE_LOGO_HEADER }, references: { default: 'header1.xml' } });
    stamp(pkg);

    const part = headerParts(pkg).default!;
    expect(paragraphCount(pkg, part)).toBe(2);
    expect(alignmentOf(pkg, part, 1)).toBeUndefined();
  });

  it('opens a line of its own above a header that already says something', () => {
    const printed = header('<w:p><w:r><w:t>DEV JHA</w:t></w:r></w:p>');
    const pkg = paper({ headers: { 'header1.xml': printed }, references: { default: 'header1.xml' } });
    stamp(pkg);

    const part = headerParts(pkg).default!;
    expect(paragraphCount(pkg, part)).toBe(2);
    expect(alignmentOf(pkg, part, 0)).toBe('right');
    // The paper's own header line is untouched, and still below the set name.
    const paragraphs = childElements(XmlPart.parse(pkg.getPartText(part)!).root, NS.w, 'p');
    expect(visibleText(paragraphs[0]!)).toBe('SET A');
    expect(visibleText(paragraphs[1]!)).toBe('DEV JHA');
  });

  it('gives a paper with no header at all one for each half of the job', () => {
    const pkg = paper();
    expect(stamp(pkg)).toBe(2);

    const parts = headerParts(pkg);
    expect(parts.default).toBeDefined();
    expect(stampIn(pkg, parts.first!)?.bold).toBe(true);
    expect(stampIn(pkg, parts.default!)?.bold).toBe(false);
    expect(pkg.names()).toContain('word/header1.xml');
    expect(pkg.names()).toContain('word/header2.xml');
  });

  it('keeps a first-page header the author wrote', () => {
    const authors = header('<w:p><w:r><w:t>Model Test Paper</w:t></w:r></w:p>');
    const pkg = paper({
      headers: { 'header1.xml': WATERMARK_HEADER, 'header2.xml': authors },
      references: { default: 'header1.xml', first: 'header2.xml' },
      titlePg: true,
    });
    stamp(pkg);

    // The paper already asks for a different first page, so that header is stamped where
    // it stands - replacing it would throw away what the author put there.
    const parts = headerParts(pkg);
    expect(parts.first).toBe('word/header2.xml');
    expect(visibleText(XmlPart.parse(pkg.getPartText(parts.first!)!).root)).toContain('Model Test Paper');
    expect(stampIn(pkg, parts.first!)?.bold).toBe(true);
  });

  it('stamps the even-page header only when the paper uses one', () => {
    const shape = {
      headers: { 'header1.xml': WATERMARK_HEADER, 'header2.xml': WATERMARK_HEADER },
      references: { default: 'header1.xml', even: 'header2.xml' },
    };
    const off = paper(shape);
    stamp(off);
    // "Different odd and even pages" is off, so Word never draws the even header.
    expect(stampIn(off, 'word/header2.xml')).toBeUndefined();

    const on = paper({ ...shape, evenAndOddHeaders: true });
    stamp(on);
    expect(stampIn(on, 'word/header2.xml')).toEqual({ text: 'SET A', bold: false, size: 24 });
  });

  it('is the size the paper writes its own headers in', () => {
    const sized = (shape: PaperShape): (number | undefined)[] => {
      const pkg = paper({
        ...shape,
        headers: { 'header1.xml': WATERMARK_HEADER },
        references: { default: 'header1.xml' },
      });
      stamp(pkg);
      const parts = headerParts(pkg);
      return [stampIn(pkg, parts.first!)?.size, stampIn(pkg, parts.default!)?.size];
    };

    // Page one sits 2pt above the paper's own size; the pages after it sit at it.
    expect(sized({ documentDefaultSize: 24 })).toEqual([28, 24]);
    expect(sized({ documentDefaultSize: 20 })).toEqual([24, 20]);
    // The Header style's own size wins over the document default.
    expect(sized({ documentDefaultSize: 24, headerStyleSize: 18 })).toEqual([22, 18]);
    // Nothing to read: Word's own default of 11pt.
    expect(sized({ documentDefaultSize: false })).toEqual([26, 22]);
    // Never smaller than 8pt, however small the paper's own text is.
    expect(sized({ documentDefaultSize: 12 })).toEqual([16, 16]);
  });

  it('does nothing twice, so a set built from a set gains no second name', () => {
    const pkg = paper({ headers: { 'header1.xml': WATERMARK_HEADER }, references: { default: 'header1.xml' } });
    stamp(pkg);
    const after = pkg.getPartText(headerParts(pkg).default!);

    expect(stamp(pkg)).toBe(0);
    expect(pkg.getPartText(headerParts(pkg).default!)).toBe(after);
  });

  it('does nothing when there is no name to write', () => {
    const pkg = paper({ headers: { 'header1.xml': WATERMARK_HEADER }, references: { default: 'header1.xml' } });
    expect(stamp(pkg, '')).toBe(0);
    expect(pkg.getPartText('word/document.xml')).not.toContain('titlePg');
  });
});

describe('a generated set', () => {
  let workingDir = '';
  let sourceFile = '';
  const service = new GenerationService();

  beforeEach(async () => {
    workingDir = await fs.mkdtemp(path.join(os.tmpdir(), 'shuffler-header-'));
    sourceFile = path.join(workingDir, 'Sample Paper.docx');
    await fs.writeFile(sourceFile, await buildPaper(defaultSections()));
  });

  afterEach(async () => {
    await fs.rm(workingDir, { recursive: true, force: true });
  });

  it('says which set it is on every page, and still verifies', async () => {
    const result = await service.generate({
      sourceFile,
      shuffleQuestions: true,
      shuffleOptions: true,
      questionExclusions: [],
      optionExclusions: [],
      setCount: 3,
      seed: 'headers',
    });

    for (const [index, set] of result.sets.entries()) {
      expect(set.verification.ok).toBe(true);

      const pkg = await DocxPackage.fromBuffer(await fs.readFile(set.filePath));
      const expected = `SET ${'ABC'[index]}`;
      const headers = [...Array(4).keys()]
        .map((n) => pkg.getPartText(`word/header${n + 1}.xml`))
        .filter((xml): xml is string => xml !== undefined);

      expect(headers.length).toBeGreaterThanOrEqual(2);
      for (const xml of headers) expect(xml).toContain(expected);
      // One set's name never appears in another's paper.
      for (const other of ['SET A', 'SET B', 'SET C'].filter((name) => name !== expected)) {
        for (const xml of headers) expect(xml).not.toContain(other);
      }
    }
  });
});
