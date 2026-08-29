import JSZip from 'jszip';

/**
 * Builds a small but structurally faithful question paper in memory.
 *
 * The fixture deliberately reproduces the awkward layouts found in real papers:
 * four options on one line, two options per line, one option per line, options
 * auto-lettered by Word, an option holding an embedded OLE equation, and an option
 * anchoring a floating picture.
 */

const PROLOG = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const NAMESPACES = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:v="urn:schemas-microsoft-com:vml"',
  'xmlns:o="urn:schemas-microsoft-com:office:office"',
].join(' ');

export interface FixtureQuestion {
  /** Question stem text. */
  readonly stem: string;
  /** Option paragraphs. Each entry is one paragraph of the option region. */
  readonly optionParagraphs: readonly string[];
  /** Correct answer in the original paper. */
  readonly answer: 'A' | 'B' | 'C' | 'D';
  /**
   * Use Word's automatic lettering instead of typed labels. `true` uses a "(%1)" list;
   * `'plain'` uses a "%1." list, the shape papers use for lettered *statements*.
   */
  readonly autoLettered?: boolean | 'plain';
  /** Extra paragraphs before the options, auto-lettered with a second lettered list. */
  readonly secondLetteredList?: readonly string[];
  /** Which list letters `secondLetteredList`. Defaults to a second "(%1)" list. */
  readonly secondLetteredListId?: string;
  /**
   * Emit only the *first* option paragraph as an auto-lettered list item, leaving the rest
   * as typed labels - the "(A) is lettered by Word, (B)(C)(D) are typed" shape.
   */
  readonly letteredFirstOption?: boolean;
  /** Raw XML injected as an extra option paragraph (e.g. an OLE object or picture). */
  readonly rawOptionParagraph?: string;
}

/** Auto-lettered list ids emitted by the fixture. */
export const LIST = {
  /** upperLetter, "(%1)" - how option lists are written. */
  bracketedOptions: '90',
  /** upperLetter, "%1." - how lettered statement lists are written. */
  plainStatements: '91',
  /** upperLetter, "(%1)" - a second option-shaped list, to test ambiguity. */
  secondBracketed: '92',
} as const;

export interface FixtureSection {
  readonly subject: string;
  readonly startNumber: number;
  readonly numId: string;
  readonly questions: readonly FixtureQuestion[];
}

const escape = (text: string): string =>
  text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

const runs = (text: string): string =>
  text
    .split('\t')
    .map((part, index) => `${index > 0 ? '<w:r><w:tab/></w:r>' : ''}<w:r><w:t xml:space="preserve">${escape(part)}</w:t></w:r>`)
    .join('');

const paragraph = (text: string, numId?: string): string =>
  `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/>${numId ? `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr>` : ''}</w:pPr>${runs(text)}</w:p>`;

const plainParagraph = (text: string): string => `<w:p>${runs(text)}</w:p>`;

/** An option paragraph whose value is an embedded OLE equation. */
export const oleOptionParagraph = (label: string, relationshipId: string): string =>
  `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/></w:pPr>` +
  `<w:r><w:t xml:space="preserve">${label} </w:t></w:r><w:r><w:tab/></w:r>` +
  `<w:r><w:object w:dxaOrig="680" w:dyaOrig="320"><v:shape id="s${relationshipId}" style="width:34pt;height:16pt"/>` +
  `<o:OLEObject Type="Embed" ProgID="Equation.DSMT4" r:id="${relationshipId}"/></w:object></w:r></w:p>`;

/** An option paragraph that anchors a floating picture (must never be moved). */
export const floatingPictureOptionParagraph = (label: string): string =>
  `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/></w:pPr>` +
  `<w:r><w:t xml:space="preserve">${label} </w:t></w:r><w:r><w:tab/></w:r>` +
  `<w:r><w:drawing><wp:anchor distT="0" distB="0"><wp:extent cx="100" cy="100"/></wp:anchor></w:drawing></w:r>` +
  `<w:r><w:t>diagram</w:t></w:r></w:p>`;

function answerKeyTable(entries: readonly { number: number; answer: string }[], columns: number): string {
  const rowCount = Math.ceil(entries.length / columns);
  const grid = `<w:tblGrid>${'<w:gridCol w:w="500"/>'.repeat(columns * 2)}</w:tblGrid>`;
  const cell = (value: string): string => `<w:tc><w:tcPr><w:tcW w:w="500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${value}</w:t></w:r></w:p></w:tc>`;

  const rows: string[] = [];
  for (let row = 0; row < rowCount; row++) {
    const cells: string[] = [];
    for (let column = 0; column < columns; column++) {
      const entry = entries[column * rowCount + row];
      cells.push(cell(entry ? String(entry.number) : ''), cell(entry ? entry.answer : ''));
    }
    rows.push(`<w:tr>${cells.join('')}</w:tr>`);
  }
  return `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>${grid}${rows.join('')}</w:tbl>`;
}

export interface FixtureOptions {
  /**
   * Leave out the "ANSWER KEY" heading, so the key is found only by recognising the table.
   * Papers written this way are supported and are the case where the page break before the
   * key has to go on the table rather than on a heading paragraph.
   */
  readonly omitAnswerKeyHeading?: boolean;
  /** Raw XML inserted immediately before the answer key, e.g. a typed page break. */
  readonly beforeAnswerKey?: string;
  /** Raw XML inserted before every subject heading, e.g. a typed page break. */
  readonly beforeEachSubject?: string;
  /** Raw XML inserted once, before the first subject heading, e.g. a cover line. */
  readonly beforeFirstSubject?: string;
}

/** A paragraph with no properties, for building raw XML in tests. */
export const plainParagraphXml = (text: string): string => plainParagraph(text);

export function buildDocumentXml(
  sections: readonly FixtureSection[],
  options: FixtureOptions = {},
): string {
  const body: string[] = [];
  const key: { number: number; answer: string }[] = [];

  sections.forEach((section, sectionIndex) => {
    if (sectionIndex === 0 && options.beforeFirstSubject) body.push(options.beforeFirstSubject);
    if (options.beforeEachSubject) body.push(options.beforeEachSubject);
    body.push(plainParagraph(section.subject));
    body.push('<w:p/>');
    section.questions.forEach((question, index) => {
      body.push(paragraph(question.stem, section.numId));
      for (const extra of question.secondLetteredList ?? []) {
        body.push(paragraph(extra, question.secondLetteredListId ?? LIST.secondBracketed));
      }
      const optionListId =
        question.autoLettered === true
          ? LIST.bracketedOptions
          : question.autoLettered === 'plain'
            ? LIST.plainStatements
            : undefined;
      question.optionParagraphs.forEach((optionParagraph, position) => {
        const lettered = question.letteredFirstOption
          ? position === 0
            ? LIST.bracketedOptions
            : undefined
          : optionListId;
        body.push(paragraph(optionParagraph, lettered));
      });
      if (question.rawOptionParagraph) body.push(question.rawOptionParagraph);
      key.push({ number: section.startNumber + index, answer: question.answer });
    });
    // Spacer paragraphs that close a subject.
    body.push('<w:p/>', '<w:p/>', '<w:p/>');
  });

  if (options.beforeAnswerKey) body.push(options.beforeAnswerKey);
  if (!options.omitAnswerKeyHeading) {
    body.push(plainParagraph('ANSWER KEY'));
    body.push(plainParagraph('MODEL TEST PAPER'));
  }
  body.push(answerKeyTable(key.sort((a, b) => a.number - b.number), 2));
  body.push('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>');

  return `${PROLOG}<w:document ${NAMESPACES}><w:body>${body.join('')}</w:body></w:document>`;
}

function buildNumberingXml(sections: readonly FixtureSection[]): string {
  const abstracts: string[] = [];
  const nums: string[] = [];

  sections.forEach((section, index) => {
    abstracts.push(
      `<w:abstractNum w:abstractNumId="${index}"><w:lvl w:ilvl="0"><w:start w:val="${section.startNumber}"/>` +
        `<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>`,
    );
    nums.push(`<w:num w:numId="${section.numId}"><w:abstractNumId w:val="${index}"/></w:num>`);
  });

  // Auto-lettered lists: options "(A)", lettered statements "A.", and a second
  // option-shaped list used to test that ambiguity is refused rather than guessed.
  const letterList = (id: string, text: string): void => {
    abstracts.push(
      `<w:abstractNum w:abstractNumId="${id}"><w:lvl w:ilvl="0"><w:start w:val="1"/>` +
        `<w:numFmt w:val="upperLetter"/><w:lvlText w:val="${text}"/></w:lvl></w:abstractNum>`,
    );
    nums.push(`<w:num w:numId="${id}"><w:abstractNumId w:val="${id}"/></w:num>`);
  };
  letterList(LIST.bracketedOptions, '(%1)');
  letterList(LIST.plainStatements, '%1.');
  letterList(LIST.secondBracketed, '(%1)');

  return `${PROLOG}<w:numbering ${NAMESPACES}>${abstracts.join('')}${nums.join('')}</w:numbering>`;
}

export async function buildPaper(
  sections: readonly FixtureSection[],
  options: FixtureOptions = {},
): Promise<Buffer> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `${PROLOG}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
      '</Types>',
  );
  zip.file(
    '_rels/.rels',
    `${PROLOG}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>',
  );
  zip.file(
    'word/_rels/document.xml.rels',
    `${PROLOG}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' +
      '<Relationship Id="rId101" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/oleObject1.bin"/>' +
      '<Relationship Id="rId102" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/oleObject2.bin"/>' +
      '</Relationships>',
  );
  zip.file('word/document.xml', buildDocumentXml(sections, options));
  zip.file('word/numbering.xml', buildNumberingXml(sections));
  zip.file('word/embeddings/oleObject1.bin', 'not-a-real-ole');
  zip.file('word/embeddings/oleObject2.bin', 'not-a-real-ole');
  return zip.generateAsync({ type: 'nodebuffer' });
}

/** The default two-subject paper used by most tests. */
export function defaultSections(): FixtureSection[] {
  return [
    {
      subject: 'PHYSICS',
      startNumber: 1,
      numId: '1',
      questions: [
        {
          stem: 'What is the SI unit of force?',
          optionParagraphs: ['(A)\tnewton\t(B)\tjoule\t(C)\twatt\t(D)\tpascal'],
          answer: 'A',
        },
        {
          stem: 'A body falls freely. Its acceleration is',
          optionParagraphs: ['(A)\tzero\t(B)\tg', '(C)\t2g\t(D)\tg/2'],
          answer: 'B',
        },
        {
          stem: 'Which quantity is a vector?',
          optionParagraphs: ['(A)\tspeed', '(B)\tmass', '(C)\tvelocity', '(D)\ttime'],
          answer: 'C',
        },
        {
          stem: 'Pick the correct statement',
          optionParagraphs: ['only magnitude changes', 'only direction changes', 'both change', 'neither changes'],
          answer: 'D',
          autoLettered: true,
        },
      ],
    },
    {
      subject: 'CHEMISTRY',
      startNumber: 5,
      numId: '2',
      questions: [
        {
          stem: 'Avogadro number is',
          optionParagraphs: ['(A)\t6.022e23\t(B)\t6.022e21', '(C)\t3.011e23\t(D)\tNone of these'],
          answer: 'A',
        },
        {
          stem: 'The molar mass of water is',
          optionParagraphs: ['(A)\t18 g\t(B)\t16 g\t(C)\t2 g\t(D)\t32 g'],
          answer: 'A',
        },
        {
          stem: 'Choose the correct equation',
          optionParagraphs: ['(A)\tfirst'],
          rawOptionParagraph: oleOptionParagraph('(B)', 'rId101') + oleOptionParagraph('(C)', 'rId102') + '<w:p><w:pPr/><w:r><w:t xml:space="preserve">(D) </w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>fourth</w:t></w:r></w:p>',
          answer: 'B',
        },
      ],
    },
  ];
}
