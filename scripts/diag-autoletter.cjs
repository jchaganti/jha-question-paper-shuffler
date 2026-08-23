// Diagnostic: for questions whose options were not found, show the numbering lists used
// inside the question block (numId, level-0 format, items in this block, items in the doc).
const { DocxPackage } = require('../dist/core/docx/DocxPackage.js');
const { PaperParser, paragraphNumId } = require('../dist/core/parse/PaperParser.js');
const { NumberingIndex } = require('../dist/core/parse/NumberingIndex.js');
const { OptionBlockParser } = require('../dist/core/options/OptionBlockParser.js');
const { NS, visibleText } = require('../dist/core/docx/xml.js');

(async () => {
  const pkg = await DocxPackage.load(process.argv[2]);
  const numbering = new NumberingIndex(pkg.numberingPart());
  const paper = new PaperParser().parsePart(pkg.documentPart(), numbering);
  const parser = new OptionBlockParser();

  const docCounts = new Map();
  for (const section of paper.sections) {
    for (const block of section.blocks) {
      for (const node of block.nodes) {
        const numId = paragraphNumId(node);
        if (numId) docCounts.set(numId, (docCounts.get(numId) ?? 0) + 1);
      }
    }
  }

  let handled = 0;
  for (const section of paper.sections) {
    for (const block of section.blocks) {
      const parsed = parser.parse(block);
      if (parsed.ok || parsed.reason !== 'options-not-found') continue;

      const lists = new Map();
      for (const node of block.nodes) {
        if (node === block.questionParagraph) continue;
        const numId = paragraphNumId(node);
        if (!numId) continue;
        if (!lists.has(numId)) lists.set(numId, []);
        lists.get(numId).push(node);
      }

      const described = [...lists.entries()].map(([numId, nodes]) => {
        const def = numbering.get(numId);
        return `numId ${numId} fmt=${def ? def.format : '?'} text="${def ? def.levelText : '?'}" here=${nodes.length} doc=${docCounts.get(numId)}`;
      });
      const letterLists = [...lists.entries()].filter(([numId, nodes]) => {
        const format = numbering.get(numId)?.format ?? '';
        return /Letter/i.test(format) && nodes.length === 4;
      });
      if (letterLists.length === 1) handled++;
      console.log(
        `Q${String(block.printedNumber).padEnd(4)} ${letterLists.length === 1 ? 'HANDLED  ' : 'still odd'} ` +
          `${described.join(' | ') || '(no numbered paragraphs)'}`,
      );
      if (letterLists.length !== 1) {
        for (const node of block.nodes.slice(1, 7)) {
          if (node.namespaceURI === NS.w && node.localName === 'p') {
            console.log(`        | ${visibleText(node).replace(/\s+/g, ' ').trim().slice(0, 90)}`);
          }
        }
      }
    }
  }
  console.log(`\nWould become shufflable: ${handled}`);
})();
