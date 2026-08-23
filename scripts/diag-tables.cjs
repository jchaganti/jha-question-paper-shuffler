// Diagnostic: for questions whose options sit inside a table, show the table's cell
// paragraphs and whether they carry Word numbering (i.e. would be auto-lettered).
// usage: node scripts/diag-tables.cjs <file.docx> <q1,q2,...>
const { DocxPackage } = require('../dist/core/docx/DocxPackage.js');
const { PaperParser, paragraphNumId } = require('../dist/core/parse/PaperParser.js');
const { NumberingIndex } = require('../dist/core/parse/NumberingIndex.js');
const { NS, childElements, visibleText } = require('../dist/core/docx/xml.js');

(async () => {
  const pkg = await DocxPackage.load(process.argv[2]);
  const wanted = new Set(process.argv[3].split(',').map(Number));
  const numbering = new NumberingIndex(pkg.numberingPart());
  const paper = new PaperParser().parsePart(pkg.documentPart(), numbering);

  for (const section of paper.sections) {
    for (const block of section.blocks) {
      if (!wanted.has(block.printedNumber)) continue;
      console.log(`\n=========== Q${block.printedNumber} key=${paper.answerKey.answerOf(block.printedNumber)}`);
      console.log(`  STEM "${visibleText(block.questionParagraph).replace(/\s+/g, ' ').trim().slice(0, 90)}"`);
      block.nodes.slice(1).forEach((node) => {
        if (node.localName === 'tbl') {
          const rows = childElements(node, NS.w, 'tr');
          console.log(`  TABLE ${rows.length} rows`);
          rows.forEach((row, r) => {
            const cells = childElements(row, NS.w, 'tc').map((cell) => {
              const paragraphs = childElements(cell, NS.w, 'p');
              const nums = paragraphs.map((p) => paragraphNumId(p)).filter(Boolean);
              const fmt = nums.length
                ? `{list ${nums.map((n) => `${n}:${numbering.get(n) ? numbering.get(n).format + ' "' + numbering.get(n).levelText + '"' : '?'}`).join(',')}}`
                : '';
              return `${fmt}${visibleText(cell).replace(/\s+/g, ' ').trim().slice(0, 46)}`;
            });
            console.log(`    r${r}: ${cells.join('  ||  ')}`);
          });
          return;
        }
        const numId = paragraphNumId(node);
        const def = numId ? numbering.get(numId) : undefined;
        const text = visibleText(node).replace(/\t/g, '»').replace(/\s+$/, '');
        if (!text.trim() && !numId) return;
        console.log(`  P${def ? ` {list ${numId}:${def.format} "${def.levelText}"}` : ''} "${text.slice(0, 90)}"`);
      });
    }
  }
})();
