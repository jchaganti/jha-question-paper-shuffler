// Diagnostic: show exactly what the option region of a question looks like, and which
// "(A)"-style labels qualified as labels (start of paragraph, or straight after a tab).
// usage: node scripts/diag-labels.cjs <file.docx> <q1,q2,...>
const { DocxPackage } = require('../dist/core/docx/DocxPackage.js');
const { PaperParser, paragraphNumId } = require('../dist/core/parse/PaperParser.js');
const { NumberingIndex } = require('../dist/core/parse/NumberingIndex.js');
const { NS, visibleText } = require('../dist/core/docx/xml.js');

const LABEL_RE = /\(([A-Ea-e])\)|([A-Ea-e])\)/g;

function qualified(text, index) {
  let i = index - 1;
  while (i >= 0 && text[i] === ' ') i--;
  if (i < 0) return true;
  return text[i] === '\t' || text[i] === '\n';
}

(async () => {
  const pkg = await DocxPackage.load(process.argv[2]);
  const wanted = new Set(process.argv[3].split(',').map(Number));
  const numbering = new NumberingIndex(pkg.numberingPart());
  const paper = new PaperParser().parsePart(pkg.documentPart(), numbering);

  for (const section of paper.sections) {
    for (const block of section.blocks) {
      if (!wanted.has(block.printedNumber)) continue;
      console.log(`\n=========== Q${block.printedNumber} [${section.subject}] key=${paper.answerKey.answerOf(block.printedNumber)}`);
      block.nodes.forEach((node, index) => {
        const kind = node.localName === 'tbl' ? 'TABLE' : node.localName === 'p' ? 'P' : node.localName;
        if (kind === 'TABLE') {
          console.log(`  [${index}] TABLE  ${visibleText(node).replace(/\s+/g, ' ').trim().slice(0, 100)}`);
          return;
        }
        const numId = paragraphNumId(node);
        const def = numId ? numbering.get(numId) : undefined;
        const list = def ? ` list(numId=${numId} ${def.format} "${def.levelText}")` : '';
        const raw = visibleText(node);
        const shown = raw.replace(/\t/g, '»').replace(/\n/g, '¶');
        const labels = [];
        LABEL_RE.lastIndex = 0;
        let m;
        while ((m = LABEL_RE.exec(raw)) !== null) {
          const letter = (m[1] ?? m[2]).toUpperCase();
          labels.push(`${letter}${qualified(raw, m.index) ? '' : '(unqualified)'}@${m.index}`);
        }
        const marker = node === block.questionParagraph ? 'STEM' : 'P   ';
        console.log(`  [${index}] ${marker}${list}${labels.length ? ` labels=[${labels.join(' ')}]` : ''}`);
        if (shown.trim()) console.log(`        "${shown.slice(0, 160)}"`);
      });
    }
  }
})();
