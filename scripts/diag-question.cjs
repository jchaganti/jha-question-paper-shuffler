// Diagnostic: print one question's stem and its four option contents.
// usage: node scripts/diag-question.cjs <file.docx> <questionNumber>
const { DocxPackage } = require('../dist/core/docx/DocxPackage.js');
const { PaperParser } = require('../dist/core/parse/PaperParser.js');
const { NumberingIndex } = require('../dist/core/parse/NumberingIndex.js');
const { OptionSetParser } = require('../dist/core/options/OptionSetParser.js');
const { visibleText } = require('../dist/core/docx/xml.js');

(async () => {
  const pkg = await DocxPackage.load(process.argv[2]);
  const wanted = Number(process.argv[3]);
  const paper = new PaperParser().parsePart(pkg.documentPart(), new NumberingIndex(pkg.numberingPart()));
  const numbering = new NumberingIndex(pkg.numberingPart());
  const parser = new OptionSetParser(numbering);
  for (const section of paper.sections) {
    for (const block of section.blocks) {
      if (block.printedNumber !== wanted) continue;
      console.log(`Q${block.printedNumber} [${section.subject}] key=${paper.answerKey.answerOf(wanted)}`);
      console.log('  stem:', visibleText(block.questionParagraph).replace(/\s+/g, ' ').trim().slice(0, 120));
      const parsed = parser.parse(block);
      if (!parsed.ok) {
        console.log('  options not parsed:', parsed.reason, '-', parsed.detail);
        return;
      }
      console.log('  layout:', parsed.options.layout);
      parsed.options.signatures.forEach((signature, index) => {
        console.log(`  (${'ABCD'[index]}) ${signature.split('|')[0]}`);
      });
      return;
    }
  }
  console.log('not found');
})();
