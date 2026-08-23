// Diagnostic: report questions whose stem fingerprint is not unique.
const { DocxPackage } = require('../dist/core/docx/DocxPackage.js');
const { PaperParser } = require('../dist/core/parse/PaperParser.js');
const { NumberingIndex } = require('../dist/core/parse/NumberingIndex.js');
const { questionSignature } = require('../dist/core/generate/Signatures.js');

(async () => {
  const pkg = await DocxPackage.load(process.argv[2]);
  const paper = new PaperParser().parsePart(pkg.documentPart(), new NumberingIndex(pkg.numberingPart()));
  const seen = new Map();
  for (const section of paper.sections) {
    for (const block of section.blocks) {
      const signature = questionSignature(block);
      if (!seen.has(signature)) seen.set(signature, []);
      seen.get(signature).push(block.printedNumber);
    }
  }
  for (const [signature, numbers] of seen) {
    if (numbers.length > 1) console.log(numbers.join(' & '), '=>', JSON.stringify(signature.slice(0, 200)));
  }
})();
