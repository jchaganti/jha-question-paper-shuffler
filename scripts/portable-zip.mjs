/**
 * Packs `release/win-unpacked` into one zip that can be emailed or dropped on a shared
 * drive: the recipient extracts it anywhere and runs the .exe. Nothing is installed, no
 * administrator rights are needed, and nothing is written to the registry.
 *
 *   npm run portable        # after npm run pack
 *
 * Uses JSZip - already a dependency, because the tool reads and writes .docx packages -
 * rather than an external archiver, so this works on a machine whose security software
 * will not let `7za.exe` run. That is exactly the machine this was written on: see
 * "Building the installer" in README.md.
 */
import { createReadStream, createWriteStream } from 'node:fs';
import { readdir, readFile, stat } from 'node:fs/promises';
import path from 'node:path';
import { pipeline } from 'node:stream/promises';
import JSZip from 'jszip';

const root = path.resolve('release', 'win-unpacked');
const { version } = JSON.parse(await readFile('package.json', 'utf8'));

/** Every file under `dir`, as paths relative to it. */
async function walk(dir, prefix = '') {
  const out = [];
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    const rel = prefix ? `${prefix}/${entry.name}` : entry.name;
    if (entry.isDirectory()) out.push(...(await walk(full, rel)));
    else out.push({ full, rel });
  }
  return out;
}

/** What to tell whoever receives the zip, before they double-click anything. */
const READ_ME = `Question Paper Shuffler ${version}

To run it
---------
1. Extract this whole folder somewhere you can write to - your Desktop, Documents,
   or a shared drive. Do not run it from inside the zip.
2. Double-click "Question Paper Shuffler.exe".

Windows may show a blue "Windows protected your PC" box the first time, because this
program is not code-signed. Click "More info", then "Run anyway". That message is about
the missing signature, not about anything the program does.

Nothing is installed. To remove it, delete the folder.

What it does
------------
Give it one Word question paper (.docx) with its answer key at the end, and it writes
however many shuffled sets you ask for, each with its own correct answer key. Press
"Dry run" first: it writes nothing and tells you exactly what would happen.

The window opens with a section called "How to write the Word document so that every
question can be shuffled" - worth reading once by whoever types the papers.
`;

async function main() {
  if (!(await stat(root).catch(() => null))) {
    throw new Error(`${root} does not exist. Run "npm run pack" first.`);
  }

  const files = await walk(root);
  const zip = new JSZip();
  // One folder inside the zip, so extracting never scatters 125 files into Downloads.
  const folder = `Question Paper Shuffler ${version}`;
  zip.file(`${folder}/READ ME FIRST.txt`, READ_ME);
  for (const file of files) {
    zip.file(`${folder}/${file.rel}`, createReadStream(file.full));
  }

  const out = path.resolve('release', `Question Paper Shuffler ${version} (portable).zip`);
  await pipeline(
    zip.generateNodeStream({
      type: 'nodebuffer',
      // Streamed rather than assembled in memory: the folder is about 270 MB.
      streamFiles: true,
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    }),
    createWriteStream(out),
  );

  const { size } = await stat(out);
  console.log(`wrote ${out} (${Math.round(size / 1024 / 1024)} MB, ${files.length + 1} entries)`);
}

await main();
