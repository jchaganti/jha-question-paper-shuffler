/**
 * Runs every paper in the sample corpus and prints one row per paper, so that any parser
 * change can be compared against the table in handoff.md §6 before and after.
 *
 *   npm run build:main && node scripts/regress.mjs
 *   node scripts/regress.mjs --generate        # also writes 2 sets each and verifies them
 *   node scripts/regress.mjs --dir "C:/other/papers"
 *
 * The papers are copied into a scratch directory first: the tool writes its output folder
 * next to the source file, and the corpus folder must not be littered.
 *
 * A silent drop in "shuffled" means a new false-positive refusal; a silent rise means a
 * check stopped firing. Either way, read the reason counts underneath.
 */
import { cp, mkdtemp, readdir, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { GenerationService } from '../dist/core/generate/GenerationService.js';

const args = process.argv.slice(2);
const flag = (name) => args.includes(`--${name}`);
const value = (name, fallback) => {
  const at = args.indexOf(`--${name}`);
  return at >= 0 && args[at + 1] ? args[at + 1] : fallback;
};

const corpus = value('dir', 'C:/ps/q-paper');
const generate = flag('generate');
const seed = value('seed', 'regress');

async function main() {
  const entries = (await readdir(corpus, { withFileTypes: true }))
    .filter((e) => e.isFile() && e.name.endsWith('.docx') && !e.name.startsWith('~$'))
    .map((e) => e.name)
    .sort();

  const scratch = await mkdtemp(path.join(tmpdir(), 'qps-regress-'));
  const service = new GenerationService();
  const rows = [];

  try {
    for (const name of entries) {
      const copy = path.join(scratch, name);
      await cp(path.join(corpus, name), copy);
      rows.push(await runOne(service, name, copy));
    }
  } finally {
    await rm(scratch, { recursive: true, force: true });
  }

  const width = Math.max(...rows.map((row) => row.name.length), 5);
  console.log(
    `\n${'Paper'.padEnd(width)} ${'Q'.padStart(4)} ${'shuffled'.padStart(9)} ${'pinned'.padStart(7)}  status`,
  );
  for (const row of rows) {
    console.log(
      `${row.name.padEnd(width)} ${String(row.questions).padStart(4)} ` +
        `${String(row.shuffled).padStart(9)} ${String(row.pinned).padStart(7)}  ${row.status}`,
    );
    for (const [reason, count] of row.reasons) {
      console.log(`${' '.repeat(width)}   ${String(count).padStart(4)} x ${reason}`);
    }
  }
}

async function runOne(service, name, file) {
  const label = name.replace(/\.docx$/i, '').slice(0, 46);
  const request = {
    sourceFile: file,
    setCount: 2,
    shuffleQuestions: true,
    shuffleOptions: true,
    questionExclusions: [],
    optionExclusions: [],
    seed,
    keepQuestionsWhole: true,
  };

  let report;
  try {
    report = await service.dryRun(request);
  } catch (error) {
    return { name: label, questions: 0, shuffled: 0, pinned: 0, reasons: [], status: `REFUSED: ${message(error)}` };
  }

  const counts = new Map();
  for (const item of report.optionsKeptByTool) {
    counts.set(item.reason, (counts.get(item.reason) ?? 0) + 1);
  }
  const row = {
    name: label,
    questions: report.subjects.reduce((sum, subject) => sum + subject.questionCount, 0),
    shuffled: report.optionsToShuffle,
    pinned: report.questionsKeptByTool.reduce((sum, group) => sum + group.questionNumbers.length, 0),
    reasons: [...counts].sort((a, b) => b[1] - a[1]),
    status: 'read',
  };

  if (!generate) return row;

  try {
    const result = await service.generate(request);
    const failed = result.sets.filter((set) => !set.verification.ok);
    row.status =
      failed.length === 0
        ? 'verified'
        : `VERIFY FAILED: ${failed.map((set) => set.verification.checks.filter((c) => !c.ok).map((c) => c.detail).join('; ')).join(' | ')}`;
  } catch (error) {
    row.status = `GENERATE FAILED: ${message(error)}`;
  }
  return row;
}

function message(error) {
  return error instanceof Error ? error.message : String(error);
}

await main();
