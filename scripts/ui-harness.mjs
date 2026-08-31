/**
 * Builds dist/web/renderer/harness.html: the real UI (index.html + styles.css +
 * renderer.js) driven by a stubbed preload bridge that returns sample data.
 *
 * Lets the renderer's own code paths - dry-run panel, progress bars, results list - be
 * opened in a plain browser without Electron and without touching a real paper.
 *
 *   npm run build && node scripts/ui-harness.mjs && start dist/web/renderer/harness.html
 */
import { readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';

const dir = path.resolve('dist/web/renderer');
const html = await readFile(path.join(dir, 'index.html'), 'utf8');
const renderer = await readFile(path.join(dir, 'renderer.js'), 'utf8');

// The page watermark is a linked file in the real app. The harness may be opened from a
// path the browser will not resolve relative URLs against, so inline it like the rest.
const watermark = await readFile(path.join(dir, 'paper-stack.svg'), 'utf8');
const css = (await readFile(path.join(dir, 'styles.css'), 'utf8')).replace(
  "url('paper-stack.svg')",
  `url("data:image/svg+xml;base64,${Buffer.from(watermark).toString('base64')}")`,
);
if (/^\s*import\s/m.test(renderer)) {
  throw new Error('renderer.js now has runtime imports; the harness can no longer inline it.');
}

const advisory = (questionNumber, subject, kind, detail) => ({ questionNumber, subject, kind, detail });
const skip = (questionNumber, subject, reason, detail) => ({ questionNumber, subject, reason, detail });
// Detail/fix text is only shown grouped, so the per-question note keeps a stub of each.
const note = (questionNumber, subject, issue) => ({ questionNumber, subject, issue, detail: issue, fix: issue });

const paper = {
  sourceFile: 'D:\\papers\\MTP-2-PCB-XI-2027_4961.docx',
  questionCount: 180,
  subjects: [
    { subject: 'PHYSICS', firstQuestionNumber: 1, lastQuestionNumber: 45, questionCount: 45 },
    { subject: 'CHEMISTRY', firstQuestionNumber: 46, lastQuestionNumber: 90, questionCount: 45 },
    { subject: 'BIOLOGY', firstQuestionNumber: 91, lastQuestionNumber: 180, questionCount: 90 },
  ],
  unshufflableOptions: [
    skip(12, 'PHYSICS', 'options-not-found', 'No "(A)...(D)" labels found; the options are probably auto-lettered by Word.'),
    skip(50, 'CHEMISTRY', 'unexpected-label-sequence', 'Expected labels A,B,C,D but found "ABD".'),
    skip(103, 'BIOLOGY', 'options-inside-table', 'Option labels were found inside a table.'),
  ],
  advisories: [
    advisory(49, 'CHEMISTRY', 'assertion-reason', 'Assertion-Reason / Statement-I-II style question.'),
    advisory(59, 'CHEMISTRY', 'references-other-option', 'Option text refers to another option: "Both (A) and (B)"'),
    advisory(136, 'BIOLOGY', 'catch-all-option', 'Contains a catch-all option: "None of the above"'),
  ],
  layoutNotes: [
    note(116, 'BIOLOGY', 'mixed-auto-and-typed-labels'),
    note(127, 'BIOLOGY', 'mixed-auto-and-typed-labels'),
    note(128, 'BIOLOGY', 'mixed-auto-and-typed-labels'),
    note(131, 'BIOLOGY', 'mixed-auto-and-typed-labels'),
    note(8, 'PHYSICS', 'label-not-after-tab'),
    note(26, 'PHYSICS', 'label-not-after-tab'),
    note(89, 'CHEMISTRY', 'several-lettered-lists'),
  ],
  layoutNoteGroups: [
    {
      issue: 'mixed-auto-and-typed-labels',
      label: 'Some options are lettered by Word and the rest typed by hand',
      fix: 'Letter all four options the same way: either let Word letter all four, or type all four labels.',
      questionNumbers: [116, 127, 128, 131],
    },
    {
      issue: 'label-not-after-tab',
      label: 'An option label has no tab in front of it',
      fix: 'Press Tab before each option label, so the label always follows a tab.',
      questionNumbers: [8, 26],
    },
    {
      issue: 'several-lettered-lists',
      label: 'More than one lettered list could have been the options',
      fix:
        'Letter only the options with a bracketed "(A) (B) (C) (D)" list, and give any other lettered ' +
        'list a different style, such as "A." or "(i) (ii)".',
      questionNumbers: [89],
    },
  ],
};

const stub = `
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const PAPER = ${JSON.stringify(paper)};
let progressListener = () => {};

window.shuffler = {
  pickSourceFile: async () => PAPER.sourceFile,
  inspect: async () => ({ ok: true, value: PAPER }),
  dryRun: async (request) => {
    await sleep(250);
    const kept = new Set(request.optionExclusions);
    const unshufflable = new Set(PAPER.unshufflableOptions.map((s) => s.questionNumber));
    const pinned = new Set(request.questionExclusions);
    const subjects = PAPER.subjects.map((s) => {
      const numbers = Array.from({ length: s.questionCount }, (_u, i) => s.firstQuestionNumber + i);
      return {
        subject: s.subject,
        questionCount: s.questionCount,
        movable: request.shuffleQuestions ? numbers.filter((n) => !pinned.has(n)).length : 0,
        optionsShuffled: request.shuffleOptions
          ? numbers.filter((n) => !kept.has(n) && !unshufflable.has(n)).length
          : 0,
      };
    });
    return {
      ok: true,
      value: {
        paper: PAPER,
        setCount: request.setCount,
        shuffleQuestions: request.shuffleQuestions,
        shuffleOptions: request.shuffleOptions,
        subjects,
        questionsEligibleToMove: subjects.reduce((a, s) => a + s.movable, 0),
        optionsToShuffle: subjects.reduce((a, s) => a + s.optionsShuffled, 0),
        optionsKeptByUser: [...kept].sort((a, b) => a - b),
        optionsKeptByTool: PAPER.unshufflableOptions,
        suggestedForExclusion: PAPER.advisories.filter((a) => !kept.has(a.questionNumber)),
        warnings: request.questionExclusions.includes(999)
          ? ['This paper has no question 999, so those entries will have no effect.']
          : [],
      },
    };
  },
  generate: async (request) => {
    const setCount = request.setCount;
    const steps = [
      ['options', 'Reading the paper', 0.1],
      ['options', 'Shuffling options', 0.3],
      ['ordering', 'Re-ordering questions and updating the answer key', 0.6],
      ['packaging', 'Packaging the document', 0.85],
      ['writing', 'Writing the file', 0.9],
      ['verifying', 'Verifying', 1],
    ];
    progressListener({ stage: 'planning', message: 'Planning the sets...', fraction: 0 });
    const sets = [];
    for (let setNumber = 1; setNumber <= setCount; setNumber++) {
      for (const [stage, message, setFraction] of steps) {
        progressListener({
          stage, message, setNumber, setCount, setFraction,
          fraction: (setNumber - 1 + setFraction) / setCount,
        });
        await sleep(120);
      }
      sets.push({
        setNumber,
        fileName: 'MTP-2-PCB-XI-2027_4961 - Set ' + String(setNumber).padStart(2, '0') + '.docx',
        filePath: 'D:\\\\papers\\\\question-sets - 31-08-2026-13-21\\\\set.docx',
        seed: 'demo#set' + setNumber,
        questionsMoved: 170 + setNumber,
        optionsShuffled: 160,
        skipped: PAPER.unshufflableOptions,
        mappings: [],
        verification: { ok: setNumber !== 3, checks: [
          { name: 'Question count unchanged', ok: true, detail: '180 of 180 questions' },
          { name: 'New answer key points at the originally correct text', ok: setNumber !== 3, detail: setNumber === 3 ? 'Q7 key says B (was Q22 answer C)' : '164 answers checked' },
        ] },
      });
    }
    progressListener({ stage: 'done', message: 'Generated ' + setCount + ' set(s)', setNumber: setCount, setCount, setFraction: 1, fraction: 1 });
    return {
      ok: true,
      value: {
        outputFolder: 'D:\\\\papers\\\\question-sets - 31-08-2026-13-21',
        reportFile: 'D:\\\\papers\\\\question-sets - 31-08-2026-13-21\\\\_generation-report.pdf',
        seed: request.seed || 'c594e1fd-be30-4738-8804-870967699570',
        paper: PAPER,
        sets,
      },
    };
  },
  revealFolder: async () => {},
  onProgress: (listener) => { progressListener = listener; },
};
`;

const harness = html
  .replace(/<meta http-equiv="Content-Security-Policy"[^>]*>/, '')
  .replace('<link rel="stylesheet" href="styles.css" />', `<style>${css}</style>`)
  // Inlined rather than linked so the harness works even when opened as a snapshot.
  .replace(
    '<script type="module" src="renderer.js"></script>',
    `<script>${stub}</script>\n    <script type="module">${renderer}</script>`,
  );

const out = path.join(dir, 'harness.html');
await writeFile(out, harness);
console.log(`wrote ${out}`);
