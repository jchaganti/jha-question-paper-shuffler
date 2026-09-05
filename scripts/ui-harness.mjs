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
// The real headings, so the harness cannot drift from what the app shows.
import { SKIP_REASON_LABEL } from '../dist/shared/skipReasons.js';

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
const skip = (questionNumber, subject, reason, detail, fix) => ({ questionNumber, subject, reason, detail, fix });
// Detail/fix text is only shown grouped, so the per-question note keeps a stub of each.
const note = (questionNumber, subject, issue) => ({ questionNumber, subject, issue, detail: issue, fix: issue });

const paper = {
  sourceFile: 'D:\\papers\\MTP-2-PCB-XI-2027_4961.docx',
  questionCount: 180,
  // 1..180: this sample paper numbers straight through, as the real ones do.
  questionNumbers: Array.from({ length: 180 }, (_unused, index) => index + 1),
  // One row per run of questions that shuffles on its own - a subject, or a section of one.
  groups: [
    { group: 'PHYSICS - SECTION A', firstQuestionNumber: 1, lastQuestionNumber: 32, questionCount: 32 },
    { group: 'PHYSICS - SECTION B', firstQuestionNumber: 33, lastQuestionNumber: 45, questionCount: 13 },
    { group: 'CHEMISTRY - SECTION A', firstQuestionNumber: 46, lastQuestionNumber: 77, questionCount: 32 },
    { group: 'CHEMISTRY - SECTION B', firstQuestionNumber: 78, lastQuestionNumber: 90, questionCount: 13 },
    { group: 'BIOLOGY - PART 1 SECTION A', firstQuestionNumber: 91, lastQuestionNumber: 135, questionCount: 45 },
    { group: 'BIOLOGY - PART 2 SECTION A', firstQuestionNumber: 136, lastQuestionNumber: 180, questionCount: 45 },
  ],
  unshufflableOptions: [
    skip(
      12,
      'PHYSICS',
      'options-not-found',
      'No option labels were found under this question - neither typed "(A)" to "(D)" nor a lettered list made by Word.',
      'Check that this question has four options and that each one starts with its label.',
    ),
    skip(
      50,
      'CHEMISTRY',
      'unexpected-label-sequence',
      'This question should carry the labels (A) (B) (C) (D), but reading it gives "ABD" - (C) could not be found.',
      'Check that this question has exactly four options, each label used once, each at the start of its line or straight after a Tab.',
    ),
    skip(
      103,
      'BIOLOGY',
      'options-inside-table',
      'The option labels are inside a table. An option is moved by moving its text along the line it sits on, and a table cell is not a line of text.',
      'Select the table, then on the Layout tab choose Convert to Text, separating with tabs.',
    ),
  ],
  pinnedQuestions: [
    {
      questionNumbers: [41, 42],
      subject: 'PHYSICS',
      detail:
        'A floating picture is anchored in the last paragraph of question 41, so it is drawn over ' +
        'whatever follows it - question 42. Moving either question away from the other would take ' +
        'the picture with it, leaving one question without its artwork.',
      fix:
        'In Word, click the picture at the end of question 41 and drag its anchor marker into the ' +
        'question the picture illustrates - or set its Layout Options to "In line with text", which ' +
        'anchors it exactly where it sits.',
    },
  ],
  advisories: [
    advisory(49, 'CHEMISTRY', 'assertion-reason', 'Assertion-Reason / Statement-I-II style question.'),
    advisory(59, 'CHEMISTRY', 'references-other-option', 'Option text refers to another option: "Both (A) and (B)"'),
    advisory(136, 'BIOLOGY', 'catch-all-option', 'Contains a catch-all option: "None of the above"'),
  ],
  // The dry-run panel shows these grouped, exactly as the main process sends them.
  unshufflableGroups: [],
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

// Grouped the same way the main process groups them, so the panel is driven by the same
// shape it sees in the real app rather than by a hand-written second copy.
paper.unshufflableGroups = [...new Set(paper.unshufflableOptions.map((item) => item.reason))].map((reason) => {
  const questions = paper.unshufflableOptions.filter((item) => item.reason === reason);
  return {
    reason,
    label: SKIP_REASON_LABEL[reason],
    fix: questions[0].fix,
    questionNumbers: questions.map((item) => item.questionNumber).sort((a, b) => a - b),
    questions,
  };
});

// The real accounting, inlined into the stub rather than re-implemented in it: the whole
// point of that module is that the three surfaces cannot disagree about the arithmetic, and
// a hand-written copy here would be a fourth answer. It has no runtime imports, so dropping
// the `export` keywords is all it takes to run in the page.
const accounting = (await readFile('dist/web/shared/accounting.js', 'utf8')).replace(
  /^export /gm,
  '',
);

// The real letterer too, for the same reason: the stub must name sets the way the app does.
const { setSuffix } = await import('../dist/core/generate/OutputFolder.js');
const letterer = `const setSuffix = ${setSuffix.toString()};`;

const stub = `
${accounting}
${letterer}
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const PAPER = ${JSON.stringify(paper)};
let progressListener = () => {};

window.shuffler = {
  pickSourceFile: async () => PAPER.sourceFile,
  inspect: async () => ({ ok: true, value: PAPER }),
  dryRun: async (request) => {
    await sleep(250);
    const questionAccounting = accountFor({
      questionNumbers: PAPER.questionNumbers,
      active: request.shuffleQuestions,
      keptByTool: PAPER.pinnedQuestions.flatMap((group) => group.questionNumbers),
      keptByUser: request.questionExclusions,
      labels: QUESTION_LABELS,
    });
    const optionAccounting = accountFor({
      questionNumbers: PAPER.questionNumbers,
      active: request.shuffleOptions,
      keptByTool: PAPER.unshufflableOptions.map((s) => s.questionNumber),
      keptByUser: request.optionExclusions,
      labels: OPTION_LABELS,
    });
    const questionsKept = new Set([...questionAccounting.keptByTool, ...questionAccounting.keptByUser]);
    const optionsKept = new Set([...optionAccounting.keptByTool, ...optionAccounting.keptByUser]);
    const groups = PAPER.groups.map((s) => {
      const numbers = Array.from({ length: s.questionCount }, (_u, i) => s.firstQuestionNumber + i);
      return {
        group: s.group,
        questionCount: s.questionCount,
        movable: request.shuffleQuestions ? numbers.filter((n) => !questionsKept.has(n)).length : 0,
        optionsShuffled: request.shuffleOptions
          ? numbers.filter((n) => !optionsKept.has(n)).length
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
        groups,
        questionsEligibleToMove: groups.reduce((a, s) => a + s.movable, 0),
        optionsToShuffle: groups.reduce((a, s) => a + s.optionsShuffled, 0),
        questionAccounting,
        optionAccounting,
        optionsKeptByTool: PAPER.unshufflableOptions,
        questionsKeptByTool: request.shuffleQuestions ? PAPER.pinnedQuestions : [],
        suggestedForExclusion: PAPER.advisories.filter((a) => !optionsKept.has(a.questionNumber)),
        warnings: [
          ...(unknownNumbers(PAPER.questionNumbers, request.questionExclusions).length > 0
            ? [staleListWarning(unknownNumbers(PAPER.questionNumbers, request.questionExclusions), 'keep these question numbers in place')]
            : []),
          ...(unknownNumbers(PAPER.questionNumbers, request.optionExclusions).length > 0
            ? [staleListWarning(unknownNumbers(PAPER.questionNumbers, request.optionExclusions), 'keep the option order of these questions')]
            : []),
        ],
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
          stage, message, setNumber, setLabel: 'Set ' + setSuffix(setNumber), setCount, setFraction,
          fraction: (setNumber - 1 + setFraction) / setCount,
        });
        await sleep(120);
      }
      sets.push({
        setNumber,
        label: 'Set ' + setSuffix(setNumber),
        fileName: 'MTP-2-PCB-XI-2027_4961 - 02-08-2026-09-15-Set-' + setSuffix(setNumber) + '.docx',
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
