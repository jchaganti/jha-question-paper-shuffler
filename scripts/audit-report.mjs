/**
 * Checks that a dry-run report never contradicts itself.
 *
 *   npm run build:main && node scripts/audit-report.mjs
 *   node scripts/audit-report.mjs --dir "C:/other/papers"
 *
 * Every number in the report is a view of the same paper, so they all have to agree: the
 * section table has to sum to the headline, the headline has to account for every question,
 * finding 1 has to hold exactly the questions the headline says were unreadable, and a
 * question number the report mentions has to exist in the paper. A report that says
 * "options would be shuffled for 77 of 100" and then explains only 3 of the other 23 is
 * wrong even when every individual number in it is right.
 *
 * Each paper is checked under several settings, including an exclusion list belonging to a
 * different paper - the case that started this: numbers left over in the UI from the paper
 * looked at before. Nothing is written; dry runs touch no files.
 */
import { readdir } from 'node:fs/promises';
import path from 'node:path';
import { GenerationService } from '../dist/core/generate/GenerationService.js';

const args = process.argv.slice(2);
const at = args.indexOf('--dir');
const corpus = at >= 0 && args[at + 1] ? args[at + 1] : 'C:/ps/q-paper';

/** The settings each paper is checked under. */
const SETTINGS = [
  ['as it comes', {}],
  // Numbers from some other paper, plus some of this one's.
  ['a list left over from another paper', {
    questionExclusions: [1, 999],
    optionExclusions: [3, 10, 11, 102, 112, 117, 124, 158, 193],
  }],
  ['options not shuffled', { shuffleOptions: false }],
  ['questions not shuffled', { shuffleQuestions: false }],
];

const sum = (values) => values.reduce((a, b) => a + b, 0);
const ascending = (values) => [...new Set(values)].sort((a, b) => a - b);
const same = (a, b) => a.length === b.length && a.every((value, index) => value === b[index]);

async function main() {
  const papers = (await readdir(corpus, { withFileTypes: true }))
    .filter((e) => e.isFile() && e.name.endsWith('.docx') && !e.name.startsWith('~$'))
    .map((e) => e.name)
    .sort();

  const service = new GenerationService();
  let checked = 0;
  const failures = [];

  for (const name of papers) {
    for (const [label, overrides] of SETTINGS) {
      let report;
      try {
        report = await service.dryRun({
          sourceFile: path.join(corpus, name),
          setCount: 2,
          shuffleQuestions: true,
          shuffleOptions: true,
          questionExclusions: [],
          optionExclusions: [],
          ...overrides,
        });
      } catch (error) {
        // A paper the tool refuses outright has no report to audit, and says so clearly.
        if (label === 'as it comes') {
          console.log(`- ${name.slice(0, 44)}: refused, nothing to audit`);
        }
        break;
      }
      for (const problem of audit(report, overrides)) {
        failures.push(`${name.slice(0, 44)} [${label}] ${problem}`);
      }
      checked++;
    }
  }

  console.log(`\n${checked} report(s) audited across ${papers.length} paper(s).`);
  if (failures.length === 0) {
    console.log('No contradictions.');
    return;
  }
  for (const failure of failures) console.log(`  ${failure}`);
  console.log(`\n${failures.length} contradiction(s).`);
  process.exitCode = 1;
}

/** Yields one message per way this report disagrees with itself. */
function* audit(report, overrides) {
  const paper = report.paper;
  const say = function* (ok, message) {
    if (!ok) yield message;
  };

  // The paper's own numbering, which every other number is checked against.
  yield* say(
    paper.questionNumbers.length === paper.questionCount,
    `questionNumbers has ${paper.questionNumbers.length} but questionCount is ${paper.questionCount}`,
  );
  yield* say(
    same(paper.questionNumbers, ascending(paper.questionNumbers)),
    'questionNumbers is not ascending and free of repeats',
  );
  yield* say(
    sum(report.groups.map((group) => group.questionCount)) === paper.questionCount,
    'the section table does not sum to the number of questions',
  );
  yield* say(
    same(paper.groups.map((g) => g.group), report.groups.map((g) => g.group)),
    'the paper summary and the dry run disagree about the sections',
  );
  for (const group of paper.groups) {
    yield* say(
      group.lastQuestionNumber - group.firstQuestionNumber + 1 === group.questionCount,
      `${group.group} shows ${group.firstQuestionNumber}-${group.lastQuestionNumber} but holds ${group.questionCount} questions`,
    );
  }

  // The table against the headline.
  yield* say(
    sum(report.groups.map((g) => g.movable)) === report.questionsEligibleToMove,
    'the "free to move" column does not sum to the headline',
  );
  yield* say(
    sum(report.groups.map((g) => g.optionsShuffled)) === report.optionsToShuffle,
    'the "options shuffled" column does not sum to the headline',
  );
  for (const group of report.groups) {
    yield* say(
      group.movable <= group.questionCount && group.optionsShuffled <= group.questionCount,
      `${group.group} claims to shuffle more questions than it has`,
    );
  }

  // The headline against itself: every question in exactly one outcome.
  for (const [what, accounting, headline] of [
    ['questions', report.questionAccounting, report.questionsEligibleToMove],
    ['options', report.optionAccounting, report.optionsToShuffle],
  ]) {
    yield* say(accounting.total === paper.questionCount, `${what}: the accounting is not over the whole paper`);
    yield* say(accounting.shuffled === headline, `${what}: the accounting and the headline disagree`);
    if (accounting.active) {
      yield* say(
        accounting.shuffled + accounting.keptByTool.length + accounting.keptByUser.length === accounting.total,
        `${what}: ${accounting.total - accounting.shuffled - accounting.keptByTool.length - accounting.keptByUser.length} question(s) unaccounted for`,
      );
      yield* say(
        accounting.keptByTool.every((n) => !accounting.keptByUser.includes(n)),
        `${what}: a question is counted as kept twice`,
      );
      yield* say(accounting.summary?.endsWith(`= ${accounting.total}`) === true, `${what}: the summary does not add up to the paper`);
    } else {
      yield* say(
        accounting.shuffled === 0 && accounting.keptByTool.length === 0 && accounting.keptByUser.length === 0,
        `${what}: something is counted as shuffled although the shuffle is off`,
      );
      yield* say(accounting.summary === undefined, `${what}: a sum is shown although the shuffle is off`);
    }
    for (const n of [...accounting.keptByTool, ...accounting.keptByUser]) {
      yield* say(paper.questionNumbers.includes(n), `${what}: question ${n} is not in this paper`);
    }
  }

  // Finding 1: the questions that could not be read.
  const skipped = ascending(report.optionsKeptByTool.map((item) => item.questionNumber));
  yield* say(
    skipped.length === report.optionsKeptByTool.length,
    'finding 1 lists the same question twice',
  );
  yield* say(
    !report.shuffleOptions || same(skipped, [...report.optionAccounting.keptByTool]),
    'finding 1 and the headline disagree about what could not be read',
  );
  yield* say(
    same(ascending(paper.unshufflableGroups.flatMap((g) => g.questionNumbers)), skipped),
    'the grouped view of finding 1 loses or invents a question',
  );
  for (const group of paper.unshufflableGroups) {
    yield* say(
      same([...group.questionNumbers], ascending(group.questions.map((q) => q.questionNumber))),
      `"${group.label}" lists different questions from the ones it holds`,
    );
    yield* say(group.label !== group.reason, `finding 1 shows the internal name "${group.reason}"`);
    yield* say(group.fix.length > 20, `"${group.label}" has no usable fix`);
    yield* say(
      (new Set(group.questions.map((q) => q.detail)).size === 1) === (group.sharedDetail !== undefined),
      `"${group.label}" collapses details that are not all the same, or fails to collapse ones that are`,
    );
  }

  // Finding 2: read, but worth correcting.
  const noted = ascending(paper.layoutNotes.map((note) => note.questionNumber));
  yield* say(
    same(ascending(paper.layoutNoteGroups.flatMap((g) => g.questionNumbers)), noted),
    'the grouped view of finding 2 loses or invents a question',
  );
  for (const n of noted) {
    yield* say(paper.questionNumbers.includes(n), `finding 2 mentions question ${n}, which is not in this paper`);
  }
  yield* say(
    noted.every((n) => !skipped.includes(n)),
    'a question is reported both as unreadable and as read',
  );

  // Finding 3: read fine, but position-dependent.
  for (const item of report.suggestedForExclusion) {
    yield* say(paper.questionNumbers.includes(item.questionNumber), `finding 3 mentions question ${item.questionNumber}, which is not in this paper`);
    yield* say(!skipped.includes(item.questionNumber), `finding 3 suggests keeping question ${item.questionNumber}, which is already kept`);
    yield* say(
      !(overrides.optionExclusions ?? []).includes(item.questionNumber),
      `finding 3 suggests question ${item.questionNumber}, which the user has already listed`,
    );
  }

  // Finding 4: held in place by a picture. One question can be held by two pictures, so it
  // appears in two groups - the count of questions is the distinct one.
  const pinned = ascending(report.questionsKeptByTool.flatMap((g) => g.questionNumbers));
  for (const n of pinned) {
    yield* say(paper.questionNumbers.includes(n), `finding 4 mentions question ${n}, which is not in this paper`);
  }
  yield* say(
    !report.shuffleQuestions || same(pinned, [...report.questionAccounting.keptByTool]),
    'finding 4 and the headline disagree about what is held in place',
  );

  // Warnings: every number that does nothing is named, and only those.
  const named = (report.warnings.join(' ').match(/\d+/g) ?? []).map(Number);
  const foreign = ascending(
    [...(overrides.questionExclusions ?? []), ...(overrides.optionExclusions ?? [])].filter(
      (n) => !paper.questionNumbers.includes(n),
    ),
  );
  for (const n of foreign) {
    yield* say(named.includes(n), `question ${n} is not in this paper but nothing says so`);
  }
}

await main();
