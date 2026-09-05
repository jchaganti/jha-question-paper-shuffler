/**
 * Headless entry point, useful for batch runs and for testing against real papers:
 *
 *   node dist/cli/cli.js --file "paper.docx" --sets 3 --shuffle-questions --shuffle-options \
 *        --keep-question-positions 10,11 --keep-option-order 1,4,8 --seed demo
 *
 * Questions are kept whole on one page unless --allow-page-splits is passed.
 *   node dist/cli/cli.js --file "paper.docx" --inspect
 */
import { GenerationService, parseNumberList } from '../core/generate/GenerationService';
import { questionsWithLayoutNotes } from '../shared/layoutNotes';
import type { GenerationRequest } from '../shared/types';

interface Args {
  readonly [key: string]: string | boolean;
}

function parseArgs(argv: readonly string[]): Args {
  const out: Record<string, string | boolean> = {};
  for (let i = 0; i < argv.length; i++) {
    const token = argv[i]!;
    if (!token.startsWith('--')) continue;
    const key = token.slice(2);
    const next = argv[i + 1];
    if (next === undefined || next.startsWith('--')) out[key] = true;
    else {
      out[key] = next;
      i++;
    }
  }
  return out;
}

async function main(): Promise<void> {
  const args = parseArgs(process.argv.slice(2));
  const file = typeof args.file === 'string' ? args.file : '';
  if (!file) {
    console.error('Usage: --file <paper.docx> [--sets N (2-100, default 2)] [--shuffle-questions] [--shuffle-options]');
    console.error('       [--keep-question-positions 1,2] [--keep-option-order 3,4] [--seed text]');
    console.error('       [--allow-page-splits] [--inspect] [--dry-run]');
    process.exitCode = 2;
    return;
  }

  const service = new GenerationService();

  if (args.inspect) {
    const summary = await service.inspect(file);
    console.log(`Questions: ${summary.questionCount}`);
    for (const group of summary.groups) {
      console.log(
        `  ${group.group}: ${group.questionCount} questions (${group.firstQuestionNumber}-${group.lastQuestionNumber})`,
      );
    }
    console.log(`\nOptions that cannot be shuffled (${summary.unshufflableOptions.length}):`);
    for (const group of summary.unshufflableGroups) {
      console.log(`  ${group.label} - questions ${group.questionNumbers.join(', ')}`);
      console.log(`    fix: ${group.fix}`);
    }
    console.log(`\nReview suggested (${summary.advisories.length}):`);
    for (const item of summary.advisories) {
      console.log(`  Q${item.questionNumber} [${item.subject}] ${item.kind}: ${item.detail}`);
    }
    console.log(`\nShuffled, but worth correcting in the Word document (${summary.layoutNotes.length}):`);
    for (const item of summary.layoutNotes) {
      console.log(`  Q${item.questionNumber} [${item.subject}] ${item.issue}: ${item.detail}`);
    }
    return;
  }

  const request: GenerationRequest = {
    sourceFile: file,
    setCount: typeof args.sets === 'string' ? Number(args.sets) : 2,
    shuffleQuestions: args['shuffle-questions'] === true,
    shuffleOptions: args['shuffle-options'] === true,
    questionExclusions: parseNumberList(typeof args['keep-question-positions'] === 'string' ? args['keep-question-positions'] : ''),
    optionExclusions: parseNumberList(typeof args['keep-option-order'] === 'string' ? args['keep-option-order'] : ''),
    seed: typeof args.seed === 'string' ? args.seed : undefined,
    keepQuestionsWhole: args['allow-page-splits'] !== true,
  };

  if (args['dry-run']) {
    const report = await service.dryRun(request);
    console.log(`Dry run - nothing will be written.\n`);
    console.log(`Sets to generate: ${report.setCount}`);
    console.log(`Shuffle questions: ${report.shuffleQuestions ? 'yes' : 'no'}   Shuffle options: ${report.shuffleOptions ? 'yes' : 'no'}`);
    // Wide enough for "CHEMISTRY - PART 2 SECTION A": a subject divided into sections is
    // reported one row per section, because that is what shuffles as a unit.
    const groupWidth = Math.max(7, ...report.groups.map((group) => group.group.length));
    console.log(
      `\n${'Section'.padEnd(groupWidth)} ${'Questions'.padStart(9)} ${'Free to move'.padStart(13)} ${'Options shuffled'.padStart(17)}`,
    );
    for (const group of report.groups) {
      console.log(
        `${group.group.padEnd(groupWidth)} ${String(group.questionCount).padStart(9)} ` +
          `${String(group.movable).padStart(13)} ${String(group.optionsShuffled).padStart(17)}`,
      );
    }
    // Both totals are shown as arithmetic that adds up to the paper, so a question the
    // report does not otherwise mention is never left unaccounted for.
    console.log(`\nTotals, out of ${report.paper.questionCount} question(s) in this paper:`);
    console.log(`  Questions: ${report.questionAccounting.summary ?? 'not being shuffled.'}`);
    console.log(`  Options:   ${report.optionAccounting.summary ?? 'not being shuffled.'}`);
    if (report.questionAccounting.keptByUser.length > 0) {
      console.log(`  You asked to keep in place: ${report.questionAccounting.keptByUser.join(', ')}`);
    }
    if (report.optionAccounting.keptByUser.length > 0) {
      console.log(`  You asked to keep in order: ${report.optionAccounting.keptByUser.join(', ')}`);
    }
    // Three sections, in this order and always shown, so the same three questions are
    // answered every run: what could not be read, what was read but is worth tidying, and
    // what was read fine but may not mean the same once its options move.
    console.log(
      `\n1. ${report.optionsKeptByTool.length} question(s) whose options cannot be shuffled with certainty.`,
    );
    console.log('   These questions could not be read, so they keep their original option order');
    console.log('   and their answer is unchanged. Each problem is listed once, with the questions');
    console.log('   it affects and what to change in Word.');
    // Grouped by problem: a paper typed one way goes wrong the same way many times over, and
    // the fix is worth reading once rather than once per question.
    for (const group of report.paper.unshufflableGroups) {
      console.log(`\n   ${group.label} (${group.questionNumbers.length})`);
      console.log(`     questions: ${group.questionNumbers.join(', ')}`);
      // When every question in the group says the same thing, say it once.
      if (group.sharedDetail) console.log(`     ${group.sharedDetail}`);
      else {
        for (const item of group.questions) {
          console.log(`       Q${item.questionNumber} [${item.subject}] ${item.detail}`);
        }
      }
      console.log(`     fix: ${group.fix}`);
    }

    const groups = report.paper.layoutNoteGroups;
    const affected = questionsWithLayoutNotes(report.paper.layoutNotes);
    console.log(`\n2. ${affected.length} question(s) shuffled, but worth correcting in the Word document.`);
    console.log('   These were shuffled correctly; their layout had to be worked out, so correcting');
    console.log('   the source removes the guesswork next time.');
    for (const group of groups) {
      console.log(`     ${group.label}`);
      console.log(`       questions: ${group.questionNumbers.join(', ')}`);
      console.log(`       fix: ${group.fix}`);
    }

    console.log(
      `\n3. ${report.suggestedForExclusion.length} question(s) worth keeping in their original option order.`,
    );
    console.log('   Their option text looks position-dependent, so shuffling may change what it means.');
    for (const item of report.suggestedForExclusion) {
      console.log(`     Q${item.questionNumber} [${item.subject}] ${item.kind}: ${item.detail}`);
    }
    if (report.suggestedForExclusion.length > 0) {
      console.log(`\n     --keep-option-order ${report.suggestedForExclusion.map((i) => i.questionNumber).join(',')}`);
    }

    // A fourth finding, about question *position* rather than options: these questions are
    // held where they are so that a picture anchored across their boundary stays with the
    // question it illustrates.
    console.log(
      `\n4. ${report.questionsKeptByTool.length} picture(s) anchored between two questions, holding them in place.`,
    );
    console.log('   These questions keep their original positions; everything else shuffles around them.');
    for (const group of report.questionsKeptByTool) {
      console.log(`     Q${group.questionNumbers.join(', Q')} [${group.subject}] ${group.detail}`);
      console.log(`       fix: ${group.fix}`);
    }
    for (const warning of report.warnings) console.log(`\nWARNING: ${warning}`);
    return;
  }

  const result = await service.generate(request, (event) => {
    const percent = event.fraction === undefined ? '' : ` ${Math.round(event.fraction * 100)}%`;
    const set = event.setNumber ? ` set ${event.setNumber}/${event.setCount}` : '';
    console.log(`  [${event.stage}]${set}${percent} ${event.message}`);
  });
  console.log(`\nOutput folder: ${result.outputFolder}`);
  console.log(`Seed: ${result.seed}   (re-run with --seed "${result.seed}" to reproduce these sets)`);
  for (const set of result.sets) {
    console.log(
      `  ${set.fileName}: moved ${set.questionsMoved} questions, shuffled options of ${set.optionsShuffled} - ` +
        `verification ${set.verification.ok ? 'PASSED' : 'FAILED'}`,
    );
    for (const check of set.verification.checks) {
      if (!check.ok) console.log(`      FAILED: ${check.name} -> ${check.detail}`);
    }
  }
  console.log(`  report: ${result.reportFile}`);
}

main().catch((error: unknown) => {
  console.error(error instanceof Error ? `${error.name}: ${error.message}` : String(error));
  process.exitCode = 1;
});
