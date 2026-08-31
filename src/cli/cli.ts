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
    console.error('Usage: --file <paper.docx> [--sets N] [--shuffle-questions] [--shuffle-options]');
    console.error('       [--keep-question-positions 1,2] [--keep-option-order 3,4] [--seed text]');
    console.error('       [--allow-page-splits] [--inspect] [--dry-run]');
    process.exitCode = 2;
    return;
  }

  const service = new GenerationService();

  if (args.inspect) {
    const summary = await service.inspect(file);
    console.log(`Questions: ${summary.questionCount}`);
    for (const subject of summary.subjects) {
      console.log(
        `  ${subject.subject}: ${subject.questionCount} questions (${subject.firstQuestionNumber}-${subject.lastQuestionNumber})`,
      );
    }
    console.log(`\nOptions that cannot be shuffled (${summary.unshufflableOptions.length}):`);
    for (const item of summary.unshufflableOptions) {
      console.log(`  Q${item.questionNumber} [${item.subject}] ${item.reason}: ${item.detail}`);
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
    setCount: typeof args.sets === 'string' ? Number(args.sets) : 1,
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
    console.log(`\n${'Subject'.padEnd(12)} ${'Questions'.padStart(9)} ${'Free to move'.padStart(13)} ${'Options shuffled'.padStart(17)}`);
    for (const subject of report.subjects) {
      console.log(
        `${subject.subject.padEnd(12)} ${String(subject.questionCount).padStart(9)} ` +
          `${String(subject.movable).padStart(13)} ${String(subject.optionsShuffled).padStart(17)}`,
      );
    }
    console.log(
      `\nTotals: ${report.questionsEligibleToMove} question(s) free to move, ` +
        `options shuffled for ${report.optionsToShuffle} question(s).`,
    );
    // Three sections, in this order and always shown, so the same three questions are
    // answered every run: what could not be read, what was read but is worth tidying, and
    // what was read fine but may not mean the same once its options move.
    console.log(
      `\n1. ${report.optionsKeptByTool.length} question(s) whose options cannot be shuffled with certainty.`,
    );
    console.log('   These questions could not be parsed, so they keep their original option order');
    console.log('   and their answer is unchanged.');
    for (const item of report.optionsKeptByTool) {
      console.log(`     Q${item.questionNumber} [${item.subject}] ${item.reason}: ${item.detail}`);
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
    console.log(
      `\nOptions already kept because you asked (${report.optionsKeptByUser.length}): ` +
        `${report.optionsKeptByUser.join(', ') || '-'}`,
    );
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
