import { promises as fs } from 'node:fs';
import path from 'node:path';
import type { GenerationRequest, GenerationResult } from '../../shared/types';

export const REPORT_FILE_NAME = '_generation-report.md';

/**
 * Writes an audit trail next to the generated sets: what was shuffled, what was left
 * alone and why, the seed needed to reproduce the run, and the full question mapping
 * of every set.
 */
export class ReportWriter {
  async write(folder: string, request: GenerationRequest, result: Omit<GenerationResult, 'reportFile'>): Promise<string> {
    const file = path.join(folder, REPORT_FILE_NAME);
    await fs.writeFile(file, this.render(request, result), 'utf8');
    return file;
  }

  render(request: GenerationRequest, result: Omit<GenerationResult, 'reportFile'>): string {
    const lines: string[] = [];
    const { paper } = result;

    lines.push('# Question set generation report', '');
    lines.push(`- Source paper: \`${request.sourceFile}\``);
    lines.push(`- Generated: ${new Date().toISOString()}`);
    lines.push(`- Sets: ${result.sets.length}`);
    lines.push(
      `- Seed: \`${result.seed}\`` +
        (request.seed?.trim()
          ? ''
          : ' (generated for this run)') +
        ' — enter this seed with the same settings to produce these exact same sets again.',
    );
    lines.push(`- Shuffle questions: ${request.shuffleQuestions ? 'yes' : 'no'}`);
    lines.push(`- Shuffle options: ${request.shuffleOptions ? 'yes' : 'no'}`);
    lines.push(`- Questions excluded from re-ordering: ${format(request.questionExclusions)}`);
    lines.push(`- Questions whose options were kept: ${format(request.optionExclusions)}`);
    lines.push('');

    lines.push('## Paper structure', '');
    lines.push('| Subject | Questions | Range |');
    lines.push('| --- | --- | --- |');
    for (const subject of paper.subjects) {
      lines.push(
        `| ${subject.subject} | ${subject.questionCount} | ${subject.firstQuestionNumber}-${subject.lastQuestionNumber} |`,
      );
    }
    lines.push('', `Total questions: ${paper.questionCount}`, '');

    if (paper.unshufflableOptions.length > 0) {
      lines.push('## Options that could not be shuffled', '');
      lines.push('These questions keep their original option order in every set.', '');
      lines.push('| Question | Subject | Reason |');
      lines.push('| --- | --- | --- |');
      for (const item of paper.unshufflableOptions) {
        lines.push(`| ${item.questionNumber} | ${item.subject} | ${item.reason}: ${item.detail} |`);
      }
      lines.push('');
    }

    if (paper.advisories.length > 0) {
      lines.push('## Review suggested', '');
      lines.push(
        'The option order of these questions may carry meaning. Consider adding them to the ' +
          '"options not shuffled" exclusion list.',
        '',
      );
      lines.push('| Question | Subject | Why |');
      lines.push('| --- | --- | --- |');
      for (const item of paper.advisories) {
        lines.push(`| ${item.questionNumber} | ${item.subject} | ${item.detail} |`);
      }
      lines.push('');
    }

    for (const set of result.sets) {
      lines.push(`## Set ${String(set.setNumber).padStart(2, '0')}`, '');
      lines.push(`- File: \`${set.fileName}\``);
      lines.push(`- Seed used for this set: \`${set.seed}\` (derived from the run seed)`);
      lines.push(`- Questions moved: ${set.questionsMoved}`);
      lines.push(`- Questions with shuffled options: ${set.optionsShuffled}`);
      lines.push(`- Verification: ${set.verification.ok ? 'PASSED' : 'FAILED'}`);
      for (const check of set.verification.checks) {
        lines.push(`  - ${check.ok ? 'ok' : 'FAILED'} - ${check.name} (${check.detail})`);
      }
      lines.push('');
      lines.push('| New Q | Was Q | Original answer | New answer | Option mapping |');
      lines.push('| --- | --- | --- | --- | --- |');
      for (const mapping of set.mappings) {
        lines.push(
          `| ${mapping.newNumber} | ${mapping.originalNumber} | ${mapping.originalAnswer} | ${mapping.newAnswer} | ${mapping.optionMapping || '(kept)'} |`,
        );
      }
      lines.push('');
    }

    return lines.join('\n');
  }
}

function format(values: readonly number[]): string {
  return values.length === 0 ? '(none)' : values.join(', ');
}
