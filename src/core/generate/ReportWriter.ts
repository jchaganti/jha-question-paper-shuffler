import { promises as fs } from 'node:fs';
import path from 'node:path';
import { questionsWithLayoutNotes } from '../../shared/layoutNotes';
import { renderAnswer } from '../../shared/answerStyle';
import type { GenerationRequest, GenerationResult } from '../../shared/types';
import type { ReportBlock, ReportDocument, ReportFact } from '../report/ReportDocument';
import { renderReportPdf } from '../report/ReportPdf';
import { readableTimestamp } from './OutputFolder';

export const REPORT_FILE_NAME = '_generation-report.pdf';

/**
 * Writes an audit trail next to the generated sets: what was shuffled, what was left
 * alone and why, the seed needed to reproduce the run, and the full question mapping
 * of every set.
 *
 * `build` produces the report as structure and `renderReportPdf` typesets it, so what the
 * report says is decided here and how it looks is decided in one place next door.
 */
export class ReportWriter {
  async write(
    folder: string,
    request: GenerationRequest,
    result: Omit<GenerationResult, 'reportFile'>,
  ): Promise<string> {
    const file = path.join(folder, REPORT_FILE_NAME);
    await fs.writeFile(file, renderReportPdf(this.build(request, result)));
    return file;
  }

  build(request: GenerationRequest, result: Omit<GenerationResult, 'reportFile'>): ReportDocument {
    const { paper } = result;
    const blocks: ReportBlock[] = [];

    blocks.push({ kind: 'heading', text: 'This run' });
    blocks.push({
      kind: 'facts',
      items: [
        { label: 'Source paper', value: request.sourceFile, mono: true },
        { label: 'Sets', value: String(result.sets.length) },
        {
          label: 'Seed',
          value:
            result.seed +
            (request.seed?.trim() ? '' : '  (generated for this run)'),
          mono: true,
        },
        { label: 'Shuffle questions', value: request.shuffleQuestions ? 'yes' : 'no' },
        { label: 'Shuffle options', value: request.shuffleOptions ? 'yes' : 'no' },
        { label: 'Questions kept in place', value: format(request.questionExclusions) },
        { label: 'Option order kept', value: format(request.optionExclusions) },
        {
          label: 'One question per page',
          value: request.keepQuestionsWhole === false ? 'no' : 'yes',
        },
      ],
    });
    blocks.push({
      kind: 'paragraph',
      text:
        'Enter the seed above with the same settings and the same paper to produce these ' +
        'exact sets again. Keep this report as confidential as the paper itself: it holds ' +
        'the answer mapping of every set.',
    });

    blocks.push({ kind: 'heading', text: 'Paper structure' });
    blocks.push({
      kind: 'table',
      columns: [{ header: 'Subject' }, { header: 'Questions', align: 'right' }, { header: 'Range' }],
      rows: paper.subjects.map((subject) => [
        subject.subject,
        String(subject.questionCount),
        `${subject.firstQuestionNumber}-${subject.lastQuestionNumber}`,
      ]),
    });
    blocks.push({ kind: 'paragraph', text: `Total questions: ${paper.questionCount}` });

    if (paper.unshufflableOptions.length > 0) {
      blocks.push({ kind: 'heading', text: 'Options that could not be shuffled' });
      blocks.push({
        kind: 'paragraph',
        text: 'These questions keep their original option order, and their answer, in every set.',
      });
      blocks.push({
        kind: 'table',
        columns: [
          { header: 'Q', align: 'right' },
          { header: 'Subject' },
          { header: 'Reason', weight: 1 },
        ],
        rows: paper.unshufflableOptions.map((item) => [
          String(item.questionNumber),
          item.subject,
          `${item.reason}: ${item.detail}`,
        ]),
      });
    }

    if (paper.layoutNoteGroups.length > 0) {
      blocks.push({ kind: 'heading', text: 'Shuffled, but worth correcting in the Word document' });
      blocks.push({
        kind: 'paragraph',
        text:
          `${questionsWithLayoutNotes(paper.layoutNotes).length} question(s) were shuffled ` +
          'normally, but their option labels had to be worked out from an inconsistent ' +
          'layout. Correcting the source paper removes the guesswork next time.',
      });
      blocks.push({
        kind: 'table',
        columns: [
          { header: 'Problem', weight: 1 },
          { header: 'Questions', weight: 1 },
          { header: 'Fix', weight: 2 },
        ],
        rows: paper.layoutNoteGroups.map((group) => [
          group.label,
          group.questionNumbers.join(', '),
          group.fix,
        ]),
      });
    }

    if (paper.advisories.length > 0) {
      blocks.push({ kind: 'heading', text: 'Review suggested' });
      blocks.push({
        kind: 'paragraph',
        text:
          'The option order of these questions may carry meaning. Consider adding them to ' +
          'the "keep the option order" list.',
      });
      blocks.push({
        kind: 'table',
        columns: [{ header: 'Q', align: 'right' }, { header: 'Subject' }, { header: 'Why', weight: 1 }],
        rows: paper.advisories.map((item) => [
          String(item.questionNumber),
          item.subject,
          item.detail,
        ]),
      });
    }

    for (const set of result.sets) {
      const label = `Set ${String(set.setNumber).padStart(2, '0')}`;
      blocks.push({ kind: 'heading', text: label });
      const facts: ReportFact[] = [
        { label: 'File', value: set.fileName, mono: true },
        { label: 'Seed for this set', value: `${set.seed}  (derived from the run seed)`, mono: true },
        { label: 'Questions moved', value: String(set.questionsMoved) },
        { label: 'Options shuffled', value: String(set.optionsShuffled) },
      ];
      if (set.questionsKeptWhole > 0) {
        facts.push({ label: 'Kept whole on a page', value: String(set.questionsKeptWhole) });
      }
      facts.push({
        label: 'Verification',
        value: set.verification.ok ? 'PASSED' : 'FAILED',
        tone: set.verification.ok ? 'ok' : 'bad',
      });
      for (const check of set.verification.checks) {
        facts.push({
          label: '',
          value: `${check.ok ? 'ok' : 'FAILED'} - ${check.name} (${check.detail})`,
          tone: check.ok ? undefined : 'bad',
        });
      }
      blocks.push({ kind: 'facts', items: facts });

      blocks.push({ kind: 'heading', text: `${label} - question and answer mapping`, level: 2 });
      // Answers are named the way the paper names them, so this table can be read straight
      // against the generated set's key - never "A" for a paper that writes "(1)".
      blocks.push({
        kind: 'table',
        columns: [
          { header: 'New Q', align: 'right' },
          { header: 'Was Q', align: 'right' },
          { header: 'Original answer' },
          { header: 'New answer' },
          { header: 'Option mapping', weight: 1 },
        ],
        rows: set.mappings.map((mapping) => [
          String(mapping.newNumber),
          String(mapping.originalNumber),
          renderAnswer(mapping.originalAnswer, paper.answerStyle),
          renderAnswer(mapping.newAnswer, paper.answerStyle),
          mapping.optionMapping || '(kept)',
        ]),
      });
    }

    return {
      title: 'Question set generation report',
      // The run's own timestamp, the same one carried in the folder and file names.
      subtitle:
        `${path.basename(request.sourceFile)}   ·   generated ` +
        readableTimestamp(new Date(result.generatedAt)),
      blocks,
    };
  }
}

function format(values: readonly number[]): string {
  return values.length === 0 ? '(none)' : values.join(', ');
}
