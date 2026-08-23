import { promises as fs } from 'node:fs';
import path from 'node:path';
import type {
  DryRunReport,
  DryRunSubject,
  GeneratedSet,
  GenerationRequest,
  GenerationResult,
  PaperSummary,
  ProgressEvent,
  SkippedOptionShuffle,
  SubjectSummary,
} from '../../shared/types';
import { DocxPackage } from '../docx/DocxPackage';
import type { IOptionSetParser } from '../options/OptionSet';
import { OptionSetParser } from '../options/OptionSetParser';
import { NumberingIndex } from '../parse/NumberingIndex';
import { PaperParser } from '../parse/PaperParser';
import type { ParsedPaper } from '../parse/PaperModel';
import { ShufflePlanner, resolveSeed } from '../shuffle/ShufflePlanner';
import { SetVerifier } from '../verify/SetVerifier';
import { OptionAdvisor } from './OptionAdvisor';
import { OutputFolderResolver, setFileName, setLabel } from './OutputFolder';
import { ReportWriter } from './ReportWriter';
import { SetBuilder } from './SetBuilder';

export type ProgressListener = (event: ProgressEvent) => void;

/**
 * Application service that ties parsing, planning, building, verifying and reporting
 * together. The UI and the CLI both drive this class and nothing else.
 */
export class GenerationService {
  constructor(
    private readonly parser: PaperParser = new PaperParser(),
    private readonly planner: ShufflePlanner = new ShufflePlanner(),
    private readonly builder: SetBuilder = new SetBuilder(),
    private readonly verifier: SetVerifier = new SetVerifier(),
    private readonly advisor: OptionAdvisor = new OptionAdvisor(),
    private readonly folders: OutputFolderResolver = new OutputFolderResolver(),
    private readonly reports: ReportWriter = new ReportWriter(),
  ) {}

  /** Parses the paper and reports what the generator can and cannot do with it. */
  async inspect(sourceFile: string): Promise<PaperSummary> {
    const { paper, optionParser } = await this.parse(sourceFile);
    return this.summarise(sourceFile, paper, optionParser);
  }

  /**
   * Answers "what would happen if I pressed Generate now?" without writing anything.
   *
   * Only facts that do not depend on the random seed are reported: which questions are
   * free to move, whose options will be permuted, which questions the tool refuses to
   * touch, which ones the user may want to exclude, and any mistake in the exclusion
   * lists (a question number that does not exist in this paper, for example).
   */
  async dryRun(request: GenerationRequest): Promise<DryRunReport> {
    this.validate(request);
    const { paper, optionParser } = await this.parse(request.sourceFile);
    const summary = this.summarise(request.sourceFile, paper, optionParser);

    const questionExclusions = new Set(request.questionExclusions);
    const optionExclusions = new Set(request.optionExclusions);
    const unshufflable = new Set(summary.unshufflableOptions.map((item) => item.questionNumber));
    const allNumbers = new Set(
      paper.sections.flatMap((section) => section.blocks.map((block) => block.printedNumber)),
    );

    const subjects: DryRunSubject[] = paper.sections.map((section) => {
      const numbers = section.blocks.map((block) => block.printedNumber);
      return {
        subject: section.subject,
        questionCount: numbers.length,
        movable: request.shuffleQuestions ? numbers.filter((n) => !questionExclusions.has(n)).length : 0,
        optionsShuffled: request.shuffleOptions
          ? numbers.filter((n) => !optionExclusions.has(n) && !unshufflable.has(n)).length
          : 0,
      };
    });

    const warnings: string[] = [];
    const unknownQuestions = [...questionExclusions].filter((n) => !allNumbers.has(n));
    const unknownOptions = [...optionExclusions].filter((n) => !allNumbers.has(n));
    if (unknownQuestions.length > 0) {
      warnings.push(
        `This paper has no question ${unknownQuestions.join(', ')}, so those entries in ` +
          '"keep these question numbers in place" will have no effect.',
      );
    }
    if (unknownOptions.length > 0) {
      warnings.push(
        `This paper has no question ${unknownOptions.join(', ')}, so those entries in ` +
          '"keep the option order of these questions" will have no effect.',
      );
    }
    if (paper.sections.length === 1 && paper.sections[0]?.subject === 'ALL') {
      warnings.push(
        'No subject headings were recognised, so the whole paper is treated as one subject ' +
          'and questions can move anywhere in it.',
      );
    }
    for (const subject of subjects) {
      if (request.shuffleQuestions && subject.movable === 1) {
        warnings.push(`Only one question is free to move in ${subject.subject}, so its order cannot change.`);
      }
    }

    return {
      paper: summary,
      setCount: request.setCount,
      shuffleQuestions: request.shuffleQuestions,
      shuffleOptions: request.shuffleOptions,
      subjects,
      questionsEligibleToMove: subjects.reduce((sum, subject) => sum + subject.movable, 0),
      optionsToShuffle: subjects.reduce((sum, subject) => sum + subject.optionsShuffled, 0),
      optionsKeptByUser: [...optionExclusions].filter((n) => allNumbers.has(n)).sort((a, b) => a - b),
      optionsKeptByTool: summary.unshufflableOptions,
      // No point suggesting a question whose options are already left alone.
      suggestedForExclusion: summary.advisories.filter(
        (item) => !optionExclusions.has(item.questionNumber) && !unshufflable.has(item.questionNumber),
      ),
      warnings,
    };
  }

  async generate(request: GenerationRequest, onProgress: ProgressListener = () => {}): Promise<GenerationResult> {
    this.validate(request);

    onProgress({ stage: 'reading', message: 'Reading the question paper...' });
    const buffer = await fs.readFile(request.sourceFile);
    const pkg = await DocxPackage.fromBuffer(buffer);

    onProgress({ stage: 'parsing', message: 'Understanding questions, options and answer key...' });
    const { paper, optionParser } = this.parsePackage(pkg);
    const summary = this.summarise(request.sourceFile, paper, optionParser);
    const facts = SetVerifier.factsOf(paper, optionParser);

    const unshufflable = new Set(summary.unshufflableOptions.map((item) => item.questionNumber));
    const shufflableOptionQuestions: number[] = [];
    const sections: number[][] = paper.sections.map((section) => {
      const numbers: number[] = [];
      for (const block of section.blocks) {
        numbers.push(block.printedNumber);
        if (!unshufflable.has(block.printedNumber)) shufflableOptionQuestions.push(block.printedNumber);
      }
      return numbers;
    });

    onProgress({ stage: 'planning', message: 'Planning the sets...', fraction: 0 });
    // Resolved here (not inside the planner) so the exact seed can be reported back.
    const runSeed = resolveSeed(request.seed);
    const plans = this.planner.plan({ sections, shufflableOptionQuestions, request, baseSeed: runSeed });
    const folder = await this.folders.create(request.sourceFile);

    const setCount = plans.length;
    /** Turns "x% of set n" into "y% of the whole run". */
    const emit = (stage: ProgressEvent['stage'], setNumber: number, setFraction: number, message: string): void =>
      onProgress({
        stage,
        message,
        setNumber,
        setCount,
        setFraction,
        fraction: (setNumber - 1 + setFraction) / setCount,
      });

    const sets: GeneratedSet[] = [];
    for (const plan of plans) {
      const label = `Set ${String(plan.setNumber).padStart(2, '0')}`;
      // The builder reports 0..1 for its own work, which is ~85% of a set.
      const BUILD_SHARE = 0.85;
      emit('options', plan.setNumber, 0, `Building ${label}`);

      const built = await this.builder.build({
        sourceBuffer: buffer,
        plan,
        originalAnswers: facts.answers,
        setLabel: setLabel(plan.setNumber),
        onStep: (fraction, message) =>
          emit(fraction < 0.55 ? 'options' : fraction < 0.95 ? 'ordering' : 'packaging', plan.setNumber, fraction * BUILD_SHARE, message),
      });

      const fileName = setFileName(request.sourceFile, plan.setNumber);
      const filePath = path.join(folder, fileName);
      emit('writing', plan.setNumber, 0.88, `Writing ${fileName}`);
      await fs.writeFile(filePath, built.buffer);

      emit('verifying', plan.setNumber, 0.92, `Verifying ${label}`);
      const verification = await this.verifier.verify(built.buffer, facts);
      emit('verifying', plan.setNumber, 1, `${label} verified`);

      sets.push({
        setNumber: plan.setNumber,
        fileName,
        filePath,
        seed: plan.seed,
        questionsMoved: built.questionsMoved,
        optionsShuffled: built.optionsShuffled,
        skipped: [...summary.unshufflableOptions, ...built.skipped],
        mappings: built.mappings,
        verification,
      });
    }

    const partial = { outputFolder: folder, seed: runSeed, paper: summary, sets };
    const reportFile = await this.reports.write(folder, request, partial);
    onProgress({
      stage: 'done',
      message: `Generated ${sets.length} set(s) in ${folder}`,
      setCount,
      setNumber: setCount,
      setFraction: 1,
      fraction: 1,
    });

    return { ...partial, reportFile };
  }

  private validate(request: GenerationRequest): void {
    if (!request.sourceFile.trim()) throw new Error('Choose the original question paper first.');
    if (path.extname(request.sourceFile).toLowerCase() !== '.docx') {
      throw new Error('The question paper must be a .docx file (Word 2007 or later).');
    }
    if (!Number.isInteger(request.setCount) || request.setCount < 1 || request.setCount > 100) {
      throw new Error('Number of sets must be a whole number between 1 and 100.');
    }
    if (!request.shuffleQuestions && !request.shuffleOptions) {
      throw new Error('Select at least one of "shuffle questions" or "shuffle options".');
    }
  }

  /**
   * Parses a paper and builds the option parser for it. The two travel together because
   * recognising auto-lettered options needs that paper's own numbering definitions.
   */
  private async parse(sourceFile: string): Promise<{ paper: ParsedPaper; optionParser: IOptionSetParser }> {
    const pkg = await DocxPackage.load(sourceFile);
    return this.parsePackage(pkg);
  }

  private parsePackage(pkg: DocxPackage): { paper: ParsedPaper; optionParser: IOptionSetParser } {
    const numbering = new NumberingIndex(pkg.numberingPart());
    return {
      paper: this.parser.parsePart(pkg.documentPart(), numbering),
      optionParser: new OptionSetParser(numbering),
    };
  }

  private summarise(sourceFile: string, paper: ParsedPaper, optionParser: IOptionSetParser): PaperSummary {
    const subjects: SubjectSummary[] = paper.sections.map((section) => {
      const numbers = section.blocks.map((block) => block.printedNumber);
      return {
        subject: section.subject,
        firstQuestionNumber: numbers[0] ?? 0,
        lastQuestionNumber: numbers[numbers.length - 1] ?? 0,
        questionCount: numbers.length,
      };
    });

    const unshufflableOptions: SkippedOptionShuffle[] = [];
    for (const section of paper.sections) {
      for (const block of section.blocks) {
        const parsed = optionParser.parse(block);
        if (parsed.ok) continue;
        unshufflableOptions.push({
          questionNumber: block.printedNumber,
          subject: section.subject,
          reason: parsed.reason,
          detail: parsed.detail,
        });
      }
    }

    return {
      sourceFile,
      questionCount: paper.questionCount,
      subjects,
      unshufflableOptions,
      advisories: this.advisor.advise(paper),
    };
  }
}

/** Parses "1, 4, 8-10" into [1,4,8,9,10]. Ranges are a convenience over the spec. */
export function parseNumberList(input: string): number[] {
  const out = new Set<number>();
  for (const chunk of input.split(/[,;\s]+/)) {
    const token = chunk.trim();
    if (!token) continue;
    const range = /^(\d+)\s*[-–]\s*(\d+)$/.exec(token);
    if (range) {
      const from = Number(range[1]);
      const to = Number(range[2]);
      for (let n = Math.min(from, to); n <= Math.max(from, to); n++) out.add(n);
      continue;
    }
    if (!/^\d+$/.test(token)) throw new Error(`"${token}" is not a question number.`);
    out.add(Number(token));
  }
  return [...out].sort((a, b) => a - b);
}
