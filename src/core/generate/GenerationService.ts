import { promises as fs } from 'node:fs';
import path from 'node:path';
import type {
  DryRunReport,
  DryRunGroup,
  GeneratedSet,
  GenerationRequest,
  GenerationResult,
  PaperSummary,
  ProgressEvent,
  QuestionLayoutNote,
  SkippedOptionShuffle,
  GroupSummary,
} from '../../shared/types';
import {
  OPTION_LABELS,
  QUESTION_LABELS,
  accountFor,
  staleListWarning,
  unknownNumbers,
} from '../../shared/accounting';
import { groupLayoutNotes } from '../../shared/layoutNotes';
import { groupSkippedOptions } from '../../shared/skipReasons';
import { DocxPackage } from '../docx/DocxPackage';
import type { IOptionSetParser } from '../options/OptionSet';
import { OptionSetParser } from '../options/OptionSetParser';
import { NumberingIndex } from '../parse/NumberingIndex';
import { PaperParser } from '../parse/PaperParser';
import { pinnedQuestionNumbers, questionsPinnedByPictures } from '../parse/BoundaryPictures';
import { FALLBACK_SUBJECT, type ParsedPaper, type PaperSection } from '../parse/PaperModel';
import { ShufflePlanner, resolveSeed } from '../shuffle/ShufflePlanner';
import { SetVerifier } from '../verify/SetVerifier';
import { OptionAdvisor } from './OptionAdvisor';
import {
  MAX_SETS,
  MIN_SETS,
  OutputFolderResolver,
  setDisplayName,
  setFileName,
  setLabel,
} from './OutputFolder';
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

    // Both shuffles are split the same way, once, and everything the report says is counted
    // off these four sets - so the section table, the totals and the findings cannot drift
    // apart, and no question is ever counted twice or left out.
    const questionAccounting = accountFor({
      questionNumbers: summary.questionNumbers,
      active: request.shuffleQuestions,
      keptByTool: pinnedQuestionNumbers(summary.pinnedQuestions),
      keptByUser: request.questionExclusions,
      labels: QUESTION_LABELS,
    });
    const optionAccounting = accountFor({
      questionNumbers: summary.questionNumbers,
      active: request.shuffleOptions,
      keptByTool: summary.unshufflableOptions.map((item) => item.questionNumber),
      keptByUser: request.optionExclusions,
      labels: OPTION_LABELS,
    });
    const questionsKept = new Set([...questionAccounting.keptByTool, ...questionAccounting.keptByUser]);
    const optionsKept = new Set([...optionAccounting.keptByTool, ...optionAccounting.keptByUser]);
    const alreadyLeftAlone = new Set([
      ...summary.unshufflableOptions.map((item) => item.questionNumber),
      ...request.optionExclusions,
    ]);

    const groups: DryRunGroup[] = withQuestions(paper).map((section) => {
      const numbers = section.blocks.map((block) => block.printedNumber);
      return {
        group: section.label,
        questionCount: numbers.length,
        movable: request.shuffleQuestions ? numbers.filter((n) => !questionsKept.has(n)).length : 0,
        optionsShuffled: request.shuffleOptions
          ? numbers.filter((n) => !optionsKept.has(n)).length
          : 0,
      };
    });

    const warnings: string[] = [];
    const unknownQuestions = unknownNumbers(summary.questionNumbers, request.questionExclusions);
    const unknownOptions = unknownNumbers(summary.questionNumbers, request.optionExclusions);
    if (unknownQuestions.length > 0) {
      warnings.push(staleListWarning(unknownQuestions, 'keep these question numbers in place'));
    }
    if (unknownOptions.length > 0) {
      warnings.push(staleListWarning(unknownOptions, 'keep the option order of these questions'));
    }
    if (paper.sections.every((section) => section.subject === FALLBACK_SUBJECT)) {
      warnings.push(
        groups.length === 1
          ? 'No subject headings were recognised, so the whole paper is treated as one subject ' +
            'and questions can move anywhere in it.'
          : 'No subject headings were recognised, so the whole paper is treated as one subject. ' +
            `Its ${groups.length} sections still shuffle separately.`,
      );
    }
    for (const group of groups) {
      if (request.shuffleQuestions && group.movable === 1) {
        warnings.push(`Only one question is free to move in ${group.group}, so its order cannot change.`);
      }
    }

    return {
      paper: summary,
      setCount: request.setCount,
      shuffleQuestions: request.shuffleQuestions,
      shuffleOptions: request.shuffleOptions,
      groups,
      questionsEligibleToMove: groups.reduce((sum, group) => sum + group.movable, 0),
      optionsToShuffle: groups.reduce((sum, group) => sum + group.optionsShuffled, 0),
      questionAccounting,
      optionAccounting,
      optionsKeptByTool: summary.unshufflableOptions,
      // Only worth saying when questions were going to move at all.
      questionsKeptByTool: request.shuffleQuestions ? summary.pinnedQuestions : [],
      // No point suggesting a question whose options are already left alone - whether the
      // user listed it or the tool cannot read it. Asked independently of whether options
      // are being shuffled at all, so switching the shuffle off does not resurrect
      // suggestions the user has already acted on.
      suggestedForExclusion: summary.advisories.filter(
        (item) =>
          !alreadyLeftAlone.has(item.questionNumber),
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

    // Opt-out rather than opt-in: an unsplit question is the sane default for a printed
    // paper, and older callers (the CLI, saved settings) omit the flag entirely.
    const keepQuestionsWhole = request.keepQuestionsWhole !== false;

    onProgress({ stage: 'planning', message: 'Planning the sets...', fraction: 0 });
    // Resolved here (not inside the planner) so the exact seed can be reported back.
    const runSeed = resolveSeed(request.seed);
    // Read once, so the folder and every set of this run carry the same date and time in
    // their names even if the run crosses a minute boundary.
    const runStartedAt = new Date();
    // A question pinned by a picture anchored across its boundary is excluded from the
    // re-ordering exactly as if the user had listed it, so the planner needs no new concept.
    const planRequest: GenerationRequest = {
      ...request,
      questionExclusions: [
        ...new Set([...request.questionExclusions, ...pinnedQuestionNumbers(summary.pinnedQuestions)]),
      ],
    };
    const plans = this.planner.plan({ sections, shufflableOptionQuestions, request: planRequest, baseSeed: runSeed });
    const folder = await this.folders.create(request.sourceFile, runStartedAt);

    const setCount = plans.length;
    /** Turns "x% of set n" into "y% of the whole run". */
    const emit = (stage: ProgressEvent['stage'], setNumber: number, setFraction: number, message: string): void =>
      onProgress({
        stage,
        message,
        setNumber,
        setLabel: setDisplayName(setNumber),
        setCount,
        setFraction,
        fraction: (setNumber - 1 + setFraction) / setCount,
      });

    const sets: GeneratedSet[] = [];
    for (const plan of plans) {
      const label = setDisplayName(plan.setNumber);
      // The builder reports 0..1 for its own work, which is ~85% of a set.
      const BUILD_SHARE = 0.85;
      emit('options', plan.setNumber, 0, `Building ${label}`);

      const built = await this.builder.build({
        sourceBuffer: buffer,
        plan,
        originalAnswers: facts.answers,
        setLabel: setLabel(plan.setNumber),
        keepQuestionsWhole,
        onStep: (fraction, message) =>
          emit(fraction < 0.55 ? 'options' : fraction < 0.95 ? 'ordering' : 'packaging', plan.setNumber, fraction * BUILD_SHARE, message),
      });

      const fileName = setFileName(request.sourceFile, plan.setNumber, runStartedAt);
      const filePath = path.join(folder, fileName);
      emit('writing', plan.setNumber, 0.88, `Writing ${fileName}`);
      await fs.writeFile(filePath, built.buffer);

      emit('verifying', plan.setNumber, 0.92, `Verifying ${label}`);
      const verification = await this.verifier.verify(built.buffer, facts);
      emit('verifying', plan.setNumber, 1, `${label} verified`);

      sets.push({
        setNumber: plan.setNumber,
        label,
        fileName,
        filePath,
        seed: plan.seed,
        questionsMoved: built.questionsMoved,
        optionsShuffled: built.optionsShuffled,
        questionsKeptWhole: built.questionsKeptWhole,
        skipped: [...summary.unshufflableOptions, ...built.skipped],
        mappings: built.mappings,
        verification,
      });
    }

    const partial = {
      outputFolder: folder,
      seed: runSeed,
      generatedAt: runStartedAt.toISOString(),
      paper: summary,
      sets,
    };
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
    // At least two, because one set is not a set of anything; at most 26, because the sets
    // are named by letter and that is how many letters there are.
    if (
      !Number.isInteger(request.setCount) ||
      request.setCount < MIN_SETS ||
      request.setCount > MAX_SETS
    ) {
      throw new Error(
        `Enter a whole number of sets between ${MIN_SETS} and ${MAX_SETS}. One set on its ` +
          'own is just a copy of your paper, and the sets are named Set A to Set Z, so ' +
          `${MAX_SETS} is as many as can be named.`,
      );
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
    const paper = this.parser.parsePart(pkg.documentPart(), numbering);
    // The key states how this paper names an option, so labels written the same way are
    // looked for before any other scheme.
    return { paper, optionParser: new OptionSetParser(numbering, paper.answerKey.style.scheme) };
  }

  private summarise(sourceFile: string, paper: ParsedPaper, optionParser: IOptionSetParser): PaperSummary {
    const questionNumbers = paper.sections.flatMap((section) =>
      section.blocks.map((block) => block.printedNumber),
    );
    const groups: GroupSummary[] = withQuestions(paper).map((section) => {
      const numbers = section.blocks.map((block) => block.printedNumber);
      return {
        group: section.label,
        firstQuestionNumber: numbers[0] ?? 0,
        lastQuestionNumber: numbers[numbers.length - 1] ?? 0,
        questionCount: numbers.length,
      };
    });

    const pinnedQuestions = questionsPinnedByPictures(paper);
    const unshufflableOptions: SkippedOptionShuffle[] = [];
    const layoutNotes: QuestionLayoutNote[] = [];
    for (const section of paper.sections) {
      for (const block of section.blocks) {
        const parsed = optionParser.parse(block);
        if (parsed.ok) {
          for (const note of parsed.notes) {
            layoutNotes.push({ questionNumber: block.printedNumber, subject: section.subject, ...note });
          }
          continue;
        }
        unshufflableOptions.push({
          questionNumber: block.printedNumber,
          subject: section.subject,
          reason: parsed.reason,
          detail: parsed.detail,
          fix: parsed.fix,
        });
      }
    }

    return {
      sourceFile,
      questionCount: paper.questionCount,
      questionNumbers,
      groups,
      answerStyle: paper.answerKey.style,
      unshufflableOptions,
      unshufflableGroups: groupSkippedOptions(unshufflableOptions),
      pinnedQuestions,
      // The advisor reads the options themselves, so it needs the same parser.
      advisories: this.advisor.advise(paper, optionParser),
      layoutNotes,
      layoutNoteGroups: groupLayoutNotes(layoutNotes),
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

/**
 * The runs of questions worth showing the user: every section that actually holds
 * questions.
 *
 * A subject heading with nothing under it is real - "BIOLOGY" above "BOTANY" and "ZOOLOGY"
 * in one sample paper - and its nodes still have to be written out in place, so the section
 * exists. Listing it as "BIOLOGY: 0 questions" would only puzzle whoever reads the report.
 */
function withQuestions(paper: ParsedPaper): PaperSection[] {
  return paper.sections.filter((section) => section.blocks.length > 0);
}
