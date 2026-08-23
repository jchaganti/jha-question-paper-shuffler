import type {
  DryRunReport,
  GenerationRequest,
  GenerationResult,
  PaperSummary,
  ProgressEvent,
  ShufflerApi,
} from '../shared/types';

declare global {
  interface Window {
    readonly shuffler: ShufflerApi;
  }
}

const api = window.shuffler;

const el = <T extends HTMLElement>(id: string): T => {
  const node = document.getElementById(id);
  if (!node) throw new Error(`Missing element #${id}`);
  return node as T;
};

const sourceFileInput = el<HTMLInputElement>('source-file');
const browseButton = el<HTMLButtonElement>('browse');
const summaryPanel = el<HTMLElement>('paper-summary');
const shuffleQuestions = el<HTMLInputElement>('shuffle-questions');
const shuffleOptions = el<HTMLInputElement>('shuffle-options');
const questionExclusions = el<HTMLInputElement>('question-exclusions');
const optionExclusions = el<HTMLInputElement>('option-exclusions');
const setCount = el<HTMLInputElement>('set-count');
const seed = el<HTMLInputElement>('seed');
const generateButton = el<HTMLButtonElement>('generate');
const dryRunButton = el<HTMLButtonElement>('dry-run');
const statusLabel = el<HTMLElement>('status');
const dryRunPanel = el<HTMLElement>('dry-run-panel');
const resultsPanel = el<HTMLElement>('results');

const progressPanel = el<HTMLElement>('progress-panel');
const progressSet = el<HTMLElement>('progress-set');
const progressOverallBar = el<HTMLElement>('progress-overall-bar');
const progressOverallPercent = el<HTMLElement>('progress-overall-percent');
const progressStep = el<HTMLElement>('progress-step');
const progressStepBar = el<HTMLElement>('progress-step-bar');
const progressStepPercent = el<HTMLElement>('progress-step-percent');

let busy = false;

function setStatus(message: string, isError = false): void {
  statusLabel.textContent = message;
  statusLabel.classList.toggle('status--error', isError);
}

function refreshEnabled(): void {
  const ready = !!sourceFileInput.value && (shuffleQuestions.checked || shuffleOptions.checked);
  generateButton.disabled = busy || !ready;
  dryRunButton.disabled = busy || !ready;
  questionExclusions.disabled = !shuffleQuestions.checked;
  optionExclusions.disabled = !shuffleOptions.checked;
}

function element(tag: string, className: string, content = ''): HTMLElement {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (content) node.textContent = content;
  return node;
}

function table(headers: readonly string[], rows: readonly (readonly string[])[]): HTMLTableElement {
  const node = document.createElement('table');
  const head = node.createTHead().insertRow();
  for (const header of headers) head.appendChild(element('th', '', header));
  const body = node.createTBody();
  for (const row of rows) {
    const tr = body.insertRow();
    for (const cell of row) tr.appendChild(element('td', '', cell));
  }
  return node;
}

function currentRequest(): GenerationRequest {
  return {
    sourceFile: sourceFileInput.value,
    shuffleQuestions: shuffleQuestions.checked,
    shuffleOptions: shuffleOptions.checked,
    questionExclusions: parseNumbers(questionExclusions.value),
    optionExclusions: parseNumbers(optionExclusions.value),
    setCount: Number(setCount.value),
    seed: seed.value.trim() || undefined,
  };
}

/** Merges question numbers into an exclusion field without losing what is already there. */
function addToField(field: HTMLInputElement, numbers: readonly number[]): void {
  const merged = [...new Set([...parseNumbers(field.value), ...numbers])].sort((a, b) => a - b);
  field.value = merged.join(', ');
}

// --- paper structure ------------------------------------------------------------------

function renderSummary(paper: PaperSummary): void {
  summaryPanel.hidden = false;
  summaryPanel.replaceChildren(
    element('p', 'hint', `${paper.questionCount} questions found.`),
    table(
      ['Subject', 'Questions', 'Numbers'],
      paper.subjects.map((subject) => [
        subject.subject,
        String(subject.questionCount),
        `${subject.firstQuestionNumber}–${subject.lastQuestionNumber}`,
      ]),
    ),
    element(
      'p',
      'hint',
      'Use Dry run to see exactly what would be shuffled before any file is written.',
    ),
  );
}

// --- dry run --------------------------------------------------------------------------

function renderDryRun(report: DryRunReport): void {
  dryRunPanel.hidden = false;
  dryRunPanel.replaceChildren(element('h2', '', 'Dry run · nothing was written'));

  const headline = element('p', 'headline');
  const questions = element('strong', '', String(report.questionsEligibleToMove));
  const options = element('strong', '', String(report.optionsToShuffle));
  headline.append(
    'With these settings, ',
    questions,
    ` of ${report.paper.questionCount} questions are free to move and options would be shuffled for `,
    options,
    ` questions, in each of the ${report.setCount} set(s).`,
  );
  dryRunPanel.append(headline);

  dryRunPanel.append(
    table(
      ['Subject', 'Questions', 'Free to move', 'Options shuffled'],
      report.subjects.map((subject) => [
        subject.subject,
        String(subject.questionCount),
        String(subject.movable),
        String(subject.optionsShuffled),
      ]),
    ),
  );

  for (const warning of report.warnings) {
    const notice = element('div', 'notice');
    notice.append(element('strong', '', 'Check this'), element('span', '', warning));
    dryRunPanel.append(notice);
  }

  if (report.optionsKeptByTool.length > 0) {
    const details = document.createElement('details');
    details.append(
      element(
        'summary',
        '',
        `${report.optionsKeptByTool.length} question(s) whose options cannot be shuffled with certainty`,
      ),
      element(
        'p',
        'hint',
        'These keep their original option order in every set, and their answer letter never changes. ' +
          'No action needed - they are listed so you know which ones were left alone. To avoid this in ' +
          'the next paper, see "How to write the Word document" at the top of this window.',
      ),
    );
    const list = document.createElement('ul');
    for (const item of report.optionsKeptByTool) {
      list.append(element('li', '', `Q${item.questionNumber} (${item.subject}) — ${item.detail}`));
    }
    details.append(list);
    dryRunPanel.append(details);
  }

  if (report.optionsKeptByUser.length > 0) {
    dryRunPanel.append(
      element(
        'p',
        'hint number-list',
        `Options kept because you asked: ${report.optionsKeptByUser.join(', ')}`,
      ),
    );
  }

  if (report.suggestedForExclusion.length > 0) {
    const notice = element('div', 'notice');
    notice.append(
      element(
        'strong',
        '',
        `${report.suggestedForExclusion.length} question(s) worth keeping in their original option order`,
      ),
      element(
        'span',
        '',
        'Their option text depends on its position, so shuffling can change the meaning ' +
          '("None of these", "Both (A) and (B)", assertion-reason sets).',
      ),
    );

    const list = document.createElement('ul');
    for (const item of report.suggestedForExclusion) {
      list.append(element('li', '', `Q${item.questionNumber} (${item.subject}) — ${item.detail}`));
    }
    const details = document.createElement('details');
    details.append(element('summary', '', 'Show the list'), list);
    notice.append(details);

    const addButton = document.createElement('button');
    addButton.type = 'button';
    addButton.className = 'button button--secondary';
    addButton.style.marginTop = '10px';
    addButton.textContent = `Add these ${report.suggestedForExclusion.length} to "keep the option order"`;
    addButton.addEventListener('click', () => {
      addToField(
        optionExclusions,
        report.suggestedForExclusion.map((item) => item.questionNumber),
      );
      shuffleOptions.checked = true;
      refreshEnabled();
      addButton.disabled = true;
      addButton.textContent = 'Added — run Dry run again to confirm';
      setStatus('Exclusion list updated.');
    });
    notice.append(document.createElement('br'), addButton);
    dryRunPanel.append(notice);
  } else if (report.shuffleOptions) {
    dryRunPanel.append(
      element('p', 'hint', 'No further questions look position-dependent. Nothing else to exclude.'),
    );
  }
}

// --- progress -------------------------------------------------------------------------

function resetProgress(visible: boolean): void {
  progressPanel.hidden = !visible;
  progressSet.textContent = 'Preparing…';
  progressStep.textContent = ' ';
  progressStepPercent.textContent = '';
  progressOverallPercent.textContent = '0%';
  progressOverallBar.style.width = '0%';
  progressStepBar.style.width = '0%';
}

function renderProgress(event: ProgressEvent): void {
  if (progressPanel.hidden) return;

  const overall = Math.round((event.fraction ?? 0) * 100);
  progressOverallPercent.textContent = `${overall}%`;
  progressOverallBar.style.width = `${overall}%`;

  if (event.setNumber && event.setCount) {
    const done = event.stage === 'done' ? event.setCount : event.setNumber - 1;
    progressSet.textContent =
      event.stage === 'done'
        ? `${event.setCount} of ${event.setCount} sets generated`
        : `Set ${event.setNumber} of ${event.setCount} · ${done} generated`;
  } else {
    progressSet.textContent = event.message;
  }

  if (event.setFraction !== undefined) {
    const step = Math.round(event.setFraction * 100);
    progressStepBar.style.width = `${step}%`;
    progressStepPercent.textContent = `${step}% of this set`;
    progressStep.textContent = event.message;
  } else {
    progressStep.textContent = event.message;
  }
}

// --- results --------------------------------------------------------------------------

function renderResults(result: GenerationResult): void {
  resultsPanel.hidden = false;
  resultsPanel.replaceChildren(element('h2', '', 'Generated sets'));

  const list = document.createElement('ul');
  list.className = 'file-list';
  for (const set of result.sets) {
    const item = document.createElement('li');
    const left = document.createElement('div');
    left.append(
      element('div', 'name', set.fileName),
      element(
        'div',
        'meta',
        `${set.questionsMoved} questions moved · options shuffled for ${set.optionsShuffled} questions`,
      ),
    );
    item.append(
      left,
      element(
        'span',
        `badge ${set.verification.ok ? 'badge--ok' : 'badge--bad'}`,
        set.verification.ok ? 'verified' : 'check failed',
      ),
    );
    list.append(item);
  }
  resultsPanel.append(list);

  const failed = result.sets.filter((set) => !set.verification.ok);
  if (failed.length > 0) {
    const notice = element('div', 'notice');
    notice.append(element('strong', '', 'Verification problems'));
    for (const set of failed) {
      for (const check of set.verification.checks.filter((c) => !c.ok)) {
        notice.append(element('div', '', `Set ${set.setNumber}: ${check.name} — ${check.detail}`));
      }
    }
    resultsPanel.append(notice);
  }

  const seedNotice = element('p', 'hint');
  seedNotice.append('Seed: ', element('code', '', result.seed));
  const reuse = document.createElement('button');
  reuse.type = 'button';
  reuse.className = 'button button--secondary button--inline';
  reuse.textContent = 'Reuse this seed';
  reuse.title = 'Put this seed in the Seed box, so the same settings recreate these exact sets';
  reuse.addEventListener('click', () => {
    seed.value = result.seed;
    setStatus('Seed copied into the Seed box.');
  });
  seedNotice.append(' ', reuse);
  resultsPanel.append(seedNotice);

  const footer = document.createElement('div');
  footer.className = 'results-footer';
  const openButton = document.createElement('button');
  openButton.type = 'button';
  openButton.className = 'button button--secondary';
  openButton.textContent = 'Open output folder';
  openButton.addEventListener('click', () => void api.revealFolder(result.outputFolder));
  footer.append(openButton, element('p', 'hint', result.outputFolder));
  resultsPanel.append(footer);
}

// --- wiring ---------------------------------------------------------------------------

browseButton.addEventListener('click', async () => {
  const file = await api.pickSourceFile();
  if (!file) return;
  sourceFileInput.value = file;
  summaryPanel.hidden = true;
  dryRunPanel.hidden = true;
  resultsPanel.hidden = true;
  resetProgress(false);
  setStatus('Reading the paper…');
  refreshEnabled();

  const result = await api.inspect(file);
  if (!result.ok) {
    setStatus(result.error, true);
    return;
  }
  renderSummary(result.value);
  setStatus('');
});

for (const input of [shuffleQuestions, shuffleOptions]) {
  input.addEventListener('change', refreshEnabled);
}

dryRunButton.addEventListener('click', async () => {
  busy = true;
  refreshEnabled();
  resultsPanel.hidden = true;
  resetProgress(false);
  setStatus('Checking what would happen…');

  const result = await api.dryRun(currentRequest());
  busy = false;
  refreshEnabled();

  if (!result.ok) {
    dryRunPanel.hidden = true;
    setStatus(result.error, true);
    return;
  }
  renderDryRun(result.value);
  setStatus('Dry run finished — no files were written.');
});

generateButton.addEventListener('click', async () => {
  busy = true;
  refreshEnabled();
  resultsPanel.hidden = true;
  resetProgress(true);
  setStatus('');

  const result = await api.generate(currentRequest());

  busy = false;
  refreshEnabled();

  if (!result.ok) {
    resetProgress(false);
    setStatus(result.error, true);
    return;
  }
  setStatus(`Done — ${result.value.sets.length} set(s) created.`);
  renderResults(result.value);
});

api.onProgress(renderProgress);

/** Client-side copy of the exclusion parser: "1, 4, 8-10" -> [1,4,8,9,10]. */
function parseNumbers(input: string): number[] {
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
    if (/^\d+$/.test(token)) out.add(Number(token));
  }
  return [...out].sort((a, b) => a - b);
}

refreshEnabled();
