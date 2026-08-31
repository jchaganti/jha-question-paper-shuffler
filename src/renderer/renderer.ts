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

/**
 * Page colour.
 *
 * The palettes themselves live in styles.css, one block each, selected by `data-palette`
 * on <html>: the page's Content-Security-Policy allows no inline style, so the colours
 * cannot be written onto the element from here. The choice is remembered per machine;
 * when it cannot be (storage disabled, or the file opened from a data: URL in the UI
 * harness) the app still works and simply opens on the default.
 */
const PALETTES = ['periwinkle', 'mint', 'peach', 'sand'] as const;
const DEFAULT_PALETTE = PALETTES[0];
const PALETTE_KEY = 'shuffler.palette';

const paletteSelect = el<HTMLSelectElement>('palette');

function applyPalette(name: string): void {
  const palette = (PALETTES as readonly string[]).includes(name) ? name : DEFAULT_PALETTE;
  document.documentElement.dataset.palette = palette;
  paletteSelect.value = palette;
}

applyPalette(readStoredPalette());

paletteSelect.addEventListener('change', () => {
  applyPalette(paletteSelect.value);
  try {
    localStorage.setItem(PALETTE_KEY, paletteSelect.value);
  } catch {
    // Storage unavailable: the colour still applies, it just is not remembered.
  }
});

function readStoredPalette(): string {
  try {
    return localStorage.getItem(PALETTE_KEY) ?? DEFAULT_PALETTE;
  } catch {
    return DEFAULT_PALETTE;
  }
}

const sourceFileInput = el<HTMLInputElement>('source-file');
const browseButton = el<HTMLButtonElement>('browse');
const summaryPanel = el<HTMLElement>('paper-summary');
const shuffleQuestions = el<HTMLInputElement>('shuffle-questions');
const shuffleOptions = el<HTMLInputElement>('shuffle-options');
const questionExclusions = el<HTMLInputElement>('question-exclusions');
const optionExclusions = el<HTMLInputElement>('option-exclusions');
const setCount = el<HTMLInputElement>('set-count');
const seed = el<HTMLInputElement>('seed');
const keepQuestionsWhole = el<HTMLInputElement>('keep-questions-whole');
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
    keepQuestionsWhole: keepQuestionsWhole.checked,
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

/**
 * One of the dry run's three numbered findings.
 *
 * All three are rendered the same way and shown every run, even when the count is zero, so
 * the report answers the same three questions in the same order every time: what could not
 * be read, what was read but is worth tidying, and what was read fine but may not mean the
 * same once its options move.
 */
function renderFinding(number: number, heading: string, explanation: string): HTMLDetailsElement {
  const details = document.createElement('details');
  details.className = 'finding';
  details.append(
    element('summary', '', `${number}. ${heading}`),
    element('p', 'hint', explanation),
  );
  dryRunPanel.append(details);
  return details;
}

/**
 * Finding 2: questions that *are* shuffled, but whose option labels had to be worked out
 * from an inconsistent layout. Nothing to do before generating - this is for whoever types
 * the next paper, so it is grouped by problem with the fix stated once.
 *
 * The grouping arrives ready-made from the main process (`PaperSummary.layoutNoteGroups`),
 * so the UI and the CLI cannot drift apart.
 */
function renderLayoutNotes(paper: PaperSummary): void {
  const groups = paper.layoutNoteGroups;
  // A question can carry more than one note, so count questions, not notes.
  const affected = new Set(paper.layoutNotes.map((note) => note.questionNumber)).size;
  const details = renderFinding(
    2,
    `${affected} question(s) shuffled, but worth correcting in the Word document`,
    'These are shuffled correctly and their answers are right. The tool had to work their ' +
      'option labels out from a layout that was not consistent — it can, but it should not ' +
      'have to. Tidying them up in Word removes any doubt for the next paper.',
  );
  if (groups.length === 0) return;

  const list = document.createElement('ul');
  list.className = 'notes-list';
  for (const group of groups) {
    const item = document.createElement('li');
    item.append(
      element('strong', '', group.label),
      element('div', '', `Questions: ${group.questionNumbers.join(', ')}`),
      element('div', 'hint', `Fix: ${group.fix}`),
    );
    list.append(item);
  }
  details.append(list);
}

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

  // 1. Could not be read.
  {
    const details = renderFinding(
      1,
      `${report.optionsKeptByTool.length} question(s) whose options cannot be shuffled with certainty`,
      'These questions could not be parsed, so they keep their original option order in every ' +
        'set and their answer never changes. Each one below says what stopped it and what to ' +
        'change in Word; see also "How to write the Word document" at the top of this window.',
    );
    const list = document.createElement('ul');
    for (const item of report.optionsKeptByTool) {
      list.append(element('li', '', `Q${item.questionNumber} (${item.subject}) — ${item.detail}`));
    }
    if (report.optionsKeptByTool.length > 0) details.append(list);
  }

  // 2. Read, but worth tidying.
  renderLayoutNotes(report.paper);

  // 3. Read fine, but position-dependent.
  {
    const details = renderFinding(
      3,
      `${report.suggestedForExclusion.length} question(s) worth keeping in their original option order`,
      'Their option text depends on its position, so shuffling can change the meaning ' +
        '("None of these", "Both (A) and (B)"). Nothing is wrong with ' +
        'them — this is your call, not the tool’s.',
    );

    const list = document.createElement('ul');
    for (const item of report.suggestedForExclusion) {
      list.append(element('li', '', `Q${item.questionNumber} (${item.subject}) — ${item.detail}`));
    }
    if (report.suggestedForExclusion.length > 0) details.append(list);

    if (report.optionsKeptByUser.length > 0) {
      details.append(
        element(
          'p',
          'hint',
          `Already kept because you asked: ${report.optionsKeptByUser.join(', ')}`,
        ),
      );
    }

    if (report.suggestedForExclusion.length > 0) {
      const addButton = document.createElement('button');
      addButton.type = 'button';
      addButton.className = 'button button--secondary';
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
      details.append(addButton);
    }
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
