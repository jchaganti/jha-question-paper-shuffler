import type { LockApi, LockStatus, SetOutcome, UnlockOutcome } from '../shared/lock';

/**
 * The sign-in window.
 *
 * Like `renderer.ts` this file has no runtime imports - it is loaded as a module script and
 * only `import type` may cross into it - and it reaches the main process through the one
 * bridge its preload exposes. That bridge can set a password, check a password and open the
 * app; there is nothing else on it to go wrong with.
 */
declare global {
  interface Window {
    readonly lock: LockApi;
  }
}

const api = window.lock;

const el = <T extends HTMLElement>(id: string): T => {
  const node = document.getElementById(id);
  if (!node) throw new Error(`Missing element #${id}`);
  return node as T;
};

const subtitle = el<HTMLParagraphElement>('lock-subtitle');
const message = el<HTMLParagraphElement>('lock-message');

const panels = {
  set: el<HTMLFormElement>('panel-set'),
  enter: el<HTMLFormElement>('panel-enter'),
  change: el<HTMLFormElement>('panel-change'),
  reset: el<HTMLFormElement>('panel-reset'),
} as const;
type PanelName = keyof typeof panels;

const SUBTITLES: Record<PanelName, string> = {
  set: 'Set a password for this computer.',
  enter: 'Enter your password to open the app.',
  change: 'Change the password this computer asks for.',
  reset: 'Replace a forgotten password.',
};

const field = (id: string): HTMLInputElement => el<HTMLInputElement>(id);
const setPassword = field('set-password');
const setConfirm = field('set-confirm');
const enterPassword = field('enter-password');
const changeCurrent = field('change-current');
const changeNew = field('change-new');
const changeConfirm = field('change-confirm');
const resetAdmin = field('reset-admin');
const resetNew = field('reset-new');
const resetConfirm = field('reset-confirm');

const enterSubmit = el<HTMLButtonElement>('enter-submit');

/** The panel on screen. Everything else is `hidden`, so it is out of the tab order too. */
function show(panel: PanelName): void {
  for (const [name, form] of Object.entries(panels)) {
    form.hidden = name !== panel;
    if (name !== panel) form.reset();
  }
  subtitle.textContent = SUBTITLES[panel];
  say('');
  const first = panels[panel].querySelector('input');
  if (first instanceof HTMLInputElement) first.focus();
}

type Tone = 'bad' | 'calm' | 'good';

function say(text: string, tone: Tone = 'bad'): void {
  message.textContent = text;
  message.classList.toggle('lock__message--calm', tone === 'calm');
  message.classList.toggle('lock__message--good', tone === 'good');
}

/**
 * The lock-out, counted down in front of the person waiting.
 *
 * A button that is simply dead for half a minute reads as a broken app; one that says how
 * long is left reads as a rule. Only one countdown ever runs.
 */
let countdown: number | undefined;

function holdFor(waitMs: number): void {
  if (countdown !== undefined) clearInterval(countdown);
  let left = Math.ceil(waitMs / 1000);
  const tick = (): void => {
    if (left <= 0) {
      clearInterval(countdown);
      countdown = undefined;
      enterSubmit.disabled = false;
      say('');
      enterPassword.focus();
      return;
    }
    enterSubmit.disabled = true;
    say(`Too many wrong passwords. Try again in ${spell(left)}.`, 'calm');
    left--;
  };
  tick();
  countdown = window.setInterval(tick, 1000);
}

/** "45 seconds", "2 minutes 05 seconds" - read out rather than shown as a clock. */
function spell(seconds: number): string {
  if (seconds < 60) return `${seconds} second${seconds === 1 ? '' : 's'}`;
  const minutes = Math.floor(seconds / 60);
  const rest = String(seconds % 60).padStart(2, '0');
  return `${minutes} minute${minutes === 1 ? '' : 's'} ${rest} seconds`;
}

/** The two-boxes-must-match rule, in one place because three panels need it. */
function mismatch(first: HTMLInputElement, second: HTMLInputElement): boolean {
  if (first.value === second.value) return false;
  say('The two new passwords are not the same.');
  second.value = '';
  second.focus();
  return true;
}

function afterSet(outcome: SetOutcome, panel: PanelName): void {
  if (!outcome.ok) {
    say(outcome.message);
    return;
  }
  if (panel === 'change') {
    show('enter');
    say('Password changed. Enter the new one to open the app.', 'good');
    return;
  }
  api.proceed();
}

function afterUnlock(outcome: UnlockOutcome): void {
  if (outcome.ok) {
    api.proceed();
    return;
  }
  enterPassword.value = '';
  enterPassword.focus();
  switch (outcome.reason) {
    case 'wrong':
      say(
        outcome.attemptsLeft > 0
          ? `That is not the password. ${outcome.attemptsLeft} ${
              outcome.attemptsLeft === 1 ? 'try' : 'tries'
            } left before the app makes you wait.`
          : 'That is not the password.',
      );
      return;
    case 'waiting':
      holdFor(outcome.waitMs);
      return;
    case 'other-machine':
      say(
        'This password was set on a different computer or under a different Windows ' +
          'account. Use “Forgotten your password?” to set one for this computer.',
      );
      return;
    case 'unreadable':
      say(outcome.message);
      return;
  }
}

panels.set.addEventListener('submit', (event) => {
  event.preventDefault();
  if (mismatch(setPassword, setConfirm)) return;
  void api.setPassword(setPassword.value).then((outcome) => afterSet(outcome, 'set'));
});

panels.enter.addEventListener('submit', (event) => {
  event.preventDefault();
  void api.unlock(enterPassword.value).then(afterUnlock);
});

panels.change.addEventListener('submit', (event) => {
  event.preventDefault();
  if (mismatch(changeNew, changeConfirm)) return;
  void api
    .changePassword(changeCurrent.value, changeNew.value)
    .then((outcome) => afterSet(outcome, 'change'));
});

panels.reset.addEventListener('submit', (event) => {
  event.preventDefault();
  if (mismatch(resetNew, resetConfirm)) return;
  void api
    .resetWithAdminPassword(resetAdmin.value, resetNew.value)
    .then((outcome) => afterSet(outcome, 'reset'));
});

el<HTMLButtonElement>('go-change').addEventListener('click', () => show('change'));
el<HTMLButtonElement>('go-reset').addEventListener('click', () => show('reset'));
el<HTMLButtonElement>('change-cancel').addEventListener('click', () => show('enter'));
el<HTMLButtonElement>('reset-cancel').addEventListener('click', () => show('enter'));

function start(status: LockStatus): void {
  if (!status.usable) {
    for (const form of Object.values(panels)) form.hidden = true;
    subtitle.textContent = 'The app cannot be unlocked on this account.';
    say(status.problem);
    return;
  }
  show(status.mode);
  if (status.waitMs > 0) holdFor(status.waitMs);
}

void api.status().then(start);
