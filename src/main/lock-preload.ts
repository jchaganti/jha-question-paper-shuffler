import { contextBridge, ipcRenderer } from 'electron';
import type { LockApi, LockStatus, SetOutcome, UnlockOutcome } from '../shared/lock';
import type { IPC } from './channels';

/**
 * The bridge for the sign-in window, and **only** the sign-in window.
 *
 * It is a separate preload from `preload.ts` on purpose. The lock screen runs before anyone
 * has proved who they are, so the page behind it must not be able to open a question paper,
 * write a file or start a generation run: a window loaded with this script has no way to
 * ask for any of those things, whatever its page is made to do.
 *
 * Sandboxed like the main preload, so relative `require` is impossible and the channel
 * names are inlined. The `satisfies`-style assignment below fails the build if they drift
 * from `src/main/channels.ts`.
 */
const CHANNELS = {
  readTheme: 'shuffler:read-theme',
  lockStatus: 'shuffler:lock-status',
  lockSet: 'shuffler:lock-set',
  lockUnlock: 'shuffler:lock-unlock',
  lockChange: 'shuffler:lock-change',
  lockReset: 'shuffler:lock-reset',
  lockProceed: 'shuffler:lock-proceed',
} as const;

type LockChannels = Pick<typeof IPC, keyof typeof CHANNELS>;
const channelsMatchMainProcess: LockChannels = CHANNELS;
void channelsMatchMainProcess;

const THEMES = ['light', 'dim', 'dark'] as const;
const FALLBACK_THEME = 'light';

/**
 * The appearance mode, applied before the page paints for the same reason as in the main
 * window: a sign-in box that flashes white before turning dark looks broken.
 */
const startupTheme: string =
  THEMES.find((theme) => theme === ipcRenderer.sendSync(CHANNELS.readTheme)) ?? FALLBACK_THEME;

const page = (globalThis as { document?: PreloadDocument }).document;
interface PreloadDocument {
  addEventListener(type: 'DOMContentLoaded', listener: () => void): void;
  readonly documentElement: { readonly dataset: Record<string, string> };
}

page?.addEventListener('DOMContentLoaded', () => {
  page.documentElement.dataset.theme = startupTheme;
});

const api: LockApi = {
  theme: startupTheme,
  status: () => ipcRenderer.invoke(CHANNELS.lockStatus) as Promise<LockStatus>,
  setPassword: (password) => ipcRenderer.invoke(CHANNELS.lockSet, password) as Promise<SetOutcome>,
  unlock: (password) => ipcRenderer.invoke(CHANNELS.lockUnlock, password) as Promise<UnlockOutcome>,
  changePassword: (current, next) =>
    ipcRenderer.invoke(CHANNELS.lockChange, current, next) as Promise<SetOutcome>,
  resetWithAdminPassword: (adminPassword, next) =>
    ipcRenderer.invoke(CHANNELS.lockReset, adminPassword, next) as Promise<SetOutcome>,
  proceed: () => ipcRenderer.send(CHANNELS.lockProceed),
};

contextBridge.exposeInMainWorld('lock', api);
