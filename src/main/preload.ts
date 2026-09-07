import { contextBridge, ipcRenderer } from 'electron';
import type { GenerationRequest, ProgressEvent } from '../shared/types';
import type { DEFAULT_THEME as DefaultTheme, THEMES as Themes, Theme } from '../shared/theme';
import type { IPC } from './channels';

/**
 * A sandboxed preload script cannot `require` relative modules, so the channel names
 * are inlined here. The `satisfies`-style assignment below makes the compiler fail if
 * they ever drift from `src/main/channels.ts`.
 */
const CHANNELS = {
  pickSourceFile: 'shuffler:pick-source-file',
  inspect: 'shuffler:inspect',
  dryRun: 'shuffler:dry-run',
  generate: 'shuffler:generate',
  revealFolder: 'shuffler:reveal-folder',
  progress: 'shuffler:progress',
  readTheme: 'shuffler:read-theme',
  writeTheme: 'shuffler:write-theme',
} as const;

const channelsMatchMainProcess: typeof IPC = CHANNELS;
void channelsMatchMainProcess;

/**
 * The appearance modes, inlined for the same reason as the channels above: this script is
 * **sandboxed**, so `require('../shared/theme')` throws at load and takes the whole preload
 * with it - `window.shuffler` is then never defined and nothing in the window works.
 *
 * Only `import type` may cross this boundary. The two assignments below are the guard: the
 * compiler fails if either drifts from `src/shared/theme.ts`.
 */
const THEMES = ['light', 'dim', 'dark'] as const;
const FALLBACK_THEME = 'light';

const themesMatchShared: typeof Themes = THEMES;
const fallbackMatchesShared: typeof DefaultTheme = FALLBACK_THEME;
void themesMatchShared;
void fallbackMatchesShared;

const asTheme = (value: unknown): Theme | undefined => THEMES.find((theme) => theme === value);

/**
 * The appearance mode, read before the page exists and written onto <html> the moment the
 * document is parsed - which is before the browser paints.
 *
 * Doing it here rather than in `renderer.js` is what stops the flash: the renderer is a
 * module script, so it does not run until after the first paint, and the window would show
 * one frame of the default theme first. The renderer still owns the drop-down; it only
 * needs this attribute to be right from the start.
 */
const startupTheme: Theme = asTheme(ipcRenderer.sendSync(CHANNELS.readTheme)) ?? FALLBACK_THEME;

/**
 * The page this preload script is attached to.
 *
 * Reached through `globalThis` rather than as a global, because this is the only file
 * compiled with the main process that runs inside a page - `tsconfig.main.json` leaves the
 * DOM library out on purpose, so that `main.ts` cannot reach for a browser API that will
 * not be there. Optional throughout: a preload with no document is not worth crashing over,
 * and the renderer sets the attribute again anyway when the user changes the mode.
 */
const page = (globalThis as { document?: PreloadDocument }).document;
interface PreloadDocument {
  addEventListener(type: 'DOMContentLoaded', listener: () => void): void;
  readonly documentElement: { readonly dataset: Record<string, string> };
}

page?.addEventListener('DOMContentLoaded', () => {
  page.documentElement.dataset.theme = startupTheme;
});

/**
 * The only bridge between the UI and Node. The renderer has no direct filesystem or
 * module access; it can call exactly these operations.
 */
contextBridge.exposeInMainWorld('shuffler', {
  theme: startupTheme,
  setTheme: (theme: Theme) => ipcRenderer.invoke(CHANNELS.writeTheme, theme),
  pickSourceFile: () => ipcRenderer.invoke(CHANNELS.pickSourceFile),
  inspect: (sourceFile: string) => ipcRenderer.invoke(CHANNELS.inspect, sourceFile),
  dryRun: (request: GenerationRequest) => ipcRenderer.invoke(CHANNELS.dryRun, request),
  generate: (request: GenerationRequest) => ipcRenderer.invoke(CHANNELS.generate, request),
  revealFolder: (folder: string) => ipcRenderer.invoke(CHANNELS.revealFolder, folder),
  onProgress: (listener: (event: ProgressEvent) => void) => {
    ipcRenderer.on(CHANNELS.progress, (_event, payload: ProgressEvent) => listener(payload));
  },
});
