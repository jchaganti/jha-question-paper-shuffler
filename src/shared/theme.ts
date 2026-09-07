/**
 * The three appearance modes, and the one place that decides what counts as a valid one.
 *
 * Free of runtime dependencies like the rest of `shared/`, so the main process, the preload
 * bridge and the renderer can all agree on the same three names.
 */
export const THEMES = ['light', 'dim', 'dark'] as const;
export type Theme = (typeof THEMES)[number];

/** What a fresh install opens on. */
export const DEFAULT_THEME: Theme = 'light';

/**
 * The page colour of each mode, matching `--page` in `styles.css`.
 *
 * The main process needs these before the window exists: a window painted white and then
 * repainted dark once the stylesheet arrives flashes, which is what the old colour picker
 * did (see the note in `handoff.md`). Giving `BrowserWindow` the right colour up front
 * means the first thing drawn is already the right one.
 *
 * These are the only colour values outside the stylesheet, so a test checks they still
 * match it.
 */
export const THEME_PAGE_COLOUR: Record<Theme, string> = {
  light: '#f3f3f3',
  dim: '#2e2e2e',
  dark: '#202020',
};

/** Narrows whatever was stored (or sent over IPC) to a mode this build knows. */
export function asTheme(value: unknown): Theme | undefined {
  return THEMES.find((theme) => theme === value);
}
