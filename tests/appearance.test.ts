/**
 * The three appearance modes.
 *
 * Two things here are easy to get wrong and invisible when they are: the page colour the
 * main process paints the window with must match the mode's own `--page`, or the window
 * flashes the wrong colour before the stylesheet arrives; and the number input's `max` in
 * the HTML must match `MAX_SETS`, or the box offers a number the app then refuses.
 */
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { PreferencesStore } from '../src/main/Preferences';
import { MAX_SETS, MIN_SETS } from '../src/core/generate/OutputFolder';
import { DEFAULT_THEME, THEMES, THEME_PAGE_COLOUR, asTheme } from '../src/shared/theme';

const STYLESHEET = path.join('src', 'renderer', 'styles.css');
const MARKUP = path.join('src', 'renderer', 'index.html');

describe('the appearance modes', () => {
  it('offers exactly the three the window offers', async () => {
    const html = await fs.readFile(MARKUP, 'utf8');
    const offered = [...html.matchAll(/<option value="([a-z]+)">/g)].map((match) => match[1]);

    expect(offered).toEqual([...THEMES]);
    expect(THEMES).toContain(DEFAULT_THEME);
  });

  it('paints the window in the same colour the stylesheet paints the page', async () => {
    // The main process has to know the colour before the window exists, so it cannot read
    // the stylesheet. This is what keeps the copy honest.
    const css = await fs.readFile(STYLESHEET, 'utf8');
    for (const theme of THEMES) {
      const selector =
        theme === DEFAULT_THEME
          ? new RegExp(`:root,\\s*:root\\[data-theme='light'\\][^}]*--page:\\s*(#[0-9a-f]{6})`, 'i')
          : new RegExp(`:root\\[data-theme='${theme}'\\][^}]*--page:\\s*(#[0-9a-f]{6})`, 'i');
      const match = selector.exec(css);
      expect(match, `no --page found for ${theme}`).toBeTruthy();
      expect(match![1]!.toLowerCase(), theme).toBe(THEME_PAGE_COLOUR[theme].toLowerCase());
    }
  });

  it('gives every mode a full palette, so none inherits another mode’s colours', async () => {
    const css = await fs.readFile(STYLESHEET, 'utf8');
    // Every name the light palette defines has to be defined by the other two as well;
    // a missing one would silently leave a light colour on a dark page.
    const block = (selector: RegExp): string => (selector.exec(css) ?? ['', ''])[0];
    const names = (text: string): string[] =>
      [...text.matchAll(/(--[a-z0-9-]+):/g)].map((match) => match[1]!).sort();

    const light = names(block(/:root,\s*:root\[data-theme='light'\][^}]*}/));
    expect(light.length).toBeGreaterThan(15);
    for (const theme of ['dim', 'dark'] as const) {
      const defined = names(block(new RegExp(`:root\\[data-theme='${theme}'\\][^}]*}`)));
      // --radius is a shape, not a colour, and is deliberately set once.
      expect(light.filter((name) => name !== '--radius' && !defined.includes(name)), theme).toEqual([]);
    }
  });

  it('does not follow the Windows setting any more, now that the user chooses', async () => {
    // Leaving the media query in would override a deliberate choice of Light on a machine
    // set to dark, which is the one thing a picker must never do.
    expect(await fs.readFile(STYLESHEET, 'utf8')).not.toContain('prefers-color-scheme');
  });

  it('accepts only the modes this build knows', () => {
    expect(asTheme('dim')).toBe('dim');
    expect(asTheme('DARK')).toBeUndefined();
    expect(asTheme('solarized')).toBeUndefined();
    expect(asTheme(undefined)).toBeUndefined();
  });
});

describe('the number of sets offered on screen', () => {
  it('matches what the app will actually accept', async () => {
    const html = await fs.readFile(MARKUP, 'utf8');
    const input = /<input id="set-count"[^>]*>/.exec(html)?.[0] ?? '';

    expect(input).toContain(`min="${MIN_SETS}"`);
    expect(input).toContain(`max="${MAX_SETS}"`);
  });
});

describe('remembering the chosen mode', () => {
  let folder = '';

  beforeEach(async () => {
    folder = await fs.mkdtemp(path.join(os.tmpdir(), 'shuffler-prefs-'));
  });

  afterEach(async () => {
    await fs.rm(folder, { recursive: true, force: true });
  });

  it('starts on the default when nothing has been saved', () => {
    expect(new PreferencesStore(folder).read()).toEqual({ theme: DEFAULT_THEME });
  });

  it('reads back what was written', () => {
    const store = new PreferencesStore(folder);
    expect(store.write({ theme: 'dark' })).toBe(true);
    expect(store.read()).toEqual({ theme: 'dark' });
    expect(new PreferencesStore(folder).read()).toEqual({ theme: 'dark' });
  });

  it('falls back to the default rather than failing on a damaged file', async () => {
    await fs.writeFile(path.join(folder, 'preferences.json'), '{ not json at all');
    expect(new PreferencesStore(folder).read()).toEqual({ theme: DEFAULT_THEME });

    await fs.writeFile(path.join(folder, 'preferences.json'), '{"theme":"solarized"}');
    expect(new PreferencesStore(folder).read()).toEqual({ theme: DEFAULT_THEME });
  });

  it('creates the folder it is given if it is not there yet', () => {
    const nested = path.join(folder, 'deep', 'deeper');
    expect(new PreferencesStore(nested).write({ theme: 'dim' })).toBe(true);
    expect(new PreferencesStore(nested).read()).toEqual({ theme: 'dim' });
  });
});

/**
 * The preload script is **sandboxed**: `require` works only for Electron's own modules, so
 * a relative import throws at load and takes the whole script with it. `window.shuffler` is
 * then never defined and every control in the window is dead - which is exactly what a
 * value import of `shared/theme` did. Nothing catches it at build time, and the app looks
 * fine until it is clicked, so it is checked here.
 */
describe('the preload script', () => {
  it('imports nothing but electron at runtime', async () => {
    const source = await fs.readFile(path.join('src', 'main', 'preload.ts'), 'utf8');
    const imports = [...source.matchAll(/^import\s+(?!type\b)(.+?)from\s+'([^']+)'/gm)];

    expect(imports.length).toBeGreaterThan(0);
    for (const [, clause, module] of imports) {
      // `import { type X }` is erased too, so only a clause with a value binding counts.
      const bindings = (clause ?? '').replace(/\btype\s+\w+\s*,?/g, '').replace(/[{},\s]/g, '');
      if (bindings === '') continue;
      expect(module, `preload requires "${module}" at runtime`).toBe('electron');
    }
  });

  it('keeps its inlined copies in step with the shared module', async () => {
    // The compiler checks this too (see the `typeof` assignments in preload.ts); this says
    // out loud what those two lines are for.
    const source = await fs.readFile(path.join('src', 'main', 'preload.ts'), 'utf8');
    const inlined = /const THEMES = \[([^\]]+)\] as const;/.exec(source)?.[1] ?? '';

    expect(inlined.replace(/['\s]/g, '').split(',')).toEqual([...THEMES]);
    expect(source).toContain(`const FALLBACK_THEME = '${DEFAULT_THEME}';`);
  });
});
