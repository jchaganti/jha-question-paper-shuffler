import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { DEFAULT_THEME, asTheme, type Theme } from '../shared/theme';

/**
 * The handful of choices that outlive one run of the app, kept in a small JSON file in the
 * user's own application-data folder.
 *
 * Read **synchronously**, because the window's background colour has to be known before
 * the window is created - see `THEME_PAGE_COLOUR`. The file is tiny and read once at
 * startup, so there is nothing to gain by making it async and a flash to lose.
 *
 * Nothing here is essential: a missing, unreadable or corrupt file simply means the
 * defaults. Losing a colour preference is not worth refusing to start over.
 */
export interface Preferences {
  readonly theme: Theme;
}

const FILE_NAME = 'preferences.json';

export class PreferencesStore {
  private readonly file: string;

  /** @param folder Electron's `app.getPath('userData')`. */
  constructor(folder: string) {
    this.file = path.join(folder, FILE_NAME);
  }

  read(): Preferences {
    try {
      if (!existsSync(this.file)) return { theme: DEFAULT_THEME };
      const stored: unknown = JSON.parse(readFileSync(this.file, 'utf8'));
      const theme =
        typeof stored === 'object' && stored !== null
          ? asTheme((stored as { theme?: unknown }).theme)
          : undefined;
      return { theme: theme ?? DEFAULT_THEME };
    } catch {
      // An unreadable preferences file is not a reason to fail to start.
      return { theme: DEFAULT_THEME };
    }
  }

  /** Returns whether it was actually saved, so a caller can say so rather than assume. */
  write(preferences: Preferences): boolean {
    try {
      mkdirSync(path.dirname(this.file), { recursive: true });
      writeFileSync(this.file, `${JSON.stringify(preferences, null, 2)}\n`, 'utf8');
      return true;
    } catch {
      return false;
    }
  }
}
