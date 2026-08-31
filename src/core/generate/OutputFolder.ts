import { promises as fs } from 'node:fs';
import path from 'node:path';

export const OUTPUT_FOLDER_PREFIX = 'question-sets-';

/**
 * Creates the next free `question-sets-NN` folder next to the source paper, so a new
 * run never overwrites the sets produced by an earlier run.
 */
export class OutputFolderResolver {
  async create(sourceFile: string): Promise<string> {
    const parent = path.dirname(path.resolve(sourceFile));
    for (let n = 1; n <= 999; n++) {
      const folder = path.join(parent, `${OUTPUT_FOLDER_PREFIX}${String(n).padStart(2, '0')}`);
      try {
        await fs.mkdir(folder, { recursive: false });
        return folder;
      } catch (error) {
        if ((error as NodeJS.ErrnoException).code === 'EEXIST') continue;
        throw error;
      }
    }
    throw new Error(`Could not create an output folder in ${parent}: 999 "${OUTPUT_FOLDER_PREFIX}NN" folders already exist.`);
  }
}

const pad = (value: number): string => String(value).padStart(2, '0');

/**
 * The run's date and time for use in a file name: `31-08-2026-13-21`.
 *
 * Local time, because that is the clock the user read when they pressed Generate.
 *
 * The minutes are separated with `-` rather than the `:` of a clock reading: Windows
 * forbids `:` in a file name (along with `\ / * ? " < > |`), and Word will not open a file
 * it cannot name.
 */
export function fileNameTimestamp(when: Date): string {
  return (
    `${pad(when.getDate())}-${pad(when.getMonth() + 1)}-${when.getFullYear()}` +
    `-${pad(when.getHours())}-${pad(when.getMinutes())}`
  );
}

/**
 * `<paper> - 31-08-2026-13-21-Set-01.docx`.
 *
 * Every set of one run carries the same timestamp - it is stamped once when the run
 * starts, not as each file is written, so a run that crosses a minute boundary still
 * produces one consistently named batch.
 */
export function setFileName(sourceFile: string, setNumber: number, when: Date): string {
  const base = path.basename(sourceFile, path.extname(sourceFile));
  return `${base} - ${fileNameTimestamp(when)}-Set-${pad(setNumber)}.docx`;
}

export function setLabel(setNumber: number): string {
  return `SET ${pad(setNumber)}`;
}
