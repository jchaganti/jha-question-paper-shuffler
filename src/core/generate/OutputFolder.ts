import { promises as fs } from 'node:fs';
import path from 'node:path';

export const OUTPUT_FOLDER_PREFIX = 'question-sets - ';

/**
 * Creates the folder for one run next to the source paper: `question-sets - 31-08-2026-13-21`.
 *
 * The date and time are the run's own, the same stamp every file inside the folder carries,
 * so a folder can be matched to its sets at a glance and runs sort in the order they were
 * made. Two runs started in the same minute would collide, so the second gets `-02`, the
 * third `-03`, and so on - a new run never overwrites an earlier one's sets.
 */
export class OutputFolderResolver {
  async create(sourceFile: string, when: Date): Promise<string> {
    const parent = path.dirname(path.resolve(sourceFile));
    const base = path.join(parent, `${OUTPUT_FOLDER_PREFIX}${fileNameTimestamp(when)}`);
    for (let n = 1; n <= 99; n++) {
      const folder = n === 1 ? base : `${base}-${pad(n)}`;
      try {
        await fs.mkdir(folder, { recursive: false });
        return folder;
      } catch (error) {
        if ((error as NodeJS.ErrnoException).code === 'EEXIST') continue;
        throw error;
      }
    }
    throw new Error(
      `Could not create an output folder in ${parent}: 99 runs are already recorded for ${fileNameTimestamp(when)}.`,
    );
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
 * The same instant written for a person to read: `31-08-2026 at 13:21`.
 *
 * The report is the one place with room for a real clock reading, so it uses the `:` a file
 * name may not contain - and the same local time, so the report and the folder it sits in
 * plainly describe the same run.
 */
export function readableTimestamp(when: Date): string {
  return (
    `${pad(when.getDate())}-${pad(when.getMonth() + 1)}-${when.getFullYear()}` +
    ` at ${pad(when.getHours())}:${pad(when.getMinutes())}`
  );
}

/**
 * The name of a set: `A`, `B`, `C` ... `Z`, then `AA`, `AB` ... for a run past 26.
 *
 * Sets are lettered rather than numbered because that is how a hall names them - "you have
 * Set B" - and because a set number invites confusion with a question number, which this
 * tool talks about constantly. Bijective base 26, so 26 is `Z` and 27 is `AA`; the largest
 * run the tool allows is 100 sets, which reaches `CV`.
 */
export function setSuffix(setNumber: number): string {
  if (!Number.isInteger(setNumber) || setNumber < 1) {
    throw new Error(`A set number must be a whole number from 1 upwards, not ${setNumber}.`);
  }
  let name = '';
  for (let n = setNumber; n > 0; n = Math.floor((n - 1) / 26)) {
    name = String.fromCharCode('A'.charCodeAt(0) + ((n - 1) % 26)) + name;
  }
  return name;
}

/**
 * `<paper> - 31-08-2026-13-21-Set-A.docx`.
 *
 * Every set of one run carries the same timestamp - it is stamped once when the run
 * starts, not as each file is written, so a run that crosses a minute boundary still
 * produces one consistently named batch.
 */
export function setFileName(sourceFile: string, setNumber: number, when: Date): string {
  const base = path.basename(sourceFile, path.extname(sourceFile));
  return `${base} - ${fileNameTimestamp(when)}-Set-${setSuffix(setNumber)}.docx`;
}

/**
 * What is stamped into the document itself, beside the answer-key heading: `SET A`.
 *
 * The same letter as the file name, so a printed paper can be matched to the file it came
 * from without opening anything.
 */
export function setLabel(setNumber: number): string {
  return `SET ${setSuffix(setNumber)}`;
}

/** The same set named for the UI and the report: `Set A`. */
export function setDisplayName(setNumber: number): string {
  return `Set ${setSuffix(setNumber)}`;
}
