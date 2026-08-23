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

export function setFileName(sourceFile: string, setNumber: number): string {
  const base = path.basename(sourceFile, path.extname(sourceFile));
  return `${base} - Set ${String(setNumber).padStart(2, '0')}.docx`;
}

export function setLabel(setNumber: number): string {
  return `SET ${String(setNumber).padStart(2, '0')}`;
}
