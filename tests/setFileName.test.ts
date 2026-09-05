/**
 * How a generated set is named.
 *
 * The name carries the source paper, the run's date and time, and the set's letter, so that
 * sets from different runs of the same paper never look alike in a folder listing.
 */
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import {
  OutputFolderResolver,
  fileNameTimestamp,
  readableTimestamp,
  setDisplayName,
  setFileName,
  setLabel,
  setSuffix,
} from '../src/core/generate/OutputFolder';

/** 31 August 2026, 13:21 local time - the example from the request. */
const when = new Date(2026, 7, 31, 13, 21, 45);

describe('setFileName', () => {
  it('appends the run date, time and the set letter', () => {
    expect(setFileName('C:/papers/MTP-2 Group A.docx', 1, when)).toBe(
      'MTP-2 Group A - 31-08-2026-13-21-Set-A.docx',
    );
  });

  it('letters the sets A, B, C rather than numbering them', () => {
    expect(setFileName('paper.docx', 2, when)).toContain('-Set-B.docx');
    expect(setFileName('paper.docx', 7, when)).toContain('-Set-G.docx');
    expect(setFileName('paper.docx', 12, when)).toContain('-Set-L.docx');
  });

  it('keeps the source paper name, including its dots and spaces', () => {
    expect(setFileName('C:/q/Nano MTP 2 Physics_XI-2023__3418 _11.03.2023.docx', 3, when)).toBe(
      'Nano MTP 2 Physics_XI-2023__3418 _11.03.2023 - 31-08-2026-13-21-Set-C.docx',
    );
  });

  it('never contains a character Windows forbids in a file name', () => {
    // A ":" between the hours and minutes would read like a clock but could not be saved,
    // and Word cannot open a file that cannot be named.
    const name = setFileName('paper.docx', 1, when);
    expect(name).not.toMatch(/[:\\/*?"<>|]/);
  });
});

describe('fileNameTimestamp', () => {
  it('writes day-month-year-hour-minute, each padded', () => {
    expect(fileNameTimestamp(new Date(2026, 0, 5, 9, 4))).toBe('05-01-2026-09-04');
  });

  it('uses a 24-hour clock, so afternoon runs sort after morning ones', () => {
    const morning = fileNameTimestamp(new Date(2026, 7, 31, 9, 30));
    const afternoon = fileNameTimestamp(new Date(2026, 7, 31, 21, 30));
    expect(morning).toBe('31-08-2026-09-30');
    expect(afternoon).toBe('31-08-2026-21-30');
  });

  it('reads midnight as 00', () => {
    expect(fileNameTimestamp(new Date(2026, 11, 25, 0, 0))).toBe('25-12-2026-00-00');
  });
});

describe('readableTimestamp', () => {
  it('writes the same instant with a clock reading, for the report', () => {
    expect(readableTimestamp(when)).toBe('31-08-2026 at 13:21');
  });

  it('describes the same minute as the file names do', () => {
    expect(readableTimestamp(when).replace(' at ', '-').replace(':', '-')).toBe(
      fileNameTimestamp(when),
    );
  });
});

describe('the output folder', () => {
  let workingDir = '';
  let sourceFile = '';

  beforeEach(async () => {
    workingDir = await fs.mkdtemp(path.join(os.tmpdir(), 'shuffler-folder-'));
    sourceFile = path.join(workingDir, 'Sample Paper.docx');
    await fs.writeFile(sourceFile, 'not a real paper');
  });

  afterEach(async () => {
    await fs.rm(workingDir, { recursive: true, force: true });
  });

  it('carries the run date and time, like the files inside it', async () => {
    const folder = await new OutputFolderResolver().create(sourceFile, when);
    expect(path.basename(folder)).toBe('question-sets - 31-08-2026-13-21');
    expect(path.dirname(folder)).toBe(workingDir);
    // The stamp in the folder name and in a set's name are the same string.
    expect(setFileName(sourceFile, 1, when)).toContain(fileNameTimestamp(when));
  });

  it('marks a second run started in the same minute rather than overwriting the first', async () => {
    const resolver = new OutputFolderResolver();
    const first = await resolver.create(sourceFile, when);
    const second = await resolver.create(sourceFile, when);
    const third = await resolver.create(sourceFile, when);

    expect(path.basename(first)).toBe('question-sets - 31-08-2026-13-21');
    expect(path.basename(second)).toBe('question-sets - 31-08-2026-13-21-02');
    expect(path.basename(third)).toBe('question-sets - 31-08-2026-13-21-03');
  });

  it('never contains a character Windows forbids in a folder name', async () => {
    const folder = await new OutputFolderResolver().create(sourceFile, when);
    expect(path.basename(folder)).not.toMatch(/[:\\/*?"<>|]/);
  });
});

describe('setSuffix', () => {
  it('letters the sets A to Z', () => {
    expect(setSuffix(1)).toBe('A');
    expect(setSuffix(2)).toBe('B');
    expect(setSuffix(26)).toBe('Z');
  });

  it('carries on past Z the way a spreadsheet names its columns', () => {
    // 100 is the largest run the tool allows, so CV is as far as this ever goes.
    expect(setSuffix(27)).toBe('AA');
    expect(setSuffix(28)).toBe('AB');
    expect(setSuffix(52)).toBe('AZ');
    expect(setSuffix(53)).toBe('BA');
    expect(setSuffix(100)).toBe('CV');
  });

  it('gives every set of the largest allowed run a name of its own', () => {
    const names = Array.from({ length: 100 }, (_unused, index) => setSuffix(index + 1));
    expect(new Set(names).size).toBe(100);
    for (const name of names) expect(name).toMatch(/^[A-Z]+$/);
  });

  it('refuses a set number that is not a whole number from 1 upwards', () => {
    expect(() => setSuffix(0)).toThrow(/from 1 upwards/);
    expect(() => setSuffix(-1)).toThrow(/from 1 upwards/);
    expect(() => setSuffix(1.5)).toThrow(/from 1 upwards/);
  });
});

describe('setLabel and setDisplayName', () => {
  it('name the set the same way the file name does', () => {
    // One identifier in three places: the file name, the stamp inside the document and the
    // name on screen. A printed paper is traced back to its file by its letter alone.
    expect(setLabel(1)).toBe('SET A');
    expect(setLabel(11)).toBe('SET K');
    expect(setDisplayName(1)).toBe('Set A');
    expect(setDisplayName(11)).toBe('Set K');
    expect(setFileName('paper.docx', 11, when)).toContain(`-Set-${setSuffix(11)}.docx`);
  });
});
