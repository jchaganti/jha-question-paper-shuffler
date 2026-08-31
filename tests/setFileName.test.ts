/**
 * How a generated set is named.
 *
 * The name carries the source paper, the run's date and time, and the set number, so that
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
  setFileName,
  setLabel,
} from '../src/core/generate/OutputFolder';

/** 31 August 2026, 13:21 local time - the example from the request. */
const when = new Date(2026, 7, 31, 13, 21, 45);

describe('setFileName', () => {
  it('appends the run date, time and set number', () => {
    expect(setFileName('C:/papers/MTP-2 Group A.docx', 1, when)).toBe(
      'MTP-2 Group A - 31-08-2026-13-21-Set-01.docx',
    );
  });

  it('pads the set number to two digits', () => {
    expect(setFileName('paper.docx', 7, when)).toContain('-Set-07.docx');
    expect(setFileName('paper.docx', 12, when)).toContain('-Set-12.docx');
  });

  it('keeps the source paper name, including its dots and spaces', () => {
    expect(setFileName('C:/q/Nano MTP 2 Physics_XI-2023__3418 _11.03.2023.docx', 3, when)).toBe(
      'Nano MTP 2 Physics_XI-2023__3418 _11.03.2023 - 31-08-2026-13-21-Set-03.docx',
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

describe('setLabel', () => {
  it('is unchanged - the label stamped inside the document stays "SET NN"', () => {
    // The timestamp belongs in the file name, not printed on the answer key page.
    expect(setLabel(1)).toBe('SET 01');
    expect(setLabel(11)).toBe('SET 11');
  });
});
