import { execFileSync } from 'node:child_process';
import { createHash, randomBytes } from 'node:crypto';
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { ADMIN_PASSWORD_DIGEST, ADMIN_SALT_BASE64, adminDigestOf } from '../src/main/AdminSecret';

/**
 * The installer checks the administrator's password with a SHA-256 written by hand in the
 * scripting language Windows Installer hosts - it has no crypto library to call. That
 * implementation is only correct if it agrees with the one the app uses, on every input,
 * which is what these tests establish: the real `installer/CheckPassword.js` is run under
 * Windows' own script host and its answers are compared with Node's.
 *
 * Without this the two could drift silently, and the failure would be the worst kind: an
 * installer that refuses the password everybody was given.
 */
const CHECK_PASSWORD = path.resolve('installer/CheckPassword.js');
const onWindows = process.platform === 'win32';

let workingDir = '';

beforeEach(async () => {
  workingDir = await fs.mkdtemp(path.join(os.tmpdir(), 'shuffler-msi-'));
});

afterEach(async () => {
  await fs.rm(workingDir, { recursive: true, force: true });
});

/**
 * Runs the installer's own script under cscript and returns one digest per input.
 *
 * The custom-action file is concatenated with a driver rather than imported, because the
 * Windows Script Host has no module system - which is also exactly how Windows Installer
 * loads it, as one flat script.
 */
async function digestsUnderWsh(saltHex: string, passwords: readonly string[]): Promise<string[]> {
  const source = await fs.readFile(CHECK_PASSWORD, 'utf8');
  // The inputs travel as a JSON file rather than on the command line, so that a password
  // with a quote or a space in it reaches the script unchanged.
  const inputFile = path.join(workingDir, 'input.json');
  await fs.writeFile(inputFile, JSON.stringify({ saltHex, passwords }), 'utf8');

  const driver = `
    var stream = new ActiveXObject('ADODB.Stream');
    stream.Type = 2; stream.Charset = 'utf-8'; stream.Open();
    stream.LoadFromFile(${JSON.stringify(inputFile)});
    var input = eval('(' + stream.ReadText() + ')');
    stream.Close();
    var out = [];
    for (var i = 0; i < input.passwords.length; i++) {
      out[out.length] = adminDigestOf(input.saltHex, input.passwords[i]);
    }
    WScript.Echo(out.join('\\n'));
  `;
  const script = path.join(workingDir, 'run.js');
  await fs.writeFile(script, `${source}\n${driver}`, 'utf8');

  const output = execFileSync('cscript', ['//Nologo', '//E:JScript', script], {
    encoding: 'utf8',
    windowsHide: true,
    timeout: 60_000,
  });
  return output.trim().split(/\r?\n/);
}

const saltHex = Buffer.from(ADMIN_SALT_BASE64, 'base64').toString('hex');

describe.skipIf(!onWindows)('the SHA-256 the installer carries', () => {
  it('agrees with Node on the administrator password itself', async () => {
    const [digest] = await digestsUnderWsh(saltHex, ['nextGenUI']);
    expect(digest).toBe(ADMIN_PASSWORD_DIGEST);
  });

  /**
   * The lengths either side of 55 and 119 bytes are where SHA-256's padding has to spill
   * into another block; an implementation that gets the padding wrong passes every short
   * input and fails exactly here.
   */
  it('agrees on inputs either side of every block boundary', async () => {
    const passwords = [
      '',
      'a',
      ...[38, 39, 40, 41, 42, 102, 103, 104, 105, 106].map((length) => 'x'.repeat(length)),
    ];
    const theirs = await digestsUnderWsh(saltHex, passwords);
    const ours = passwords.map((password) => adminDigestOf(password));
    expect(theirs).toEqual(ours);
  });

  it('agrees on passwords that are not plain ASCII', async () => {
    const passwords = ['pásswörd-ok', 'पासवर्ड-नया', 'pass word with spaces', 'quote"and\'apostrophe'];
    expect(await digestsUnderWsh(saltHex, passwords)).toEqual(passwords.map((p) => adminDigestOf(p)));
  });

  it('agrees on random inputs, salt and all', async () => {
    const passwords = Array.from({ length: 12 }, () => randomBytes(24).toString('base64'));
    expect(await digestsUnderWsh(saltHex, passwords)).toEqual(passwords.map((p) => adminDigestOf(p)));
  });

  /** A different salt must change every digest: the salt is really being hashed in. */
  it('follows the salt it is given', async () => {
    const otherSalt = randomBytes(16).toString('hex');
    const [withOther] = await digestsUnderWsh(otherSalt, ['nextGenUI']);
    expect(withOther).not.toBe(ADMIN_PASSWORD_DIGEST);
    const expected = createHash('sha256')
      .update(Buffer.concat([Buffer.from(otherSalt, 'hex'), Buffer.from('nextGenUI', 'utf8')]))
      .digest('hex');
    expect(withOther).toBe(expected);
  });
});
