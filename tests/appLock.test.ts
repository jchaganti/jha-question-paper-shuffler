import { createHash } from 'node:crypto';
import { promises as fs, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { AppLock, passwordProblem, waitAfter, type Vault } from '../src/main/AppLock';
import { ADMIN_PASSWORD_DIGEST, adminDigestOf, isAdminPassword } from '../src/main/AdminSecret';
import { FREE_ATTEMPTS, LONGEST_WAIT_MS, MIN_PASSWORD_LENGTH } from '../src/shared/lock';

/**
 * A stand-in for Windows' own encryption.
 *
 * It behaves the way DPAPI behaves in the one respect the app depends on: what it seals for
 * one Windows account cannot be opened by another. The `account` it is created with is that
 * account, and `decrypt` refuses anything sealed under a different one - which is how the
 * "the file was copied here" tests get to happen without a second computer.
 */
class FakeVault implements Vault {
  constructor(
    private readonly account = 'this-account',
    private readonly on = true,
  ) {}

  available(): boolean {
    return this.on;
  }

  encrypt(plain: string): Buffer {
    return Buffer.from(`${this.account}\u0000${plain}`, 'utf8');
  }

  decrypt(sealed: Buffer): string {
    const [account, ...rest] = sealed.toString('utf8').split('\u0000');
    if (account !== this.account) throw new Error('sealed for another account');
    return rest.join('\u0000');
  }
}

const PASSWORD = 'staffroom-2026';
const MACHINE = 'machine-guid-aaaa';

let folder = '';
/** Test clock, so a fifteen-minute lock-out takes no time to check. */
let clock = 0;

const lockFor = (options: { machine?: string; vault?: Vault } = {}): AppLock =>
  new AppLock(
    folder,
    options.vault ?? new FakeVault(),
    () => options.machine ?? MACHINE,
    () => clock,
  );

/** Wrong passwords, one after another, returning the last answer. */
const missTimes = (lock: AppLock, times: number): ReturnType<AppLock['unlock']> => {
  let last = lock.unlock('not-the-password');
  for (let i = 1; i < times; i++) last = lock.unlock('not-the-password');
  return last;
};

beforeEach(async () => {
  folder = await fs.mkdtemp(path.join(os.tmpdir(), 'shuffler-lock-'));
  clock = Date.parse('2026-09-28T09:00:00Z');
});

afterEach(async () => {
  await fs.rm(folder, { recursive: true, force: true });
});

describe('the password the app asks for when it starts', () => {
  it('asks for a password to be set when there is none', () => {
    const lock = lockFor();
    expect(lock.isConfigured()).toBe(false);
    expect(lock.status()).toMatchObject({ mode: 'set', usable: true, waitMs: 0 });
  });

  it('opens with the password that was set, and not with any other', () => {
    const lock = lockFor();
    expect(lock.setPassword(PASSWORD)).toEqual({ ok: true });
    expect(lock.status().mode).toBe('enter');
    expect(lock.unlock(PASSWORD)).toEqual({ ok: true });
    expect(lock.unlock('staffroom-2025')).toMatchObject({ ok: false, reason: 'wrong' });
  });

  it('remembers the password after the app is closed and opened again', () => {
    lockFor().setPassword(PASSWORD);
    // A second instance over the same folder is what the next run of the app is.
    expect(lockFor().unlock(PASSWORD)).toEqual({ ok: true });
  });

  it('refuses a password shorter than the minimum, and says how long it must be', () => {
    const lock = lockFor();
    const outcome = lock.setPassword('short');
    expect(outcome).toMatchObject({ ok: false });
    expect(outcome.ok ? '' : outcome.message).toContain(String(MIN_PASSWORD_LENGTH));
    expect(lock.isConfigured()).toBe(false);
  });

  it('will not let a second password be set over the first', () => {
    const lock = lockFor();
    lock.setPassword(PASSWORD);
    expect(lock.setPassword('something-else-entirely')).toMatchObject({ ok: false });
    expect(lock.unlock(PASSWORD)).toEqual({ ok: true });
  });
});

describe('guessing', () => {
  it('counts down the tries left before the wait starts', () => {
    const lock = lockFor();
    lock.setPassword(PASSWORD);
    expect(lock.unlock('wrong-one-here')).toEqual({ ok: false, reason: 'wrong', attemptsLeft: FREE_ATTEMPTS - 1 });
    expect(lock.unlock('wrong-two-here')).toEqual({ ok: false, reason: 'wrong', attemptsLeft: FREE_ATTEMPTS - 2 });
  });

  it('makes the next guess wait once the free tries are used up', () => {
    const lock = lockFor();
    lock.setPassword(PASSWORD);
    const outcome = missTimes(lock, FREE_ATTEMPTS + 1);
    expect(outcome).toMatchObject({ ok: false, reason: 'waiting' });
    expect(outcome.ok === false && outcome.reason === 'waiting' ? outcome.waitMs : 0).toBeGreaterThan(0);
  });

  it('refuses even the right password while the wait is running', () => {
    const lock = lockFor();
    lock.setPassword(PASSWORD);
    missTimes(lock, FREE_ATTEMPTS + 1);
    expect(lock.unlock(PASSWORD)).toMatchObject({ ok: false, reason: 'waiting' });
  });

  it('lets the right password through once the wait is over', () => {
    const lock = lockFor();
    lock.setPassword(PASSWORD);
    missTimes(lock, FREE_ATTEMPTS + 1);
    clock += LONGEST_WAIT_MS + 1;
    expect(lock.unlock(PASSWORD)).toEqual({ ok: true });
  });

  it('keeps the count across restarts, so closing the app is not a way round the wait', () => {
    lockFor().setPassword(PASSWORD);
    missTimes(lockFor(), FREE_ATTEMPTS + 1);
    expect(lockFor().status().waitMs).toBeGreaterThan(0);
    expect(lockFor().unlock(PASSWORD)).toMatchObject({ ok: false, reason: 'waiting' });
  });

  it('forgets the count after a correct password, so a day of typos does not add up', () => {
    const lock = lockFor();
    lock.setPassword(PASSWORD);
    missTimes(lock, FREE_ATTEMPTS - 1);
    expect(lock.unlock(PASSWORD)).toEqual({ ok: true });
    expect(lock.unlock('wrong-again-now')).toEqual({ ok: false, reason: 'wrong', attemptsLeft: FREE_ATTEMPTS - 1 });
  });

  it('doubles the wait, and stops doubling at a quarter of an hour', () => {
    expect(waitAfter(FREE_ATTEMPTS + 1)).toBe(30_000);
    expect(waitAfter(FREE_ATTEMPTS + 2)).toBe(60_000);
    expect(waitAfter(FREE_ATTEMPTS + 3)).toBe(120_000);
    expect(waitAfter(FREE_ATTEMPTS + 50)).toBe(LONGEST_WAIT_MS);
  });
});

describe('tied to this computer', () => {
  it('says so when the password file came from another machine', () => {
    lockFor().setPassword(PASSWORD);
    // Same file, same Windows account, different computer: only the machine identifier moves.
    const elsewhere = lockFor({ machine: 'machine-guid-bbbb' });
    expect(elsewhere.unlock(PASSWORD)).toEqual({ ok: false, reason: 'other-machine' });
  });

  it('is unreadable under another Windows account, so that account sets its own password', () => {
    lockFor().setPassword(PASSWORD);
    const otherAccount = lockFor({ vault: new FakeVault('someone-else') });
    expect(otherAccount.isConfigured()).toBe(false);
    expect(otherAccount.status().mode).toBe('set');
  });

  it('refuses to store a password at all when Windows will not encrypt the file', () => {
    const lock = lockFor({ vault: new FakeVault('this-account', false) });
    expect(lock.status()).toMatchObject({ usable: false });
    expect(lock.status().problem).not.toBe('');
    expect(lock.setPassword(PASSWORD)).toMatchObject({ ok: false });
  });

  it('never writes the password, or the machine it is tied to, to disk', () => {
    const lock = lockFor();
    lock.setPassword(PASSWORD);
    const onDisk = readFileSync(path.join(folder, 'app-lock.dat')).toString('utf8');
    expect(onDisk).not.toContain(PASSWORD);
    expect(onDisk).not.toContain(MACHINE);
  });

  /**
   * The machine is not merely *recorded* in the file and checked against - it is part of
   * what the password is hashed with. So someone who rewrites the file to claim it belongs
   * to their own machine, which is exactly what they would do, gets no further: the stored
   * bytes were never reachable from this password anywhere else.
   */
  it('still does not open when the machine stamp in the file has been forged', () => {
    lockFor().setPassword(PASSWORD);

    const thief = 'machine-guid-bbbb';
    const vault = new FakeVault();
    const record = JSON.parse(vault.decrypt(readFileSync(path.join(folder, 'app-lock.dat'))));
    record.machine = createHash('sha256').update(thief, 'utf8').digest('base64');

    const stolen = path.join(folder, 'stolen');
    mkdirSync(stolen, { recursive: true });
    writeFileSync(path.join(stolen, 'app-lock.dat'), vault.encrypt(JSON.stringify(record)));

    const forged = new AppLock(stolen, vault, () => thief, () => clock);
    expect(forged.unlock(PASSWORD)).toMatchObject({ ok: false, reason: 'wrong' });
  });
});

describe('changing the password', () => {
  it('needs the current one', () => {
    const lock = lockFor();
    lock.setPassword(PASSWORD);
    expect(lock.changePassword('not-the-old-one', 'brand-new-password')).toMatchObject({ ok: false });
    expect(lock.unlock(PASSWORD)).toEqual({ ok: true });
  });

  it('replaces it, and the old one stops working', () => {
    const lock = lockFor();
    lock.setPassword(PASSWORD);
    expect(lock.changePassword(PASSWORD, 'brand-new-password')).toEqual({ ok: true });
    expect(lock.unlock('brand-new-password')).toEqual({ ok: true });
    expect(lock.unlock(PASSWORD)).toMatchObject({ ok: false, reason: 'wrong' });
  });

  it('holds the new one to the same length rule', () => {
    const lock = lockFor();
    lock.setPassword(PASSWORD);
    expect(lock.changePassword(PASSWORD, 'tiny')).toMatchObject({ ok: false });
    expect(lock.unlock(PASSWORD)).toEqual({ ok: true });
  });

  it('counts a wrong current password towards the lock-out', () => {
    const lock = lockFor();
    lock.setPassword(PASSWORD);
    for (let i = 0; i <= FREE_ATTEMPTS; i++) lock.changePassword('wrong-one-here', 'brand-new-password');
    expect(lock.unlock(PASSWORD)).toMatchObject({ ok: false, reason: 'waiting' });
  });
});

describe('a forgotten password', () => {
  it('is replaced by the administrator password used to install the app', () => {
    const lock = lockFor();
    lock.setPassword(PASSWORD);
    expect(lock.resetWithAdminPassword('nextGenUI', 'a-fresh-password')).toEqual({ ok: true });
    expect(lock.unlock('a-fresh-password')).toEqual({ ok: true });
    expect(lock.unlock(PASSWORD)).toMatchObject({ ok: false, reason: 'wrong' });
  });

  it('is not replaced by anything else', () => {
    const lock = lockFor();
    lock.setPassword(PASSWORD);
    expect(lock.resetWithAdminPassword('nextgenui', 'a-fresh-password')).toMatchObject({ ok: false });
    expect(lock.resetWithAdminPassword('', 'a-fresh-password')).toMatchObject({ ok: false });
    expect(lock.unlock(PASSWORD)).toEqual({ ok: true });
  });

  it('clears the lock-out too, so a reset is not followed by a wait', () => {
    const lock = lockFor();
    lock.setPassword(PASSWORD);
    missTimes(lock, FREE_ATTEMPTS + 1);
    expect(lock.resetWithAdminPassword('nextGenUI', 'a-fresh-password')).toEqual({ ok: true });
    expect(lock.unlock('a-fresh-password')).toEqual({ ok: true });
  });
});

describe('the administrator password', () => {
  it('is the one the installer asks for', () => {
    expect(isAdminPassword('nextGenUI')).toBe(true);
  });

  it('is not anything near it', () => {
    for (const near of ['nextgenui', 'NextGenUI', 'nextGenUi', ' nextGenUI', 'nextGenUI ', '']) {
      expect(isAdminPassword(near), near).toBe(false);
    }
  });

  /**
   * The digest that is compiled into the MSI. If this test has to be updated, the installer
   * has to be rebuilt - which is the point of pinning it.
   */
  it('is stored as a digest, not as the word itself', () => {
    expect(ADMIN_PASSWORD_DIGEST).toMatch(/^[0-9a-f]{64}$/);
    expect(adminDigestOf('nextGenUI')).toBe(ADMIN_PASSWORD_DIGEST);
    expect(adminDigestOf('nextGenUI')).not.toContain('nextGenUI');
  });
});

describe('what counts as a usable password', () => {
  it('accepts a long one and explains a short one', () => {
    expect(passwordProblem('a-long-enough-one')).toBe('');
    expect(passwordProblem('short')).toContain(String(MIN_PASSWORD_LENGTH));
    expect(passwordProblem('        ')).toContain('spaces');
  });
});
