import { existsSync, mkdirSync, readFileSync, renameSync, unlinkSync, writeFileSync } from 'node:fs';
import { createHash, randomBytes, scryptSync, timingSafeEqual } from 'node:crypto';
import path from 'node:path';
import {
  FIRST_WAIT_MS,
  FREE_ATTEMPTS,
  LONGEST_WAIT_MS,
  MIN_PASSWORD_LENGTH,
  type LockStatus,
  type SetOutcome,
  type UnlockOutcome,
} from '../shared/lock';
import { isAdminPassword } from './AdminSecret';

/**
 * Windows' own encryption, as Electron exposes it.
 *
 * An interface rather than a direct call so that the tests can run this class without an
 * Electron process - the same stand-in the packaging code uses for `IDocxPackage`.
 */
export interface Vault {
  /** Whether the operating system will encrypt for us at all. */
  available(): boolean;
  encrypt(plain: string): Buffer;
  decrypt(sealed: Buffer): string;
}

/**
 * Something that identifies this computer and this Windows account.
 *
 * Used twice: mixed into the key the password is turned into, so a password file lifted
 * from another machine cannot be made to match, and stored alongside it, so that when one
 * *is* lifted the app can say "this was set on a different computer" instead of "wrong
 * password".
 */
export type MachineId = () => string;

const FILE_NAME = 'app-lock.dat';
const FORMAT = 1;

/** scrypt at Node's own "interactive" end: roughly a tenth of a second per attempt. */
const KEY_LENGTH = 32;
const SCRYPT = { N: 1 << 15, r: 8, p: 1, maxmem: 96 * 1024 * 1024 } as const;

interface LockRecord {
  readonly version: number;
  /** base64, 16 bytes. New on every password change. */
  readonly salt: string;
  /** base64, 32 bytes: scrypt over the password and this machine. */
  readonly verifier: string;
  /** base64 SHA-256 of the machine identifier, for the "different computer" message. */
  readonly machine: string;
  readonly setAt: string;
  /** Consecutive wrong passwords. Reset by a correct one. */
  failures: number;
  /** Epoch milliseconds before which no password may be tried. */
  lockedUntil: number;
}

/**
 * The password the app asks for when it starts.
 *
 * Chosen by whoever uses the app, not by whoever installed it, and it never leaves this
 * computer: what is stored is scrypt run over the password and a machine identifier, and
 * even that is sealed with Windows' own per-account encryption before it is written. The
 * file therefore cannot be copied to another computer, carried to another Windows account,
 * or read for the password it stands for.
 *
 * Every answer this class gives is a value, never an exception: being locked out of your
 * own app is a situation to explain, not to crash on.
 */
export class AppLock {
  private readonly file: string;
  private cached: LockRecord | undefined;

  /**
   * @param folder Electron's `app.getPath('userData')`.
   * @param vault Windows' encryption, or a stand-in under test.
   * @param machineId What identifies this computer.
   * @param now Injected so the lock-out tests do not have to sleep.
   */
  constructor(
    folder: string,
    private readonly vault: Vault,
    private readonly machineId: MachineId,
    private readonly now: () => number = Date.now,
  ) {
    this.file = path.join(folder, FILE_NAME);
  }

  /** Whether a password has been chosen on this computer yet. */
  isConfigured(): boolean {
    return this.read() !== undefined;
  }

  status(): LockStatus {
    if (!this.vault.available()) {
      return {
        mode: 'enter',
        waitMs: 0,
        usable: false,
        problem:
          'Windows will not encrypt the password file for this account, so the app cannot ' +
          'store a password safely. Sign in to Windows with the account the app was installed ' +
          'for, or ask your administrator to reinstall the app.',
      };
    }
    const record = this.read();
    return {
      mode: record ? 'enter' : 'set',
      waitMs: record ? this.waitLeft(record) : 0,
      usable: true,
      problem: '',
    };
  }

  /**
   * Chooses the password for the first time.
   *
   * Refuses if one has already been set: changing it goes through `changePassword`, which
   * asks for the old one. Otherwise anyone who reached a running app could replace it.
   */
  setPassword(password: string): SetOutcome {
    if (this.isConfigured()) {
      return { ok: false, message: 'A password has already been set on this computer.' };
    }
    return this.store(password);
  }

  unlock(password: string): UnlockOutcome {
    const record = this.read();
    if (!record) return { ok: false, reason: 'unreadable', message: 'No password has been set on this computer yet.' };
    if (record.machine !== this.machineFingerprint()) return { ok: false, reason: 'other-machine' };

    const waitMs = this.waitLeft(record);
    if (waitMs > 0) return { ok: false, reason: 'waiting', waitMs };

    if (this.matches(password, record)) {
      // A correct password clears the count, so an afternoon of typos does not add up to a
      // lock-out weeks later.
      this.write({ ...record, failures: 0, lockedUntil: 0 });
      return { ok: true };
    }

    const failures = record.failures + 1;
    const lockedUntil = failures > FREE_ATTEMPTS ? this.now() + waitAfter(failures) : 0;
    this.write({ ...record, failures, lockedUntil });
    if (lockedUntil > 0) return { ok: false, reason: 'waiting', waitMs: lockedUntil - this.now() };
    return { ok: false, reason: 'wrong', attemptsLeft: FREE_ATTEMPTS - failures };
  }

  /** Replaces the password, having been shown the current one. */
  changePassword(current: string, next: string): SetOutcome {
    const record = this.read();
    if (!record) return { ok: false, message: 'No password has been set on this computer yet.' };
    if (this.waitLeft(record) > 0) {
      return { ok: false, message: 'Too many wrong passwords. Wait for the countdown to finish, then try again.' };
    }
    if (!this.matches(current, record)) {
      const failures = record.failures + 1;
      this.write({
        ...record,
        failures,
        lockedUntil: failures > FREE_ATTEMPTS ? this.now() + waitAfter(failures) : 0,
      });
      return { ok: false, message: 'That is not the current password.' };
    }
    return this.store(next);
  }

  /**
   * The way back in for someone who has forgotten their password: the administrator types
   * the password they used to install the app, and a new one is chosen there and then.
   *
   * Deliberately the *only* reset. Deleting the file by hand would do the same thing, and
   * nothing can stop that on a computer someone already controls - but a reset that is in
   * the app, and asks for the administrator's password, is the one a school can actually
   * use, and it leaves the data alone.
   */
  resetWithAdminPassword(adminPassword: string, next: string): SetOutcome {
    if (!isAdminPassword(adminPassword)) {
      return { ok: false, message: 'That is not the administrator password used to install the app.' };
    }
    this.clear();
    return this.store(next);
  }

  /** Forgets the stored password. Used by the reset above, and by the tests. */
  private clear(): void {
    this.cached = undefined;
    try {
      if (existsSync(this.file)) unlinkSync(this.file);
    } catch {
      // A password file that will not delete is replaced by the write that follows.
    }
  }

  private store(password: string): SetOutcome {
    const problem = passwordProblem(password);
    if (problem) return { ok: false, message: problem };
    if (!this.vault.available()) {
      return { ok: false, message: this.status().problem };
    }
    const salt = randomBytes(16);
    const record: LockRecord = {
      version: FORMAT,
      salt: salt.toString('base64'),
      verifier: this.derive(password, salt).toString('base64'),
      machine: this.machineFingerprint(),
      setAt: new Date(this.now()).toISOString(),
      failures: 0,
      lockedUntil: 0,
    };
    if (!this.write(record)) {
      return { ok: false, message: `The password could not be saved to ${this.file}.` };
    }
    return { ok: true };
  }

  private matches(password: string, record: LockRecord): boolean {
    const expected = Buffer.from(record.verifier, 'base64');
    const candidate = this.derive(password, Buffer.from(record.salt, 'base64'));
    return candidate.length === expected.length && timingSafeEqual(candidate, expected);
  }

  /**
   * The password, turned into 32 bytes that cannot be turned back.
   *
   * The machine identifier is part of what is hashed, not just something checked alongside
   * it, so "tied to this machine" holds even if the file is unsealed by some future
   * mistake: the stored bytes are only reachable from this password *on this computer*.
   */
  private derive(password: string, salt: Buffer): Buffer {
    const material = `${password}\u0000${this.machineId()}`;
    return scryptSync(material, salt, KEY_LENGTH, SCRYPT);
  }

  private machineFingerprint(): string {
    return createHash('sha256').update(this.machineId(), 'utf8').digest('base64');
  }

  private waitLeft(record: LockRecord): number {
    return Math.max(0, record.lockedUntil - this.now());
  }

  private read(): LockRecord | undefined {
    if (this.cached) return this.cached;
    try {
      if (!existsSync(this.file)) return undefined;
      const sealed = readFileSync(this.file);
      const parsed: unknown = JSON.parse(this.vault.decrypt(sealed));
      const record = asRecord(parsed);
      this.cached = record;
      return record;
    } catch {
      // Sealed for a different Windows account or a different computer, or simply damaged.
      // Either way there is no password here that this machine can check.
      return undefined;
    }
  }

  private write(record: LockRecord): boolean {
    try {
      mkdirSync(path.dirname(this.file), { recursive: true });
      const sealed = this.vault.encrypt(JSON.stringify(record));
      // Written beside the real file and moved into place, so a power cut during a password
      // change cannot leave a half-written file that locks the app for good.
      const temporary = `${this.file}.new`;
      writeFileSync(temporary, sealed);
      renameSync(temporary, this.file);
      this.cached = record;
      return true;
    } catch {
      return false;
    }
  }
}

/**
 * Why this password will not do, or the empty string.
 *
 * Length only. A rule that demands a digit and a capital produces `Password1` and a note
 * on the monitor; a long one the teacher chose themselves is better, and the lock-out above
 * is what actually stops guessing.
 */
export function passwordProblem(password: string): string {
  if (password.length < MIN_PASSWORD_LENGTH) {
    return `The password must be at least ${MIN_PASSWORD_LENGTH} characters long.`;
  }
  if (password.trim() === '') return 'The password cannot be only spaces.';
  return '';
}

/** 30s after the sixth wrong password, doubling, never more than a quarter of an hour. */
export function waitAfter(failures: number): number {
  const doublings = failures - FREE_ATTEMPTS - 1;
  return Math.min(LONGEST_WAIT_MS, FIRST_WAIT_MS * 2 ** Math.max(0, doublings));
}

function asRecord(value: unknown): LockRecord | undefined {
  if (typeof value !== 'object' || value === null) return undefined;
  const raw = value as Record<string, unknown>;
  if (raw.version !== FORMAT) return undefined;
  const text = (key: string): string | undefined => (typeof raw[key] === 'string' ? (raw[key] as string) : undefined);
  const salt = text('salt');
  const verifier = text('verifier');
  const machine = text('machine');
  if (!salt || !verifier || !machine) return undefined;
  return {
    version: FORMAT,
    salt,
    verifier,
    machine,
    setAt: text('setAt') ?? '',
    failures: typeof raw.failures === 'number' ? raw.failures : 0,
    lockedUntil: typeof raw.lockedUntil === 'number' ? raw.lockedUntil : 0,
  };
}
