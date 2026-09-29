/**
 * The vocabulary of the sign-in screen, shared by the main process, the lock window's
 * preload bridge and the lock window itself.
 *
 * Free of runtime dependencies like the rest of `shared/`, so the preload script - which is
 * sandboxed and can `require` nothing but Electron - can still `import type` from it.
 */

/** The shortest password the app will accept. */
export const MIN_PASSWORD_LENGTH = 8;

/**
 * How many passwords may be tried before the app starts making the person wait, and how
 * long the wait then is.
 *
 * The wait doubles from 30 seconds to a quarter of an hour. Someone who has mistyped their
 * own password notices the first pause and stops; someone working through a list of guesses
 * gets a few hundred tries a day instead of a few hundred a second. The count is kept in
 * the same file as the password, so closing the app does not clear it.
 */
export const FREE_ATTEMPTS = 5;
export const FIRST_WAIT_MS = 30_000;
export const LONGEST_WAIT_MS = 15 * 60_000;

/** What the lock window opens as: a password to set, or a password to enter. */
export type LockMode = 'set' | 'enter';

export interface LockStatus {
  readonly mode: LockMode;
  /** Milliseconds left on a lock-out, or 0 when a password may be tried now. */
  readonly waitMs: number;
  /**
   * Whether the password can be checked at all. False when Windows will not encrypt the
   * file the password is kept in, which is the one condition the app cannot work around.
   */
  readonly usable: boolean;
  /** Why it is unusable, ready to print. Empty when `usable`. */
  readonly problem: string;
}

/**
 * The answer to "let me in". A plain discriminated union rather than a thrown error,
 * because it crosses an IPC boundary and because "wrong password" is an ordinary thing
 * for a person to do, not an exception.
 */
export type UnlockOutcome =
  | { readonly ok: true }
  | { readonly ok: false; readonly reason: 'wrong'; readonly attemptsLeft: number }
  | { readonly ok: false; readonly reason: 'waiting'; readonly waitMs: number }
  | { readonly ok: false; readonly reason: 'other-machine' }
  | { readonly ok: false; readonly reason: 'unreadable'; readonly message: string };

/** The answer to "here is my new password". */
export type SetOutcome =
  | { readonly ok: true }
  | { readonly ok: false; readonly message: string };

/** The operations the lock window is allowed to perform. Nothing else is exposed to it. */
export interface LockApi {
  readonly theme: string;
  status(): Promise<LockStatus>;
  /** First run only: choose the password this computer will ask for from now on. */
  setPassword(password: string): Promise<SetOutcome>;
  unlock(password: string): Promise<UnlockOutcome>;
  changePassword(current: string, next: string): Promise<SetOutcome>;
  /** Forgotten password: the administrator's installation password sets a new one. */
  resetWithAdminPassword(adminPassword: string, next: string): Promise<SetOutcome>;
  /** Closes the window and starts the app. Only ever called after a successful unlock. */
  proceed(): void;
}
