import { createHash, timingSafeEqual } from 'node:crypto';

/**
 * The installing administrator's password, as a salted SHA-256 digest.
 *
 * The same digest is compiled into the MSI (see `scripts/build-msi.mjs`, which reads it
 * from here so the two can never drift) and into the app, where it is the only way back in
 * for a teacher who has forgotten the password they set.
 *
 * **What this does and does not protect.** The password is a constant that ships inside
 * every copy of the installer and the app, so it cannot be a secret in the cryptographic
 * sense: anyone determined enough to disassemble the package can work out what to type.
 * Storing the digest rather than the word itself means it is not simply *readable* - opening
 * the MSI in Orca, or `dist/main/AdminSecret.js` in Notepad, shows 64 hex characters and no
 * password - which is what stops it being passed around the staff room. It is a lock on the
 * door, not a vault: it keeps the installer out of the hands of people who were not given
 * the password, and that is all it is for.
 *
 * A salt is included so the digest cannot be recognised in a table of hashes of common
 * words, and so that the same digest never appears in two different builds of this app.
 *
 * To change the password, run `npm run set-admin-password` and rebuild the installer.
 */
const ADMIN_SALT = 'xkgVjx82qCmEcB1o0l/KXw==';
const ADMIN_DIGEST = '808fc782ccf2a2201bfddaf4b45f374e2f04e0fda855112991d1dc9af60f99e8';

/**
 * The digest of a candidate password, in the form the installer also computes.
 *
 * Exported because the MSI's own check has to produce byte-identical output from an
 * entirely separate implementation - a hand-written SHA-256 in the installer's scripting
 * language - and `tests/adminSecret.test.ts` pins the two against each other.
 */
export function adminDigestOf(password: string): string {
  const salted = Buffer.concat([Buffer.from(ADMIN_SALT, 'base64'), Buffer.from(password, 'utf8')]);
  return createHash('sha256').update(salted).digest('hex');
}

/** The salt, as the installer needs it: the bytes it prefixes to what was typed. */
export const ADMIN_SALT_BASE64 = ADMIN_SALT;
/** The digest the installer compares against. */
export const ADMIN_PASSWORD_DIGEST = ADMIN_DIGEST;

/**
 * Whether `password` is the administrator's password.
 *
 * Compared byte by byte in constant time. Not because the timing of a local comparison is
 * a realistic way in - it is not - but because the alternative costs nothing and this is
 * the one place in the app where a comparison is a decision about access.
 */
export function isAdminPassword(password: string): boolean {
  const candidate = Buffer.from(adminDigestOf(password), 'hex');
  const expected = Buffer.from(ADMIN_DIGEST, 'hex');
  return candidate.length === expected.length && timingSafeEqual(candidate, expected);
}
