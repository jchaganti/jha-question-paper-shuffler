import { execFileSync } from 'node:child_process';
import os from 'node:os';
import { safeStorage } from 'electron';
import type { MachineId, Vault } from './AppLock';

/**
 * The real encryption behind the password file.
 *
 * On Windows `safeStorage` is DPAPI, which encrypts with a key the operating system keeps
 * for the signed-in account and never hands out. That is what makes the file worthless
 * anywhere else: copied to another computer, or opened under another Windows account on
 * this one, it simply will not decrypt.
 */
export const electronVault: Vault = {
  available: () => safeStorage.isEncryptionAvailable(),
  encrypt: (plain) => safeStorage.encryptString(plain),
  decrypt: (sealed) => safeStorage.decryptString(sealed),
};

/**
 * What identifies this computer.
 *
 * Windows keeps a GUID for the installation in the registry; it is written when Windows is
 * installed and does not change when the machine is renamed, joined to a domain or given a
 * new network card, which is what makes it the right thing to tie a password to. Reading it
 * costs one `reg query` at startup, and the answer is remembered for the life of the process.
 *
 * If the registry cannot be read - a locked-down machine, or a platform that is not Windows
 * - the host name and account name stand in. Weaker, because both can be changed, but the
 * file is still sealed by the operating system either way; this identifier is the second
 * lock, not the first.
 */
let remembered: string | undefined;

export const windowsMachineId: MachineId = () => {
  remembered ??= readMachineGuid() ?? `${os.hostname()}\u0000${os.userInfo().username}`;
  return remembered;
};

function readMachineGuid(): string | undefined {
  if (process.platform !== 'win32') return undefined;
  try {
    const output = execFileSync(
      'reg',
      ['query', 'HKLM\\SOFTWARE\\Microsoft\\Cryptography', '/v', 'MachineGuid'],
      { encoding: 'utf8', windowsHide: true, timeout: 5_000 },
    );
    // "    MachineGuid    REG_SZ    6a7c...-..."
    const guid = /MachineGuid\s+REG_SZ\s+(\S+)/.exec(output)?.[1];
    return guid && guid.length >= 8 ? guid : undefined;
  } catch {
    return undefined;
  }
}
