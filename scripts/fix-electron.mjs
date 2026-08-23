/**
 * Repairs a partial Electron install.
 *
 * On locked-down networks `npm install` cannot always reach github.com to download the
 * Electron binary, and even when the download succeeds (via ELECTRON_MIRROR) the
 * post-install unzip can fail silently, leaving `node_modules/electron/dist` empty.
 * This script extracts the already-downloaded zip from Electron's cache.
 *
 *   npm run fix:electron
 */
import { createRequire } from 'node:module';
import { execFileSync } from 'node:child_process';
import { existsSync, readdirSync, readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import os from 'node:os';

const require = createRequire(import.meta.url);
const electronDir = path.dirname(require.resolve('electron/package.json'));
const version = JSON.parse(readFileSync(path.join(electronDir, 'package.json'), 'utf8')).version;
const distDir = path.join(electronDir, 'dist');
const exeName = process.platform === 'win32' ? 'electron.exe' : 'electron';

if (existsSync(path.join(distDir, exeName))) {
  console.log(`Electron ${version} is already installed.`);
  process.exit(0);
}

const cacheRoot =
  process.platform === 'win32'
    ? path.join(process.env.LOCALAPPDATA ?? os.homedir(), 'electron', 'Cache')
    : path.join(os.homedir(), '.cache', 'electron');

const zipName = `electron-v${version}-${process.platform}-${process.arch}.zip`;
const candidates = existsSync(cacheRoot)
  ? readdirSync(cacheRoot)
      .map((entry) => path.join(cacheRoot, entry, zipName))
      .filter((file) => existsSync(file))
  : [];

if (candidates.length === 0) {
  console.error(
    `Could not find ${zipName} in ${cacheRoot}.\n` +
      'Download it first, for example:\n' +
      '  set ELECTRON_MIRROR=https://registry.npmmirror.com/-/binary/electron/\n' +
      '  npm rebuild electron',
  );
  process.exit(1);
}

const zip = candidates[0];
console.log(`Extracting ${zip}`);

if (process.platform === 'win32') {
  execFileSync(
    'powershell.exe',
    [
      '-NoProfile',
      '-Command',
      `Add-Type -AssemblyName System.IO.Compression.FileSystem; ` +
        `$a=[System.IO.Compression.ZipFile]::OpenRead('${zip}'); ` +
        `foreach($e in $a.Entries){ $t=Join-Path '${distDir}' $e.FullName; ` +
        `$p=Split-Path $t -Parent; if(-not (Test-Path $p)){ New-Item -ItemType Directory -Force $p | Out-Null }; ` +
        `if(-not $e.FullName.EndsWith('/')){ [System.IO.Compression.ZipFileExtensions]::ExtractToFile($e,$t,$true) } }; ` +
        `$a.Dispose()`,
    ],
    { stdio: 'inherit' },
  );
} else {
  execFileSync('unzip', ['-o', zip, '-d', distDir], { stdio: 'inherit' });
}

writeFileSync(path.join(electronDir, 'path.txt'), exeName);
console.log(existsSync(path.join(distDir, exeName)) ? 'Electron is ready.' : 'Extraction finished but the binary is still missing.');
