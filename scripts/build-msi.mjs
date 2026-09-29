// Builds the MSI installer: the one that asks for the administrator's password.
//
//   npm run msi        # build, draw the icon, package, then everything below
//
// Three steps. The app is packaged into `release/win-unpacked` by electron-builder (the
// `pack` script, run first). The file list is worked out from that folder and written as
// WiX source, because WiX has to be told about every file by name. Then WiX is run twice -
// candle to compile, light to link - and the result lands in `release/`.
//
// The administrator's password is not in this file, nor in the WiX source, nor in the MSI:
// what is passed to WiX is the salt and the SHA-256 digest from `src/main/AdminSecret.ts`,
// read from the compiled build so that the app and the installer cannot disagree about it.
import { execFileSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';

const root = path.resolve();
const appDir = path.join(root, 'release', 'win-unpacked');
const installerDir = path.join(root, 'installer');
const stageDir = path.join(root, 'release', 'msi-build');

const { version, productName } = JSON.parse(readFileSync(path.join(root, 'package.json'), 'utf8'));
const name = productName ?? 'Question Paper Shuffler';
const outputFile = path.join(root, 'release', `${name} Setup ${version}.msi`);

/* ------------------------------------------------------------------- WiX -- */

/**
 * Where candle.exe and light.exe are.
 *
 * WiX is not an npm package and is not installed with the rest of the toolchain, so this
 * looks in the three places it can reasonably be and, failing that, says exactly what to do
 * rather than failing with a missing-file error from a program the reader has never run.
 */
function findWix() {
  const candidates = [
    process.env.WIX_DIR,
    process.env.WIX ? path.join(process.env.WIX, 'bin') : undefined,
    'C:\\Program Files (x86)\\WiX Toolset v3.14\\bin',
    'C:\\Program Files (x86)\\WiX Toolset v3.11\\bin',
    ...electronBuilderWixCache(),
  ].filter(Boolean);

  for (const folder of candidates) {
    if (existsSync(path.join(folder, 'candle.exe')) && existsSync(path.join(folder, 'light.exe'))) {
      return folder;
    }
  }

  throw new Error(
    [
      'WiX was not found, so the MSI cannot be built. It is the toolset that turns',
      `${path.relative(root, path.join(installerDir, 'Product.wxs'))} into an installer, and it is not part of npm install.`,
      '',
      'Do one of these, then run `npm run msi` again:',
      '',
      '  1. Install the WiX Toolset v3.14 from https://github.com/wixtoolset/wix3/releases',
      '     (the file called wix314.exe). It puts candle.exe and light.exe where this script',
      '     looks for them.',
      '',
      '  2. If you already have a copy of those two programs, point this at the folder that',
      '     holds them:   set WIX_DIR=D:\\path\\to\\wix\\bin',
      '',
      'Looked in:',
      ...candidates.map((folder) => `  ${folder}`),
    ].join('\n'),
  );
}

/**
 * electron-builder downloads its own copy of WiX for its `msi` target and leaves it in its
 * cache. If it is there, use it - that is one less thing for whoever builds this to install.
 */
function electronBuilderWixCache() {
  const cache = path.join(os.homedir(), 'AppData', 'Local', 'electron-builder', 'Cache');
  if (!existsSync(cache)) return [];
  const found = [];
  for (const entry of readdirSync(cache)) {
    if (!entry.startsWith('wix-')) continue;
    const folder = path.join(cache, entry);
    found.push(folder, ...readdirSync(folder).map((inner) => path.join(folder, inner)));
  }
  return found.filter((folder) => statSync(folder).isDirectory());
}

/* -------------------------------------------------------------- the files -- */

/**
 * A name WiX will accept as an identifier: letters, digits, underscores and dots, starting
 * with a letter or an underscore, and no longer than 72 characters.
 *
 * The path's own hash goes on the end. Two files called `LICENSE` in different folders would
 * otherwise collide, and a name that is *stable* between builds matters more than a readable
 * one: it is what lets Windows recognise an upgrade as the same file rather than a new one.
 */
function identifier(prefix, relativePath) {
  const digest = createHash('sha1').update(relativePath.replace(/\\/g, '/')).digest('hex').slice(0, 10);
  const readable = path
    .basename(relativePath)
    .replace(/[^A-Za-z0-9_.]/g, '_')
    .slice(0, 40);
  return `${prefix}_${readable}_${digest}`.replace(/^([^A-Za-z_])/, '_$1');
}

const escape = (text) =>
  text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/**
 * Walks the packaged app and writes it out as WiX source: one `<Component>` per file, which
 * is what Windows Installer wants - a component with a single file as its key path can be
 * given an automatic GUID, repaired on its own, and left alone by an upgrade that did not
 * change it.
 */
function harvest() {
  const components = [];

  function walk(folder, relative, indent) {
    const lines = [];
    const entries = readdirSync(folder, { withFileTypes: true }).sort((a, b) => a.name.localeCompare(b.name));
    const pad = '  '.repeat(indent);

    for (const entry of entries) {
      const child = path.join(folder, entry.name);
      const childRelative = relative ? path.join(relative, entry.name) : entry.name;
      if (entry.isDirectory()) {
        const dirId = identifier('dir', childRelative);
        lines.push(`${pad}<Directory Id="${dirId}" Name="${escape(entry.name)}">`);
        lines.push(...walk(child, childRelative, indent + 1));
        lines.push(`${pad}</Directory>`);
      } else {
        const componentId = identifier('cmp', childRelative);
        const fileId = identifier('fil', childRelative);
        components.push(componentId);
        lines.push(`${pad}<Component Id="${componentId}" Guid="*" Win64="yes">`);
        lines.push(
          `${pad}  <File Id="${fileId}" Name="${escape(entry.name)}" Source="${escape(child)}" KeyPath="yes" />`,
        );
        lines.push(`${pad}</Component>`);
      }
    }
    return lines;
  }

  const tree = walk(appDir, '', 3);
  const xml = [
    '<?xml version="1.0" encoding="utf-8"?>',
    '<!-- Written by scripts/build-msi.mjs from release/win-unpacked. Do not edit. -->',
    '<Wix xmlns="http://schemas.microsoft.com/wix/2006/wi">',
    '  <Fragment>',
    '    <DirectoryRef Id="INSTALLFOLDER">',
    ...tree,
    '    </DirectoryRef>',
    '  </Fragment>',
    '  <Fragment>',
    '    <ComponentGroup Id="AppFiles">',
    ...components.map((id) => `      <ComponentRef Id="${id}" />`),
    '    </ComponentGroup>',
    '  </Fragment>',
    '</Wix>',
    '',
  ].join('\n');

  const file = path.join(stageDir, 'AppFiles.wxs');
  writeFileSync(file, xml, 'utf8');
  return { file, count: components.length };
}

/* ------------------------------------------------------------ the secret -- */

/**
 * The salt and digest the installer checks against, taken from the app's own build.
 *
 * Read from `dist/main/AdminSecret.js` rather than copied into this script, so that changing
 * the administrator's password is one edit in one file. The salt is handed over as hex
 * because the installer's check is written in a scripting language with no base64 decoder
 * worth trusting - see `installer/CheckPassword.js`.
 */
function adminSecret() {
  const compiled = path.join(root, 'dist', 'main', 'AdminSecret.js');
  if (!existsSync(compiled)) {
    throw new Error(`${path.relative(root, compiled)} is missing. Run \`npm run build\` first.`);
  }
  const { ADMIN_SALT_BASE64, ADMIN_PASSWORD_DIGEST } = createRequire(import.meta.url)(compiled);
  return {
    saltHex: Buffer.from(ADMIN_SALT_BASE64, 'base64').toString('hex'),
    digest: ADMIN_PASSWORD_DIGEST,
  };
}

/* ----------------------------------------------------------------- build -- */

function run(program, args) {
  console.log(`  ${path.basename(program)} ${args.filter((a) => !a.startsWith('-d')).join(' ')}`);
  execFileSync(program, args, { cwd: root, stdio: 'inherit', windowsHide: true });
}

if (!existsSync(appDir)) {
  throw new Error(
    `${path.relative(root, appDir)} is missing. Run \`npm run pack\` first - \`npm run msi\` does it for you.`,
  );
}

const wix = findWix();
mkdirSync(stageDir, { recursive: true });

const { file: filesSource, count } = harvest();
console.log(`packaged app     : ${path.relative(root, appDir)} (${count} files)`);
console.log(`WiX              : ${wix}`);

const { saltHex, digest } = adminSecret();

const defines = [
  `-dVersion=${version}`,
  `-dAppDir=${appDir}`,
  `-dInstallerDir=${installerDir}`,
  `-dAdminSalt=${saltHex}`,
  `-dAdminDigest=${digest}`,
];

const productObject = path.join(stageDir, 'Product.wixobj');
const filesObject = path.join(stageDir, 'AppFiles.wixobj');

run(path.join(wix, 'candle.exe'), [
  '-nologo',
  '-arch',
  'x64',
  ...defines,
  '-out',
  `${stageDir}\\`,
  path.join(installerDir, 'Product.wxs'),
  filesSource,
]);

// `-sval` skips the post-build validation suite. It needs pieces of the Windows Installer
// SDK that the downloaded WiX bundle does not always carry, and a missing .cub file would
// fail a build that is otherwise fine. Set MSI_VALIDATE=1 to turn it back on.
const validation = process.env.MSI_VALIDATE === '1' ? [] : ['-sval'];

run(path.join(wix, 'light.exe'), [
  '-nologo',
  '-ext',
  'WixUIExtension',
  ...validation,
  '-spdb',
  '-out',
  outputFile,
  productObject,
  filesObject,
]);

console.log(`\nbuilt            : ${path.relative(root, outputFile)}`);
console.log('The installer asks for the administrator password before it will install.');
