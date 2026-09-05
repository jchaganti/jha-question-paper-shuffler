/**
 * Draws the application icon: `build/icon.ico`.
 *
 * The icon is the same idea as the window watermark (`src/renderer/paper-stack.svg`) -
 * three question papers fanned out, each with a different answer filled in - because that
 * is what the tool does, and because the icon should look like the app it opens.
 *
 * It is drawn here rather than exported from a drawing program so that the repository has
 * no binary asset nobody can edit, and so that the icon can be regenerated after a change
 * of palette. There is no image library: pixels are written directly, a PNG is deflated
 * with `node:zlib`, and the ICO container is assembled by hand (Windows Vista onwards
 * reads PNG-compressed icon entries, which is what keeps this small).
 *
 *   node scripts/make-icon.mjs
 */
import { deflateSync } from 'node:zlib';
import { mkdir, writeFile } from 'node:fs/promises';
import path from 'node:path';

/** Sizes Windows asks for: the shell picks the nearest at each display size. */
const SIZES = [16, 24, 32, 48, 64, 128, 256];

/** Warm sand behind ink-blue paper, matching the app window. */
const INK = [58, 84, 150];
const INK_SOFT = [104, 128, 190];
const PAPER = [255, 253, 248];
const BACKDROP = [238, 228, 206];
const BACKDROP_EDGE = [223, 209, 178];

/** The three sheets, in the drawing order of the watermark: back, middle, front. */
const SHEETS = [
  { cx: -0.215, cy: 0.02, angle: -11, filled: 1, shade: 0.72 },
  { cx: 0.0, cy: -0.025, angle: -1.5, filled: 3, shade: 0.86 },
  { cx: 0.215, cy: 0.025, angle: 9, filled: 0, shade: 1 },
];

const SHEET_W = 0.375;
const SHEET_H = 0.52;
const CORNER = 0.035;

const mix = (a, b, t) => a.map((value, i) => value + (b[i] - value) * t);

/** Distance-based coverage of a rounded rectangle centred on the origin, in icon units. */
function roundedRect(x, y, halfW, halfH, radius) {
  const dx = Math.abs(x) - (halfW - radius);
  const dy = Math.abs(y) - (halfH - radius);
  if (dx <= 0 && dy <= 0) return -Math.min(-dx, -dy) - radius;
  const px = Math.max(dx, 0);
  const py = Math.max(dy, 0);
  return Math.hypot(px, py) - radius;
}

/** Colour and alpha at one point of the icon, in units where the icon spans -0.5..0.5. */
function sample(x, y) {
  // Rounded-square backdrop, with a hairline edge so the icon reads on any wallpaper.
  const outer = roundedRect(x, y, 0.5, 0.5, 0.115);
  if (outer > 0) return null;
  let colour = outer > -0.012 ? BACKDROP_EDGE : BACKDROP;

  for (const sheet of SHEETS) {
    const theta = (sheet.angle * Math.PI) / 180;
    const cos = Math.cos(theta);
    const sin = Math.sin(theta);
    const ox = x - sheet.cx;
    const oy = y - sheet.cy;
    // Into the sheet's own frame.
    const sx = ox * cos + oy * sin;
    const sy = -ox * sin + oy * cos;

    const edge = roundedRect(sx, sy, SHEET_W / 2, SHEET_H / 2, CORNER);
    if (edge > 0.006) continue;
    if (edge > -0.008) {
      // The sheet's own outline.
      colour = mix(INK, PAPER, 1 - sheet.shade * 0.75);
      continue;
    }
    colour = mix(PAPER, INK, (1 - sheet.shade) * 0.14);

    // Two heading rules near the top of the sheet.
    const left = -SHEET_W / 2 + 0.055;
    for (const [index, width] of [[0, 0.2], [1, 0.13]]) {
      const ruleY = -SHEET_H / 2 + 0.075 + index * 0.045;
      if (Math.abs(sy - ruleY) < 0.011 && sx > left && sx < left + width) {
        colour = mix(colour, INK_SOFT, sheet.shade);
      }
    }

    // Four option rows: a bullet and a rule, with one bullet filled in.
    for (let row = 0; row < 4; row++) {
      const rowY = -SHEET_H / 2 + 0.215 + row * 0.072;
      const bulletX = left + 0.012;
      const toBullet = Math.hypot(sx - bulletX, sy - rowY);
      const answer = row === sheet.filled;
      if (toBullet < 0.026) {
        const ring = answer || toBullet > 0.015;
        if (ring) colour = mix(colour, answer ? INK : INK_SOFT, sheet.shade);
        continue;
      }
      const width = answer ? 0.185 : 0.12 + row * 0.018;
      if (Math.abs(sy - rowY) < 0.012 && sx > bulletX + 0.038 && sx < bulletX + 0.038 + width) {
        colour = mix(colour, answer ? INK : INK_SOFT, sheet.shade * (answer ? 1 : 0.75));
      }
    }
  }

  return colour;
}

/** RGBA pixels for one square size, supersampled so the edges are smooth. */
function render(size) {
  const grid = size <= 32 ? 4 : 3;
  const pixels = Buffer.alloc(size * size * 4);
  for (let py = 0; py < size; py++) {
    for (let px = 0; px < size; px++) {
      let r = 0;
      let g = 0;
      let b = 0;
      let hits = 0;
      for (let sy = 0; sy < grid; sy++) {
        for (let sx = 0; sx < grid; sx++) {
          const x = (px + (sx + 0.5) / grid) / size - 0.5;
          const y = (py + (sy + 0.5) / grid) / size - 0.5;
          const colour = sample(x, y);
          if (!colour) continue;
          r += colour[0];
          g += colour[1];
          b += colour[2];
          hits++;
        }
      }
      const at = (py * size + px) * 4;
      if (hits === 0) continue;
      pixels[at] = Math.round(r / hits);
      pixels[at + 1] = Math.round(g / hits);
      pixels[at + 2] = Math.round(b / hits);
      pixels[at + 3] = Math.round((hits / (grid * grid)) * 255);
    }
  }
  return pixels;
}

const crcTable = Array.from({ length: 256 }, (_unused, n) => {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c >>> 0;
});

function crc32(buffer) {
  let c = 0xffffffff;
  for (const byte of buffer) c = crcTable[(c ^ byte) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function chunk(type, body) {
  const length = Buffer.alloc(4);
  length.writeUInt32BE(body.length);
  const typed = Buffer.concat([Buffer.from(type, 'latin1'), body]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(typed));
  return Buffer.concat([length, typed, crc]);
}

/** A minimal 8-bit RGBA PNG: one IHDR, one IDAT, one IEND. */
function png(size, pixels) {
  const header = Buffer.alloc(13);
  header.writeUInt32BE(size, 0);
  header.writeUInt32BE(size, 4);
  header[8] = 8; // bit depth
  header[9] = 6; // colour type: RGBA
  // Each scanline is prefixed with its filter type; 0 (None) keeps this readable, and the
  // images are small enough that a cleverer filter would save nothing worth the code.
  const raw = Buffer.alloc(size * (size * 4 + 1));
  for (let row = 0; row < size; row++) {
    raw[row * (size * 4 + 1)] = 0;
    pixels.copy(raw, row * (size * 4 + 1) + 1, row * size * 4, (row + 1) * size * 4);
  }
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', header),
    chunk('IDAT', deflateSync(raw, { level: 9 })),
    chunk('IEND', Buffer.alloc(0)),
  ]);
}

/** ICONDIR + one ICONDIRENTRY per size, each pointing at a PNG payload. */
function ico(images) {
  const header = Buffer.alloc(6);
  header.writeUInt16LE(0, 0);
  header.writeUInt16LE(1, 2); // type: icon
  header.writeUInt16LE(images.length, 4);

  const entries = [];
  let offset = 6 + images.length * 16;
  for (const { size, data } of images) {
    const entry = Buffer.alloc(16);
    entry[0] = size >= 256 ? 0 : size; // 0 means 256
    entry[1] = size >= 256 ? 0 : size;
    entry[2] = 0; // colours in palette
    entry[3] = 0; // reserved
    entry.writeUInt16LE(1, 4); // colour planes
    entry.writeUInt16LE(32, 6); // bits per pixel
    entry.writeUInt32LE(data.length, 8);
    entry.writeUInt32LE(offset, 12);
    entries.push(entry);
    offset += data.length;
  }

  return Buffer.concat([header, ...entries, ...images.map((image) => image.data)]);
}

const images = SIZES.map((size) => ({ size, data: png(size, render(size)) }));
const out = path.resolve('build');
await mkdir(out, { recursive: true });
await writeFile(path.join(out, 'icon.ico'), ico(images));
// electron-builder also uses a 512px PNG for the installer sidebar and Linux builds.
await writeFile(path.join(out, 'icon.png'), png(512, render(512)));
console.log(`wrote ${path.join(out, 'icon.ico')} (${SIZES.join(', ')} px) and icon.png`);
