/**
 * Reads the character that a `w:sym` prints.
 *
 * Word writes `<w:sym w:font="Symbol" w:char="F028"/>` when a character is inserted with
 * Insert > Symbol. There is no `w:t`, so a naive reading of the paragraph loses the
 * character: `(A)` reads as `A)`, and `at an angle θ` reads as `at an angle `.
 *
 * That character is not, however, unknowable. `w:char` is the code point *in the font's own
 * encoding*, and the Symbol font's encoding is a published standard - the same Adobe
 * specification that gives `Pdf.ts` its widths. Code 0x28 in Symbol is `parenleft`: it is a
 * bracket, and saying so is reading the file, not guessing at it.
 *
 * Only Symbol is decoded. Wingdings, Webdings and the like are picture fonts whose codes
 * name arrows, hands and boxes rather than characters, so there is nothing to read; a
 * question whose option boundary depends on one of those is still refused.
 *
 * Word stores the code in the private-use area (0xF000 + code) when the font is flagged as
 * symbolic, and as the bare code otherwise. Both spellings are accepted.
 */

/**
 * Adobe's Symbol encoding, written as the runs of consecutive codes that have a character
 * to read. Each entry is the code of its first character.
 *
 * The gaps are deliberate. 0x60, 0xBD-0xBE and 0xE6-0xF0 are not characters at all: they
 * are the *pieces* Symbol carries for drawing a tall bracket, brace or integral sign out of
 * a top, a middle and a bottom. A piece is a fragment of ink rather than something a reader
 * could type, so it reads as nothing - exactly as every `w:sym` did before this table.
 */
const RUNS: readonly (readonly [number, string])[] = [
  // 0x21 ! through 0x5F _ : punctuation, digits, and the upper-case Greek alphabet.
  [0x21, '!∀#∃%&∋()∗+,−./0123456789:;<=>?≅ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ[∴]⊥_'],
  // 0x61 α through 0x7E ∼ : the lower-case Greek alphabet and the remaining brackets.
  [0x61, 'αβχδεφγηιϕκλμνοπθρστυϖωξψζ{|}∼'],
  // 0xA0 € through 0xBC … : currency, card suits, arrows and the common operators.
  [0xa0, '€ϒ′≤⁄∞ƒ♣♦♥♠↔←↑→↓°±″≥×∝∂•÷≠≡≈…'],
  // 0xBF ↵ through 0xE5 ∑ : set theory, logic, double arrows, and both ®©™ families.
  [0xbf, '↵ℵℑℜ℘⊗⊕∅∩∪⊃⊇⊄⊂⊆∈∉∠∇®©™∏√⋅¬∧∨⇔⇐⇑⇒⇓◊〈®©™∑'],
  // 0xF1 〉 and 0xF2 ∫ , sitting among the bracket pieces.
  [0xf1, '〉∫'],
];

/** Symbol's space, which no run above can hold without the runs becoming unreadable. */
const SPACE_CODE = 0x20;

/** Word's private-use block for a symbolic font: 0xF000 plus the code within the font. */
const PRIVATE_USE = 0xf000;

const BY_CODE = new Map<number, string>([[SPACE_CODE, ' ']]);
for (const [first, characters] of RUNS) {
  [...characters].forEach((character, offset) => BY_CODE.set(first + offset, character));
}

/**
 * The character `char` prints in `font`, or `undefined` when the file does not say.
 *
 * `char` is the raw `w:char` attribute - hexadecimal, as Word writes it.
 */
export function symbolCharacter(font: string, char: string): string | undefined {
  if (font.trim().toLowerCase() !== 'symbol') return undefined;
  const raw = Number.parseInt(char, 16);
  if (!Number.isFinite(raw)) return undefined;
  const code = raw >= PRIVATE_USE && raw <= PRIVATE_USE + 0xff ? raw - PRIVATE_USE : raw;
  return BY_CODE.get(code);
}
