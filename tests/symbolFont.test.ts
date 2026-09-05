/**
 * The Symbol encoding table is data, and a wrong entry would quietly rewrite whatever the
 * paper says. These tests pin the entries the corpus actually uses, the boundary between
 * "readable" and "nothing to read", and the one mistake that would look right until it was
 * not: Symbol's codes are *not* ASCII. 0x71 prints θ, and reading it as "q" would invent a
 * lower-case option label out of a physics variable.
 */
import { describe, expect, it } from 'vitest';
import { symbolCharacter } from '../src/core/docx/SymbolFont';

describe('reading a character out of the Symbol font', () => {
  it('reads the brackets, which is what makes an option label readable', () => {
    expect(symbolCharacter('Symbol', 'F028')).toBe('(');
    expect(symbolCharacter('Symbol', 'F029')).toBe(')');
  });

  it('accepts both spellings of the code', () => {
    // Word writes the private-use form when the font is flagged symbolic, the bare code
    // otherwise. They mean the same character.
    expect(symbolCharacter('Symbol', 'F028')).toBe(symbolCharacter('Symbol', '0028'));
    expect(symbolCharacter('Symbol', '28')).toBe('(');
  });

  it('reads the Greek letters, not the ASCII letters at the same codes', () => {
    expect(symbolCharacter('Symbol', 'F071')).toBe('θ');
    expect(symbolCharacter('Symbol', 'F077')).toBe('ω');
    expect(symbolCharacter('Symbol', 'F061')).toBe('α');
    expect(symbolCharacter('Symbol', 'F062')).toBe('β');
    expect(symbolCharacter('Symbol', 'F057')).toBe('Ω');
  });

  it('reads the operators and arrows the papers use', () => {
    expect(symbolCharacter('Symbol', 'F0B0')).toBe('°');
    expect(symbolCharacter('Symbol', 'F0B4')).toBe('×');
    expect(symbolCharacter('Symbol', 'F0AE')).toBe('→');
    expect(symbolCharacter('Symbol', 'F02D')).toBe('−');
    expect(symbolCharacter('Symbol', 'F0A2')).toBe('′');
    expect(symbolCharacter('Symbol', 'F0A5')).toBe('∞');
    expect(symbolCharacter('Symbol', 'F0D6')).toBe('√');
    expect(symbolCharacter('Symbol', 'F0F2')).toBe('∫');
  });

  it('keeps the digits and punctuation Symbol shares with ASCII', () => {
    expect(symbolCharacter('Symbol', 'F030')).toBe('0');
    expect(symbolCharacter('Symbol', 'F03D')).toBe('=');
    expect(symbolCharacter('Symbol', 'F020')).toBe(' ');
  });

  it('reads nothing from a glyph piece, which is ink rather than a character', () => {
    // 0xE6-0xF0 and 0xF3 onwards are the top, middle and bottom of a tall bracket or
    // integral sign. None of them is something an author could have typed.
    expect(symbolCharacter('Symbol', 'F0E6')).toBeUndefined();
    expect(symbolCharacter('Symbol', 'F0EF')).toBeUndefined();
    expect(symbolCharacter('Symbol', 'F0F4')).toBeUndefined();
    expect(symbolCharacter('Symbol', 'F060')).toBeUndefined();
    expect(symbolCharacter('Symbol', 'F0BD')).toBeUndefined();
  });

  it('reads nothing from an unused code', () => {
    expect(symbolCharacter('Symbol', 'F080')).toBeUndefined();
    expect(symbolCharacter('Symbol', 'F09F')).toBeUndefined();
    expect(symbolCharacter('Symbol', 'F01F')).toBeUndefined();
  });

  it('reads nothing from a picture font, whose codes name drawings', () => {
    expect(symbolCharacter('Wingdings', 'F028')).toBeUndefined();
    expect(symbolCharacter('Wingdings 2', 'F028')).toBeUndefined();
    expect(symbolCharacter('Webdings', 'F028')).toBeUndefined();
    expect(symbolCharacter('MT Extra', 'F028')).toBeUndefined();
  });

  it('does not care how the font name is capitalised or spaced', () => {
    expect(symbolCharacter('symbol', 'F028')).toBe('(');
    expect(symbolCharacter(' Symbol ', 'F028')).toBe('(');
  });

  it('reads nothing when the code is missing or not a number', () => {
    expect(symbolCharacter('Symbol', '')).toBeUndefined();
    expect(symbolCharacter('Symbol', 'not-hex')).toBeUndefined();
  });
});
