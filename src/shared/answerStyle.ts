/**
 * How a paper names one of the four options.
 *
 * Papers do not agree: letters (`A`, `(A)`, `a.`), digits naming the option's position
 * (`1`, `2.`), or roman numerals doing the same (`iii)`, `(IV)`). The style is read off the
 * document rather than assumed, and everything written back for the user - the answer key
 * cells and the generation report alike - is rendered through the style that was found, so
 * the output speaks the same vocabulary as the paper.
 *
 * Internally an option is always identified by its slot letter A-D. This module is the
 * single boundary between that vocabulary and the paper's own.
 *
 * Free of runtime dependencies so the main process, the CLI and the renderer can all use it.
 */

import { OPTION_LETTERS, type OptionLetter } from './types';

/** The family of symbol a paper uses to name one of the four options. */
export type AnswerScheme = 'letter' | 'digit' | 'roman';

/** A scheme plus the case and decoration found alongside it, e.g. `(` `iii` `)`. */
export interface AnswerStyle {
  readonly scheme: AnswerScheme;
  readonly upperCase: boolean;
  readonly prefix: string;
  readonly suffix: string;
}

/** A bare upper-case letter - the style assumed when a paper states none. */
export const PLAIN_LETTER_STYLE: AnswerStyle = { scheme: 'letter', upperCase: true, prefix: '', suffix: '' };

/** "i", "ii", "iii", "iv" - the roman numerals naming options A-D, lower case. */
const ROMAN_BY_INDEX = ['i', 'ii', 'iii', 'iv'] as const;

/**
 * Reads one written answer, deducing how it names an option: a bare or decorated letter
 * ("A", "(A)", "a."), a digit naming the option's position ("1", "2)"), or a roman numeral
 * doing the same ("iii)", "(IV)"). Returns the canonical A-D slot plus the exact style
 * found, so the same value can be written back the way it was read.
 */
export function parseAnswer(raw: string): { letter: OptionLetter; style: AnswerStyle } | undefined {
  const text = raw.trim();
  if (!text) return undefined;

  const bracketed = /^\((.+)\)$/.exec(text);
  const prefix = bracketed ? '(' : '';
  const suffix = bracketed ? ')' : (/[.)]$/.exec(text)?.[0] ?? '');
  const core = (bracketed ? bracketed[1]! : suffix ? text.slice(0, -suffix.length) : text).trim();

  if (/^[A-Da-d]$/.test(core)) {
    return {
      letter: core.toUpperCase() as OptionLetter,
      style: { scheme: 'letter', upperCase: core === core.toUpperCase(), prefix, suffix },
    };
  }
  if (/^[1-4]$/.test(core)) {
    return {
      letter: OPTION_LETTERS[Number(core) - 1]!,
      style: { scheme: 'digit', upperCase: true, prefix, suffix },
    };
  }
  const romanIndex = ROMAN_BY_INDEX.indexOf(core.toLowerCase() as (typeof ROMAN_BY_INDEX)[number]);
  if (romanIndex >= 0) {
    return {
      letter: OPTION_LETTERS[romanIndex]!,
      style: { scheme: 'roman', upperCase: core === core.toUpperCase(), prefix, suffix },
    };
  }
  return undefined;
}

/** Renders slot `letter` in the exact scheme, case and decoration of `style`. */
export function renderAnswer(letter: OptionLetter, style: AnswerStyle = PLAIN_LETTER_STYLE): string {
  const index = OPTION_LETTERS.indexOf(letter);
  const core =
    style.scheme === 'digit'
      ? String(index + 1)
      : style.scheme === 'roman'
        ? style.upperCase
          ? ROMAN_BY_INDEX[index]!.toUpperCase()
          : ROMAN_BY_INDEX[index]!
        : style.upperCase
          ? letter
          : letter.toLowerCase();
  return `${style.prefix}${core}${style.suffix}`;
}
