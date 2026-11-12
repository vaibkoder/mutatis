/**
 * host-web — stateless list-number formatter tests.
 *
 * These are the textbook half of the renderer and MUST pass now: `formatInt`
 * across every supported `<w:numFmt>` (incl. z→aa rollover, roman, decimalZero
 * padding, bullet/none→'') and `applyLvlText` %k substitution.
 */
import { describe, it, expect } from 'vitest';
import { formatInt, applyLvlText } from '@host-web/numbering-format.js';

describe('formatInt — decimal / decimalZero', () => {
  it('decimal renders the raw integer', () => {
    expect(formatInt(1, 'decimal')).toBe('1');
    expect(formatInt(10, 'decimal')).toBe('10');
    expect(formatInt(412, 'decimal')).toBe('412');
  });

  it('decimalZero 0-pads to a minimum width of 2', () => {
    expect(formatInt(1, 'decimalZero')).toBe('01');
    expect(formatInt(9, 'decimalZero')).toBe('09');
    expect(formatInt(10, 'decimalZero')).toBe('10');
    expect(formatInt(100, 'decimalZero')).toBe('100');
  });
});

describe('formatInt — letters (bijective base-26, z→aa rollover)', () => {
  it('lowerLetter: 1→a, 26→z, 27→aa, 28→ab, 52→az, 53→ba', () => {
    expect(formatInt(1, 'lowerLetter')).toBe('a');
    expect(formatInt(2, 'lowerLetter')).toBe('b');
    expect(formatInt(26, 'lowerLetter')).toBe('z');
    expect(formatInt(27, 'lowerLetter')).toBe('aa');
    expect(formatInt(28, 'lowerLetter')).toBe('ab');
    expect(formatInt(52, 'lowerLetter')).toBe('az');
    expect(formatInt(53, 'lowerLetter')).toBe('ba');
    expect(formatInt(702, 'lowerLetter')).toBe('zz');
    expect(formatInt(703, 'lowerLetter')).toBe('aaa');
  });

  it('upperLetter is lowerLetter upper-cased', () => {
    expect(formatInt(1, 'upperLetter')).toBe('A');
    expect(formatInt(26, 'upperLetter')).toBe('Z');
    expect(formatInt(27, 'upperLetter')).toBe('AA');
  });

  it('n<=0 is out of range for letters → ""', () => {
    expect(formatInt(0, 'lowerLetter')).toBe('');
    expect(formatInt(-1, 'upperLetter')).toBe('');
  });
});

describe('formatInt — roman', () => {
  it('lowerRoman renders subtractive numerals', () => {
    expect(formatInt(1, 'lowerRoman')).toBe('i');
    expect(formatInt(4, 'lowerRoman')).toBe('iv');
    expect(formatInt(9, 'lowerRoman')).toBe('ix');
    expect(formatInt(14, 'lowerRoman')).toBe('xiv');
    expect(formatInt(40, 'lowerRoman')).toBe('xl');
    expect(formatInt(90, 'lowerRoman')).toBe('xc');
    expect(formatInt(2024, 'lowerRoman')).toBe('mmxxiv');
  });

  it('upperRoman is lowerRoman upper-cased', () => {
    expect(formatInt(4, 'upperRoman')).toBe('IV');
    expect(formatInt(2024, 'upperRoman')).toBe('MMXXIV');
  });

  it('n<=0 → "" (roman has no zero/negative)', () => {
    expect(formatInt(0, 'lowerRoman')).toBe('');
  });
});

describe('formatInt — bullet / none / unknown → ""', () => {
  it('returns "" for non-numeric formats (documented)', () => {
    expect(formatInt(1, 'bullet')).toBe('');
    expect(formatInt(1, 'none')).toBe('');
    expect(formatInt(1, 'ordinal')).toBe(''); // unsupported style
    expect(formatInt(1, '')).toBe('');
  });
});

describe('applyLvlText — %k substitution', () => {
  it('substitutes %1..%9 with the 1-based per-level token', () => {
    expect(applyLvlText('%1.', ['4'])).toBe('4.');
    expect(applyLvlText('%1.%2', ['4', '2'])).toBe('4.2');
    expect(applyLvlText('%1.%2.%3', ['4', '2', '1'])).toBe('4.2.1');
  });

  it('keeps literal text (parens, spaces, bullet glyphs) verbatim', () => {
    expect(applyLvlText('(%1)', ['b'])).toBe('(b)');
    expect(applyLvlText('Clause %1 —', ['7'])).toBe('Clause 7 —');
    expect(applyLvlText('•', [])).toBe('•'); // a bare bullet glyph, no placeholder
  });

  it('a placeholder with no supplied token renders as ""', () => {
    expect(applyLvlText('%1.%2', ['4'])).toBe('4.');
  });

  it('leaves a literal % not followed by 1-9 untouched', () => {
    expect(applyLvlText('50% off %1', ['x'])).toBe('50% off x');
    expect(applyLvlText('%0', ['z'])).toBe('%0'); // %0 is not a valid 1-based level
  });
});
