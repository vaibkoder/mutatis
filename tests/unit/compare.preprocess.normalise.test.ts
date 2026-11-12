import { describe, it, expect } from 'vitest';

import {
  normaliseText,
  normaliseTextAggressive,
  stripNumberingPrefix,
} from '@features/compare/preprocess/index.js';

describe('normaliseText (§5.2)', () => {
  it('collapses whitespace runs and trims', () => {
    expect(normaliseText('  The   Seller \t shall\npay  ')).toBe('The Seller shall pay');
  });

  it('normalises NBSP to a plain space', () => {
    expect(normaliseText('The Seller')).toBe('The Seller');
  });

  it('normalises curly quotes to straight quotes', () => {
    expect(normaliseText('“Permitted” and ‘x’')).toBe('"Permitted" and \'x\'');
  });

  it('normalises em-dash to spaced double hyphen and en-dash to hyphen', () => {
    expect(normaliseText('a—b')).toBe('a -- b');
    expect(normaliseText('2020–2021')).toBe('2020-2021');
  });

  it('preserves case (defined terms / spelling are case-sensitive)', () => {
    expect(normaliseText('Permitted Encumbrance')).toBe('Permitted Encumbrance');
  });
});

describe('normaliseTextAggressive (§5.2)', () => {
  it('lowercases and strips punctuation', () => {
    expect(normaliseTextAggressive('The Seller, shall pay!')).toBe('the seller shall pay');
  });

  it('makes spacing/punctuation-only differences compare equal', () => {
    const a = normaliseTextAggressive('The Seller shall pay.');
    const b = normaliseTextAggressive('The  Seller, shall pay');
    expect(a).toBe(b);
  });

  it('keeps a genuine word change distinct', () => {
    expect(normaliseTextAggressive('shall pay')).not.toBe(normaliseTextAggressive('may pay'));
  });
});

describe('stripNumberingPrefix (§5.0)', () => {
  it('strips a compound decimal with bracketed sub-parts and keeps trailing space in prefix', () => {
    const { prefix, body } = stripNumberingPrefix('4.2(a) The Seller shall pay');
    expect(prefix).toBe('4.2(a) ');
    expect(body).toBe('The Seller shall pay');
  });

  it('strips a deeper compound decimal', () => {
    const { prefix, body } = stripNumberingPrefix('4.2.1 Definitions apply');
    expect(prefix).toBe('4.2.1 ');
    expect(body).toBe('Definitions apply');
  });

  it('strips nested bracketed sub-parts (a)(i)', () => {
    const { prefix, body } = stripNumberingPrefix('4.2(a)(i) sub-obligation');
    expect(prefix).toBe('4.2(a)(i) ');
    expect(body).toBe('sub-obligation');
  });

  it('strips a bracketed token (iii)', () => {
    const { prefix, body } = stripNumberingPrefix('(iii) third item');
    expect(prefix).toBe('(iii) ');
    expect(body).toBe('third item');
  });

  it('strips a plain number with a trailing dot', () => {
    const { prefix, body } = stripNumberingPrefix('12. Final clause');
    expect(prefix).toBe('12. ');
    expect(body).toBe('Final clause');
  });

  it('leaves a bare leading integer (ambiguous prose) untouched', () => {
    const { prefix, body } = stripNumberingPrefix('4 horsemen rode in');
    expect(prefix).toBe('');
    expect(body).toBe('4 horsemen rode in');
  });

  it('returns empty prefix when there is no numbering', () => {
    const { prefix, body } = stripNumberingPrefix('The Seller shall pay');
    expect(prefix).toBe('');
    expect(body).toBe('The Seller shall pay');
  });
});
