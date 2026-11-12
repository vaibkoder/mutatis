/**
 * Shared normalisation (spec §4.4) — T2 acceptance.
 */
import { describe, it, expect } from 'vitest';
import { normaliseNumber, normaliseSubToken, stripParens } from '@shared/section-index/normalise.js';

describe('normaliseNumber — auto context', () => {
  it('decimal tokens are unchanged', () => {
    expect(normaliseNumber('4.2.1', 'auto')).toBe('4.2.1');
    expect(normaliseNumber('1', 'auto')).toBe('1');
  });

  it('paren-wrapped letter → normalised as sub-token', () => {
    expect(normaliseNumber('(b)', 'auto')).toBe('(b)');
    expect(normaliseNumber('(a)', 'auto')).toBe('(a)');
  });

  it('paren-wrapped roman → lowercase', () => {
    expect(normaliseNumber('(IV)', 'auto')).toBe('(iv)');
    expect(normaliseNumber('(III)', 'auto')).toBe('(iii)');
  });
});

describe('normaliseNumber — recital context', () => {
  it('uppercase letter, no parens', () => {
    expect(normaliseNumber('(A)', 'recital')).toBe('A');
    expect(normaliseNumber('B.', 'recital')).toBe('B');
  });
});

describe('normaliseNumber — part context', () => {
  it('uppercase roman', () => {
    expect(normaliseNumber('II', 'part')).toBe('II');
    expect(normaliseNumber('iii', 'part')).toBe('III');
  });
});

describe('normaliseNumber — schedule-or-annex context', () => {
  it('letters uppercase, numbers unchanged', () => {
    expect(normaliseNumber('A', 'schedule-or-annex')).toBe('A');
    expect(normaliseNumber('a', 'schedule-or-annex')).toBe('A');
    expect(normaliseNumber('2', 'schedule-or-annex')).toBe('2');
  });
});

describe('normaliseSubToken', () => {
  it('produces lowercase paren form', () => {
    expect(normaliseSubToken('b')).toBe('(b)');
    expect(normaliseSubToken('(b)')).toBe('(b)');
    expect(normaliseSubToken('(IV)')).toBe('(iv)');
  });
});

describe('stripParens', () => {
  it('strips leading parens and trailing dots', () => {
    expect(stripParens('(b)')).toBe('b');
    expect(stripParens('4.2.')).toBe('4.2');
  });
});
