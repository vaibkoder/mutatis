/**
 * C4 — Validators and formatters tests.
 * C-T2: EFFECTIVE_DATE renders "06 June 2026" (day-first).
 * C-T3: 8-char UK_CH accepted; SG_UEN rejected under EW and vice versa.
 */
import { describe, it, expect } from 'vitest';
import {
  validateRegNumber,
  validateUkCh,
  validateSgUen,
  getEntityTypeOptions,
  validateFieldValue,
} from '@features/content/templates/validators.js';
import {
  formatDateDayFirst,
  validateDate,
  formatValueForInsert,
  getRegNumberPlaceholder,
  getRegNumberLabel,
} from '@features/content/templates/format.js';

describe('C4 — formatDateDayFirst (C-T2)', () => {
  it('formats ISO 8601 date as DD Month YYYY', () => {
    expect(formatDateDayFirst('2026-06-06')).toBe('06 June 2026');
    expect(formatDateDayFirst('2026-01-01')).toBe('01 January 2026');
    expect(formatDateDayFirst('2026-12-31')).toBe('31 December 2026');
  });

  it('formats day-first D/M/YYYY as DD Month YYYY', () => {
    expect(formatDateDayFirst('6/6/2026')).toBe('06 June 2026');
    expect(formatDateDayFirst('06/06/2026')).toBe('06 June 2026');
  });

  it('formats from a Date object', () => {
    expect(formatDateDayFirst(new Date(2026, 5, 6))).toBe('06 June 2026');
  });

  it('returns null for invalid input', () => {
    expect(formatDateDayFirst('not-a-date')).toBeNull();
    expect(formatDateDayFirst('')).toBeNull();
  });

  it('never produces US month-first format', () => {
    const result = formatDateDayFirst('2026-06-06');
    // Must NOT be "06/06/2026" in MM/DD order
    expect(result).not.toMatch(/^06\s*\/\s*06/);
    // Must be day-first "06 June 2026"
    expect(result).toBe('06 June 2026');
  });
});

describe('C4 — formatValueForInsert', () => {
  it('formats date type as DD Month YYYY', () => {
    expect(formatValueForInsert('date', '2026-06-06')).toBe('06 June 2026');
  });

  it('formats boolean as Yes/No', () => {
    expect(formatValueForInsert('boolean', true)).toBe('Yes');
    expect(formatValueForInsert('boolean', false)).toBe('No');
    expect(formatValueForInsert('boolean', 'true')).toBe('Yes');
  });

  it('returns trimmed string for text type', () => {
    expect(formatValueForInsert('text', '  Acme Ltd  ')).toBe('Acme Ltd');
  });
});

describe('C4 — validateUkCh', () => {
  it('accepts 8-char alphanumeric UK_CH numbers', () => {
    expect(validateUkCh('00012345').valid).toBe(true);
    expect(validateUkCh('12345678').valid).toBe(true);
    expect(validateUkCh('SC012345').valid).toBe(true);
    expect(validateUkCh('NI012345').valid).toBe(true);
    expect(validateUkCh('OC012345').valid).toBe(true);
  });

  it('rejects numbers shorter or longer than 8 chars', () => {
    expect(validateUkCh('1234567').valid).toBe(false);   // 7 chars
    expect(validateUkCh('123456789').valid).toBe(false); // 9 chars
    expect(validateUkCh('').valid).toBe(false);
  });

  it('is case-insensitive (normalises to uppercase)', () => {
    expect(validateUkCh('sc012345').valid).toBe(true);
  });
});

describe('C4 — validateSgUen', () => {
  it('accepts valid SG UEN formats', () => {
    // Business UEN: 9 digits + checksum letter
    expect(validateSgUen('201912345A').valid).toBe(true);
    // Local company: 4-digit year + 5 digits + letter
    expect(validateSgUen('19990012345K').valid).toBe(false); // too long for nnnnnnnnX
    // Ensure pattern-based (read from config).
  });

  it('rejects UK CH numbers as SG UEN', () => {
    expect(validateSgUen('00012345').valid).toBe(false);
  });
});

describe('C4 — validateRegNumber jurisdiction gating (C-T3)', () => {
  it('accepts UK_CH under EW profile', () => {
    const r = validateRegNumber('00012345', 'UK_CH', 'EW');
    expect(r.valid).toBe(true);
  });

  it('rejects UK_CH under SG profile (C-T3)', () => {
    const r = validateRegNumber('00012345', 'UK_CH', 'SG');
    expect(r.valid).toBe(false);
    expect(r.error).toMatch(/Singapore/);
  });

  it('accepts SG_UEN under SG profile', () => {
    // 201912345A matches /^\d{9}[A-Z]$/
    const r = validateRegNumber('201912345A', 'SG_UEN', 'SG');
    expect(r.valid).toBe(true);
  });

  it('rejects SG_UEN under EW profile (C-T3)', () => {
    const r = validateRegNumber('201912345A', 'SG_UEN', 'EW');
    expect(r.valid).toBe(false);
    expect(r.error).toMatch(/England/);
  });
});

describe('C4 — getEntityTypeOptions', () => {
  it('returns EW entity types (no Inc./LLC)', () => {
    const opts = getEntityTypeOptions('EW');
    expect(opts).toContain('Limited');
    expect(opts).toContain('Ltd');
    expect(opts).toContain('plc');
    expect(opts).not.toContain('Inc.');
    expect(opts).not.toContain('LLC');
    expect(opts).not.toContain('EIN');
  });

  it('returns SG entity types (no Inc./LLC)', () => {
    const opts = getEntityTypeOptions('SG');
    expect(opts).toContain('Pte. Ltd.');
    expect(opts).not.toContain('Inc.');
    expect(opts).not.toContain('LLC');
    expect(opts).not.toContain('EIN');
  });

  it('SG options do not include EW-only terms', () => {
    const sg = getEntityTypeOptions('SG');
    expect(sg).not.toContain('plc');
  });

  it('EW options do not include SG-only terms', () => {
    const ew = getEntityTypeOptions('EW');
    expect(ew).not.toContain('Pte. Ltd.');
  });
});

describe('C4 — validateDate', () => {
  it('returns null for valid dates', () => {
    expect(validateDate('2026-06-06', 'Effective Date')).toBeNull();
    expect(validateDate('', 'Effective Date')).toBeNull(); // empty = required check elsewhere
  });

  it('returns error for invalid dates', () => {
    expect(validateDate('not-a-date', 'Effective Date')).toBeTruthy();
  });
});

describe('C4 — reg-number labels and placeholders', () => {
  it('UK_CH label mentions England & Wales', () => {
    expect(getRegNumberLabel('UK_CH')).toMatch(/England/);
  });

  it('SG_UEN label mentions Singapore', () => {
    expect(getRegNumberLabel('SG_UEN')).toMatch(/Singapore/);
  });

  it('UK_CH placeholder shows alphanumeric example', () => {
    expect(getRegNumberPlaceholder('UK_CH')).toMatch(/\d{6}/);
  });
});

describe('C4 — validateFieldValue', () => {
  it('returns error when required field is empty', () => {
    expect(validateFieldValue('', { required: true }, 'Client Name')).toBeTruthy();
  });

  it('returns null when required field has value', () => {
    expect(validateFieldValue('Acme', { required: true }, 'Client Name')).toBeNull();
  });

  it('respects maxLen constraint', () => {
    expect(validateFieldValue('A'.repeat(121), { maxLen: 120 }, 'Name')).toBeTruthy();
    expect(validateFieldValue('A'.repeat(120), { maxLen: 120 }, 'Name')).toBeNull();
  });

  it('respects pattern constraint', () => {
    expect(validateFieldValue('abc', { pattern: '^[A-Z]+$' }, 'Code')).toBeTruthy();
    expect(validateFieldValue('ABC', { pattern: '^[A-Z]+$' }, 'Code')).toBeNull();
  });
});
