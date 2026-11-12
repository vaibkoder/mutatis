/**
 * Tests for B1 — Clause-entry schema validation.
 *
 * Covers acceptance checks B-T1 (save with required fields enforced),
 * staleness warning, OTHER jurisdiction enforcement, shared-backend gate.
 *
 * No Office JS required — all pure validation logic.
 */

import { describe, it, expect } from 'vitest';
import {
  validateClauseMetadata,
  validateTitle,
  validateSelectionNonEmpty,
} from '@features/content/library/validation.js';
import type { ClauseMetadata } from '@features/content/library/types.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const VALID_META: ClauseMetadata = {
  jurisdiction: 'SG',
  practiceArea: 'Commercial',
  lastReviewed: '2025-06-01',
  approvedBy: 'J. Tan (Partner)',
  visibility: 'personal',
};

const NOW = new Date('2026-06-08T00:00:00Z');

// ---------------------------------------------------------------------------
// validateTitle
// ---------------------------------------------------------------------------

describe('validateTitle', () => {
  it('returns null for a non-empty title', () => {
    expect(validateTitle('Limitation of Liability')).toBeNull();
  });

  it('returns an error for an empty string', () => {
    const err = validateTitle('');
    expect(err).not.toBeNull();
    expect(err?.field).toBe('title');
  });

  it('returns an error for whitespace-only', () => {
    expect(validateTitle('   ')).not.toBeNull();
  });

  it('returns an error for null/undefined', () => {
    expect(validateTitle(null)).not.toBeNull();
    expect(validateTitle(undefined)).not.toBeNull();
  });
});

// ---------------------------------------------------------------------------
// validateSelectionNonEmpty
// ---------------------------------------------------------------------------

describe('validateSelectionNonEmpty', () => {
  it('returns true when html is non-empty', () => {
    expect(validateSelectionNonEmpty('<p>hello</p>', '')).toBe(true);
  });

  it('returns true when plainText is non-empty', () => {
    expect(validateSelectionNonEmpty('', 'hello')).toBe(true);
  });

  it('returns false when both are empty/whitespace', () => {
    expect(validateSelectionNonEmpty('', '')).toBe(false);
    expect(validateSelectionNonEmpty('  ', '  ')).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// validateClauseMetadata — required fields
// ---------------------------------------------------------------------------

describe('validateClauseMetadata — required fields (B-T1)', () => {
  it('passes when all required fields are present and valid', () => {
    const result = validateClauseMetadata(VALID_META, false, NOW);
    expect(result.valid).toBe(true);
    expect(result.errors).toHaveLength(0);
  });

  it('errors when jurisdiction is missing', () => {
    const meta = { ...VALID_META, jurisdiction: undefined } as unknown as Partial<ClauseMetadata>;
    const result = validateClauseMetadata(meta, false, NOW);
    expect(result.valid).toBe(false);
    expect(result.errors.some(e => e.field === 'jurisdiction')).toBe(true);
  });

  it('errors when jurisdiction is an invalid value', () => {
    const meta = { ...VALID_META, jurisdiction: 'US' } as unknown as Partial<ClauseMetadata>;
    const result = validateClauseMetadata(meta, false, NOW);
    expect(result.valid).toBe(false);
    expect(result.errors.some(e => e.field === 'jurisdiction')).toBe(true);
  });

  it('errors when jurisdiction === OTHER but jurisdictionOther is empty', () => {
    const meta: Partial<ClauseMetadata> = { ...VALID_META, jurisdiction: 'OTHER', jurisdictionOther: '' };
    const result = validateClauseMetadata(meta, false, NOW);
    expect(result.valid).toBe(false);
    expect(result.errors.some(e => e.field === 'jurisdictionOther')).toBe(true);
  });

  it('passes when jurisdiction === OTHER and jurisdictionOther is provided', () => {
    const meta: Partial<ClauseMetadata> = { ...VALID_META, jurisdiction: 'OTHER', jurisdictionOther: 'New York' };
    const result = validateClauseMetadata(meta, false, NOW);
    expect(result.valid).toBe(true);
  });

  it('errors when lastReviewed is missing', () => {
    const meta = { ...VALID_META, lastReviewed: '' };
    const result = validateClauseMetadata(meta, false, NOW);
    expect(result.valid).toBe(false);
    expect(result.errors.some(e => e.field === 'lastReviewed')).toBe(true);
  });

  it('errors when lastReviewed is in the future', () => {
    const meta = { ...VALID_META, lastReviewed: '2027-01-01' };
    const result = validateClauseMetadata(meta, false, NOW);
    expect(result.valid).toBe(false);
    expect(result.errors.some(e => e.field === 'lastReviewed')).toBe(true);
  });

  it('errors when lastReviewed is not a valid date', () => {
    const meta = { ...VALID_META, lastReviewed: 'not-a-date' };
    const result = validateClauseMetadata(meta, false, NOW);
    expect(result.valid).toBe(false);
    expect(result.errors.some(e => e.field === 'lastReviewed')).toBe(true);
  });

  it('accepts lastReviewed = today', () => {
    const meta = { ...VALID_META, lastReviewed: '2026-06-08' };
    const result = validateClauseMetadata(meta, false, NOW);
    expect(result.valid).toBe(true);
  });

  it('errors when approvedBy is empty', () => {
    const meta = { ...VALID_META, approvedBy: '' };
    const result = validateClauseMetadata(meta, false, NOW);
    expect(result.valid).toBe(false);
    expect(result.errors.some(e => e.field === 'approvedBy')).toBe(true);
  });

  it('errors when practiceArea is empty', () => {
    const meta = { ...VALID_META, practiceArea: '' };
    const result = validateClauseMetadata(meta, false, NOW);
    expect(result.valid).toBe(false);
    expect(result.errors.some(e => e.field === 'practiceArea')).toBe(true);
  });

  it('errors when visibility is missing', () => {
    const meta = { ...VALID_META, visibility: undefined } as unknown as Partial<ClauseMetadata>;
    const result = validateClauseMetadata(meta, false, NOW);
    expect(result.valid).toBe(false);
    expect(result.errors.some(e => e.field === 'visibility')).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// validateClauseMetadata — shared-backend gate (B-T1)
// ---------------------------------------------------------------------------

describe('validateClauseMetadata — shared backend gate', () => {
  it('errors when visibility=shared but backend not configured', () => {
    const meta = { ...VALID_META, visibility: 'shared' as const };
    const result = validateClauseMetadata(meta, false, NOW);
    expect(result.valid).toBe(false);
    expect(result.errors.some(e => e.field === 'visibility')).toBe(true);
  });

  it('passes when visibility=shared AND backend is configured', () => {
    const meta = { ...VALID_META, visibility: 'shared' as const };
    const result = validateClauseMetadata(meta, true, NOW);
    expect(result.valid).toBe(true);
  });

  it('passes with visibility=personal regardless of backend', () => {
    const result = validateClauseMetadata(VALID_META, false, NOW);
    expect(result.valid).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// validateClauseMetadata — staleness warning (non-blocking)
// ---------------------------------------------------------------------------

describe('validateClauseMetadata — staleness warning', () => {
  it('warns when lastReviewed is older than 24 months but does not block save', () => {
    // 25 months before NOW (2026-06-08)
    const meta = { ...VALID_META, lastReviewed: '2024-05-01' };
    const result = validateClauseMetadata(meta, false, NOW);
    expect(result.valid).toBe(true);  // not blocked
    expect(result.warnings.some(w => w.field === 'lastReviewed')).toBe(true);
  });

  it('does not warn when lastReviewed is within 24 months', () => {
    const meta = { ...VALID_META, lastReviewed: '2025-01-01' };  // ~17 months before NOW
    const result = validateClauseMetadata(meta, false, NOW);
    expect(result.valid).toBe(true);
    expect(result.warnings.some(w => w.field === 'lastReviewed')).toBe(false);
  });

  it('respects a custom staleness window', () => {
    // 13 months before NOW → stale under 12-month window, fine under 24-month
    const meta = { ...VALID_META, lastReviewed: '2025-05-01' };
    const result12 = validateClauseMetadata(meta, false, NOW, 12);
    const result24 = validateClauseMetadata(meta, false, NOW, 24);
    expect(result12.warnings.some(w => w.field === 'lastReviewed')).toBe(true);
    expect(result24.warnings.some(w => w.field === 'lastReviewed')).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// EW jurisdiction passes (not just SG)
// ---------------------------------------------------------------------------

describe('validateClauseMetadata — EW jurisdiction', () => {
  it('passes with jurisdiction=EW', () => {
    const meta = { ...VALID_META, jurisdiction: 'EW' as const };
    const result = validateClauseMetadata(meta, false, NOW);
    expect(result.valid).toBe(true);
  });
});
