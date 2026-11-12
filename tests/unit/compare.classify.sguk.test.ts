/**
 * Feature 08 — Compare · SG/UK forcing-rule tests (Task B2).
 *
 * Covers T6 (en-US→en-GB spelling variant), T8 (table-cell value forced
 * numerical), and T9 (signature block forced substantive). Fixtures are
 * hand-built DiffOps.
 */

import { describe, it, expect } from 'vitest';
import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import type { CompareConfig } from '@shared/config/index.js';
import type { DiffOp, ParaRecord, WordDiff } from '../../src/features/compare/model.js';
import { classifyOp } from '../../src/features/compare/classify/classifier.js';
import { applySgUkRules } from '../../src/features/compare/classify/sguk-rules.js';
import { detectSpellingVariant } from '../../src/features/compare/classify/spelling-variant.js';
import { applySignatureBlockRule } from '../../src/features/compare/classify/signature-block.js';

const CONFIG: CompareConfig = getCompareConfig(SG_PROFILE);

function para(overrides: Partial<ParaRecord> = {}): ParaRecord {
  return {
    index: 0,
    partId: 'document',
    rawXml: '<w:p/>',
    normText: '',
    normTextAggressive: '',
    numberingPrefix: '',
    hash: '0',
    styleId: null,
    isTableCell: false,
    isSignatureBlock: false,
    isHeaderFooter: false,
    isFootnote: false,
    ...overrides,
  };
}

function replaceOp(wordDiffs: WordDiff[], overrides: Partial<DiffOp> = {}): DiffOp {
  return {
    id: 'op-1',
    type: 'replace',
    origRange: { start: 0, end: 0 },
    revRange: { start: 0, end: 0 },
    partId: 'document',
    wordDiffs,
    classification: 'substantiveText',
    classes: ['substantiveText'],
    author: 'Test',
    date: '2026-06-08T00:00:00Z',
    ...overrides,
  };
}

/** Classify (B1) then force (B2) — the full per-op pipeline. */
function classifyAndForce(
  op: DiffOp,
  orig: ParaRecord[],
  rev: ParaRecord[],
): DiffOp {
  return applySgUkRules(classifyOp(op, orig, rev, CONFIG), orig, rev, CONFIG);
}

// ---------------------------------------------------------------------------
// T6 — en-US → en-GB spelling variant
// ---------------------------------------------------------------------------

describe('SG/UK — T6 en-US→en-GB spelling variant', () => {
  it('detectSpellingVariant recognises organize → organise', () => {
    const op = replaceOp([
      { op: 'del', text: 'organize' },
      { op: 'ins', text: 'organise' },
    ]);
    const variant = detectSpellingVariant(op, CONFIG);
    expect(variant).not.toBeNull();
    expect(variant?.usSpelling).toBe('organize');
    expect(variant?.gbSpelling).toBe('organise');
  });

  it('organize → organise is forced substantiveText with a spellingVariant flag', () => {
    const op = replaceOp([
      { op: 'del', text: 'organize' },
      { op: 'ins', text: 'organise' },
    ]);
    const out = classifyAndForce(op, [para()], [para()]);
    expect(out.classification).toBe('substantiveText');
    expect(out.spellingVariant?.usSpelling).toBe('organize');
    expect(out.spellingVariant?.gbSpelling).toBe('organise');
    // Never filed as filterable noise.
    expect(out.classes).not.toContain('formatting');
    expect(out.classes).not.toContain('whitespacePunct');
  });

  it('preserves original casing (Color → Colour)', () => {
    const op = replaceOp([
      { op: 'del', text: 'Color' },
      { op: 'ins', text: 'Colour' },
    ]);
    const variant = detectSpellingVariant(op, CONFIG);
    expect(variant?.usSpelling).toBe('Color');
    expect(variant?.gbSpelling).toBe('Colour');
  });

  it('returns null when the ins side is not the mapped en-GB form', () => {
    const op = replaceOp([
      { op: 'del', text: 'organize' },
      { op: 'ins', text: 'arrange' },
    ]);
    expect(detectSpellingVariant(op, CONFIG)).toBeNull();
  });

  it('a spelling-variant change can never be classified formatting', () => {
    // Even if the surrounding records would look style-only, the variant rule wins.
    const op = replaceOp([
      { op: 'del', text: 'colour' === 'colour' ? 'organize' : '' },
      { op: 'ins', text: 'organise' },
    ]);
    const out = classifyAndForce(op, [para()], [para()]);
    expect(out.classification).not.toBe('formatting');
  });
});

// ---------------------------------------------------------------------------
// T8 — table-cell value forced numerical
// ---------------------------------------------------------------------------

describe('SG/UK — T8 table-cell value forced numerical', () => {
  it('a cell amount change is forced numericalValue, never formatting', () => {
    const op = replaceOp(
      [
        { op: 'del', text: 'S$1,250,000', atomic: 'amount' },
        { op: 'ins', text: 'S$1,500,000', atomic: 'amount' },
      ],
      { partId: 'tableCell:r1c2' },
    );
    const orig = [para({ isTableCell: true, partId: 'tableCell:r1c2' })];
    const rev = [para({ isTableCell: true, partId: 'tableCell:r1c2' })];
    const out = classifyAndForce(op, orig, rev);
    expect(out.classification).toBe('numericalValue');
    expect(out.classes).not.toContain('formatting');
    expect(out.classes).not.toContain('whitespacePunct');
  });

  it('a non-numeric cell text change is forced substantiveText', () => {
    const op = replaceOp([
      { op: 'del', text: 'Borrower' },
      { op: 'ins', text: 'Guarantor' },
    ]);
    const orig = [para({ isTableCell: true })];
    const rev = [para({ isTableCell: true })];
    const out = classifyAndForce(op, orig, rev);
    expect(['numericalValue', 'substantiveText']).toContain(out.classification);
    expect(out.classes).not.toContain('formatting');
  });

  it('does not demote a higher-priority structural cell op below the floor', () => {
    const op = replaceOp([{ op: 'ins', text: 'new row text' }], {
      type: 'insert',
      origRange: null,
      revRange: { start: 0, end: 0 },
    });
    const rev = [para({ isTableCell: true, normText: 'new row text' })];
    const out = classifyAndForce(op, [], rev);
    // structural (1) outranks the numerical/substantive floor (3/6) — kept.
    expect(out.classification).toBe('structural');
  });
});

// ---------------------------------------------------------------------------
// T9 — signature block forced substantive
// ---------------------------------------------------------------------------

describe('SG/UK — T9 signature block forced substantive', () => {
  it('a witness-line edit in a signature block is forced substantiveText', () => {
    const op = replaceOp([
      { op: 'del', text: 'Jane Doe' },
      { op: 'ins', text: 'John Roe' },
    ]);
    const orig = [para({ isSignatureBlock: true, normText: 'in the presence of Jane Doe' })];
    const rev = [para({ isSignatureBlock: true, normText: 'in the presence of John Roe' })];
    const out = classifyAndForce(op, orig, rev);
    expect(out.classification).toBe('substantiveText');
    expect(out.classes).not.toContain('formatting');
    expect(out.classes).not.toContain('whitespacePunct');
  });

  it('a formatting-only signature-block change is still forced substantive (never filterable)', () => {
    // No text delta would normally be formatting — but a signature block forces it.
    const op = replaceOp([{ op: 'eq', text: 'EXECUTED as a DEED' }]);
    const orig = [para({ isSignatureBlock: true, normText: 'EXECUTED as a DEED' })];
    const rev = [para({ isSignatureBlock: true, normText: 'EXECUTED as a DEED', styleId: 'Bold' })];
    const out = applySignatureBlockRule(classifyOp(op, orig, rev, CONFIG), orig, rev);
    expect(out.classification).toBe('substantiveText');
    expect(out.classes).not.toContain('formatting');
  });

  it('an added signature block is forced structural', () => {
    const op = replaceOp([{ op: 'ins', text: 'EXECUTED as a DEED' }], {
      type: 'insert',
      origRange: null,
      revRange: { start: 0, end: 0 },
    });
    const rev = [para({ isSignatureBlock: true, normText: 'EXECUTED as a DEED' })];
    const out = classifyAndForce(op, [], rev);
    expect(out.classification).toBe('structural');
    expect(out.classes).toContain('substantiveText');
  });

  it('clears any spellingVariant flag inside a signature block', () => {
    const op = replaceOp(
      [
        { op: 'del', text: 'organize' },
        { op: 'ins', text: 'organise' },
      ],
    );
    const orig = [para({ isSignatureBlock: true })];
    const rev = [para({ isSignatureBlock: true })];
    const out = classifyAndForce(op, orig, rev);
    // Signature-block rule runs after the spelling rule and clears the sub-flag.
    expect(out.spellingVariant).toBeUndefined();
    expect(out.classification).toBe('substantiveText');
  });
});
