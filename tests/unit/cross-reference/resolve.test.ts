/**
 * Feature A — resolution tests (spec §5.1.3, T5 acceptance).
 * Tests A1-A5, A13 (broken REF classification), A14 (externalMaybe).
 */
import { describe, it, expect } from 'vitest';
import { resolveAll } from '@features/cross-reference/resolve.js';
import { buildSectionIndex } from '@shared/section-index/builder.js';
import { UK_PROFILE } from '@shared/config/index.js';
import { para } from '../../fixtures/cross-reference/paras.js';
import type { ReferenceMatch } from '@features/cross-reference/types.js';

function makeMatch(normalizedKey: string, expandedKeys?: string[]): ReferenceMatch {
  return {
    matchId: 'test-1',
    rawText: normalizedKey,
    anchor: { paragraphId: 'body-para-99', startOffset: 0, endOffset: 5 },
    paragraphId: 'body-para-99',
    unitType: 'clause',
    normalizedKey,
    expandedKeys: expandedKeys ?? [normalizedKey],
    classification: 'notFound',
    candidateTargets: [],
    isFieldBacked: false,
    staticBadge: true,
    severity: 'warning',
  };
}

describe('Feature A resolution', () => {
  it('A1: existing clause → OK', () => {
    const paras = [para(0, 'Payment', { listString: '4.2', listLevel: 1 })];
    const index = buildSectionIndex(paras, UK_PROFILE);
    const match = makeMatch('clause:4.2');
    resolveAll([match], index, UK_PROFILE);
    expect(match.classification).toBe('ok');
  });

  it('A2: deleted clause → Not Found', () => {
    const index = buildSectionIndex([], UK_PROFILE);
    const match = makeMatch('clause:4.2');
    resolveAll([match], index, UK_PROFILE);
    expect(match.classification).toBe('notFound');
  });

  it('A3: clause 5 absent but 5.1 and 5.2 present → Ambiguous (implied-parent, spec §5.1.3 step 4)', () => {
    // Realistic "no standalone 5" structure: sub-clauses 5.1/5.2 exist as list
    // items but there is no paragraph numbered exactly "5". A cite of "Clause 5"
    // is under-specified → Ambiguous, resolved via prefix lookup in byKey.
    const index = buildSectionIndex([
      para(0, 'The Seller warrants…', { listString: '5.1', listLevel: 1 }),
      para(1, 'The Seller further warrants…', { listString: '5.2', listLevel: 1 }),
    ], UK_PROFILE);
    // Confirm the fixture genuinely lacks a standalone clause:5.
    expect(index.has('clause:5')).toBe(false);
    expect(index.has('clause:5.1')).toBe(true);

    const match = makeMatch('clause:5');
    resolveAll([match], index, UK_PROFILE);
    expect(match.classification).toBe('ambiguous');
    expect(match.candidateTargets.length).toBeGreaterThan(0);
  });

  it('A3b: lettered prefix-ambiguity is also handled by resolve (clause:5(a)/5(b) without 5)', () => {
    // The builder cannot naturally emit 5(a)/5(b) without a level-0 "5", so we
    // exercise resolve.ts's lettered prefix path against a hand-built index whose
    // byKey carries 5(a)/5(b) but no standalone 5. This guards the "clause:5("
    // branch of getPrefixChildren independently of the builder.
    const mkNode = (key: string) => ({
      key, kind: 'clause' as const, rawNumber: key, displayLabel: key,
      level: 1, paraId: key, isStaticNumbered: false, bookmarkNames: [],
      getRange: () => { throw new Error('unused'); },
    });
    const byKey = new Map([
      ['clause:5(a)', [mkNode('clause:5(a)')]],
      ['clause:5(b)', [mkNode('clause:5(b)')]],
    ]);
    const index = {
      nodes: [], ordered: [], byParaId: new Map(), byKey,
      has: (k: string) => (byKey.get(k)?.length ?? 0) > 0,
      resolve: (k: string) => byKey.get(k) ?? [],
    };
    const match = makeMatch('clause:5');
    resolveAll([match], index, UK_PROFILE);
    expect(match.classification).toBe('ambiguous');
    expect(match.candidateTargets.length).toBe(2);
  });

  it('A4: paragraph 3 of Schedule 2 — compound key both halves OK', () => {
    const paras = [
      para(0, 'Schedule 2', { styleName: 'Heading 1' }),
      para(1, 'The Property', { listString: '1', listLevel: 0 }),
      para(2, 'The land', { listString: '3', listLevel: 0 }),
    ];
    const index = buildSectionIndex(paras, UK_PROFILE);
    const match = makeMatch('schedule:2/para:3');
    match.expandedKeys = ['schedule:2/para:3'];
    resolveAll([match], index, UK_PROFILE);
    expect(match.classification).toBe('ok');
  });

  it('A5: paragraph 3 of Schedule 2 but only 1-2 exist → Not Found', () => {
    const paras = [
      para(0, 'Schedule 2', { styleName: 'Heading 1' }),
      para(1, 'Item 1', { listString: '1', listLevel: 0 }),
      para(2, 'Item 2', { listString: '2', listLevel: 0 }),
    ];
    const index = buildSectionIndex(paras, UK_PROFILE);
    const match = makeMatch('schedule:2/para:3');
    match.expandedKeys = ['schedule:2/para:3'];
    resolveAll([match], index, UK_PROFILE);
    expect(match.classification).toBe('notFound');
  });

  it('A14: Schedule 9 not in doc → externalMaybe', () => {
    const paras = [
      para(0, 'Schedule 1', { styleName: 'Heading 1' }),
      para(1, 'Schedule 2', { styleName: 'Heading 1' }),
      para(2, 'Schedule 3', { styleName: 'Heading 1' }),
    ];
    const index = buildSectionIndex(paras, UK_PROFILE);
    const match = makeMatch('schedule:9');
    match.unitType = 'schedule';
    resolveAll([match], index, UK_PROFILE);
    expect(match.classification).toBe('externalMaybe');
    expect(match.severity).toBe('info');
  });

  it('ambiguous roman/letter (i): a cite resolves OK regardless of reading (literal-token key)', () => {
    // The builder marks (i) ambiguousRomanLetter but stores the literal token, so
    // the cited form "Clause 5(i)" and the stored key collapse identically — no
    // false Not Found under either the roman or letter interpretation (spec §4.4).
    const index = buildSectionIndex([
      para(0, 'Warranties', { listString: '5', listLevel: 0 }),
      para(1, 'first sub', { listString: '(i)', listLevel: 1 }),
    ], UK_PROFILE);
    expect(index.resolve('clause:5(i)')[0]?.ambiguousRomanLetter).toBe(true);
    const match = makeMatch('clause:5(i)');
    resolveAll([match], index, UK_PROFILE);
    expect(match.classification).toBe('ok');
  });

  it('A16: nested "Clause 4.2(b)(i)" resolves OK against the built index', () => {
    const paras = [
      para(0, 'Payment', { listString: '4', listLevel: 0 }),
      para(1, 'Sub', { listString: '4.2', listLevel: 1 }),
      para(2, 'Letter', { listString: '(b)', listLevel: 2 }),
      para(3, 'Roman', { listString: '(i)', listLevel: 3 }),
    ];
    const index = buildSectionIndex(paras, UK_PROFILE);
    expect(index.has('clause:4.2(b)(i)')).toBe(true);
    const match = makeMatch('clause:4.2(b)(i)');
    resolveAll([match], index, UK_PROFILE);
    expect(match.classification).toBe('ok');
  });

  it('replacement label uses configured "Schedule" not "Exhibit" (spec §7 hard rule)', () => {
    // The Section Index displayLabel should use configured term, not alias.
    const paras = [para(0, 'Schedule 1', { styleName: 'Heading 1' })];
    const index = buildSectionIndex(paras, UK_PROFILE);
    const node = index.resolve('schedule:1')[0];
    expect(node?.displayLabel).toBe('Schedule 1');
    expect(node?.displayLabel).not.toContain('Exhibit');
    expect(node?.displayLabel).not.toContain('Section');
  });
});
