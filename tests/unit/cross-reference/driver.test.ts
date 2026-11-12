/**
 * Feature 02 driver integration tests (spec §12.3 — cross-cutting acceptance).
 *
 * Tests: featureId separation (A/B never merged); Feature B rows have no fix action;
 * re-run clears both prefixes; default labels never emit "Section"/"Exhibit".
 *
 * ctx=null → ADR-0014 degraded mode (no field detection), all refs treated as static.
 */
import { describe, it, expect } from 'vitest';
import { analyseCrossReference } from '@features/cross-reference/driver.js';
import { clearFindings, allFindings } from '@shared/findings/index.js';
import { UK_PROFILE } from '@shared/config/index.js';
import { para } from '../../fixtures/cross-reference/paras.js';
import { beforeEach } from 'vitest';

beforeEach(() => clearFindings());

describe('Feature 02 driver — §12.3 cross-cutting', () => {
  it('Feature A and Feature B findings use separate featureId discriminators', async () => {
    const paras = [
      // Static-numbered clause (Feature B detects it)
      para(0, '4.2 The Buyer shall pay…'),
      // Prose reference that resolves (Feature A OK + static badge)
      para(1, 'As per Clause 1 herein.', { listString: '2', listLevel: 0 }),
    ];
    const result = await analyseCrossReference(paras, UK_PROFILE, null);
    const xrefFindings = result.xrefFindings;
    const staticFindings = result.staticNumFindings;

    // No finding from xref appears with featureId='staticNumbering'
    for (const f of xrefFindings) {
      expect(f.featureId).toBe('xref');
      expect(f.category).toMatch(/^cross-reference\//);
    }
    // No finding from staticNumbering appears with featureId='xref'
    for (const f of staticFindings) {
      expect(f.featureId).toBe('staticNumbering');
      expect(f.category).toMatch(/^static-numbering\//);
    }
  });

  it('Feature B rows have no suggestedAction (no fix offered — ADR-0006)', async () => {
    const paras = [
      para(0, '4.2 The Buyer shall pay…'),
    ];
    const result = await analyseCrossReference(paras, UK_PROFILE, null);
    for (const f of result.staticNumFindings) {
      expect(f.suggestedAction).toBeUndefined();
    }
  });

  it('clearThenEmit removes both prefixes and emits fresh results', async () => {
    const { clearThenEmit } = await import('@features/cross-reference/driver.js');
    const { addFinding } = await import('@shared/findings/index.js');

    addFinding({
      id: 'cross-reference/stale-1',
      category: 'cross-reference/not-found',
      featureId: 'xref',
      severity: 'warning',
      title: 'Old finding',
      location: { label: 'p0' },
    });
    addFinding({
      id: 'static-numbering/stale-1',
      category: 'static-numbering/manually-typed',
      featureId: 'staticNumbering',
      severity: 'info',
      title: 'Old static finding',
      location: { label: 'p0' },
    });

    await clearThenEmit([], UK_PROFILE, null);

    const remaining = allFindings.value;
    expect(remaining.some(f => f.id === 'cross-reference/stale-1')).toBe(false);
    expect(remaining.some(f => f.id === 'static-numbering/stale-1')).toBe(false);
  });

  it('default replacement labels never use "Section" or "Exhibit" (spec §7)', async () => {
    const paras = [
      para(0, 'Schedule 1', { styleName: 'Heading 1' }),
      para(1, 'Schedule 2', { styleName: 'Heading 1' }),
    ];
    const result = await analyseCrossReference(paras, UK_PROFILE, null);
    for (const node of result.index.ordered) {
      expect(node.displayLabel).not.toContain('Section');
      expect(node.displayLabel).not.toContain('Exhibit');
    }
  });

  it('statically-numbered unit is in the index and does not cause false Not Found', async () => {
    // Para 0: static "4.2" clause; Para 1: prose reference to "Clause 4.2"
    const paras = [
      para(0, '4.2 The Buyer shall pay the Purchase Price…'),
      para(1, 'As set out in Clause 4.2 above, payment is due…'),
    ];
    const result = await analyseCrossReference(paras, UK_PROFILE, null);

    // The index should have clause:4.2 (static fallback)
    expect(result.index.has('clause:4.2')).toBe(true);
    expect(result.index.resolve('clause:4.2')[0]?.isStaticNumbered).toBe(true);

    // The xref findings should NOT have a "not-found" for "Clause 4.2"
    const notFoundForClause42 = result.xrefFindings.filter(
      f => f.category === 'cross-reference/not-found' && f.title.includes('4.2'),
    );
    expect(notFoundForClause42).toHaveLength(0);
  });
});

describe('Feature 02 driver — config', () => {
  it('both SG and UK profiles carry xref and staticNumbering config', async () => {
    const { SG_PROFILE } = await import('@shared/config/index.js');
    expect(SG_PROFILE.xref.units.primaryBody).toBe('Clause');
    expect(SG_PROFILE.xref.units.annexurePrimary).toBe('Schedule');
    expect(SG_PROFILE.xref.section.isTopLevelDivision).toBe(false);
    expect(SG_PROFILE.staticNumbering.maxOoxmlProbes).toBe(200);
    expect(UK_PROFILE.xref.units.primaryBody).toBe('Clause');
    expect(UK_PROFILE.staticNumbering.headingSeverity).toBe('medium');
    expect(UK_PROFILE.staticNumbering.bodySeverity).toBe('low');
  });
});
