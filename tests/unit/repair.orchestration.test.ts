// tests/unit/repair.orchestration.test.ts
//
// Orchestration tests for scan(), apply(), and the guards.
// All detector and apply modules are replaced with vi.mock stubs — no Office JS
// or OOXML service calls are made here.

import { describe, it, expect, vi, beforeEach } from 'vitest';

// ── Mock: shared/office (wordRun) ────────────────────────────────────────────
// vi.mock is hoisted; the factory must not reference variables declared below.
vi.mock('@shared/office/index.js', () => ({
  wordRun: vi.fn(async (fn: (ctx: unknown) => Promise<unknown>) => fn({})),
}));

// ── Mock: detect modules ─────────────────────────────────────────────────────
vi.mock('@features/repair/detect/metadata.js', () => ({
  detectMetadata: vi.fn(async () => []),
}));
vi.mock('@features/repair/detect/structure.js', () => ({
  detectStructure: vi.fn(async () => []),
  detectHeuristics: vi.fn(async () => ({
    detectedLineSpacing: null,
    detectedMarginPt: null,
    detectedFontSize: null,
    detectedFontName: null,
  })),
}));
vi.mock('@features/repair/detect/tracked.js', () => ({
  detectTracked: vi.fn(async () => null),
}));
vi.mock('@features/repair/detect/ooxml-scan.js', () => ({
  detectOoxml: vi.fn(async () => ({
    findings: [],
    metrics: { fileSizeBeforeBytes: 1024, rsidCountBefore: 0 },
    bundle: { _kind: 'stub' },
  })),
}));

// ── Mock: apply modules ───────────────────────────────────────────────────────
vi.mock('@features/repair/apply/metadata.js', () => ({
  applyMetadata: vi.fn(async (_ctx: unknown, findings: unknown[]) =>
    (findings as Array<{ id: string; label: string; action: string }>).map(f => ({ ...f, result: 'done' })),
  ),
}));
vi.mock('@features/repair/apply/standardise.js', () => ({
  applyStandardisation: vi.fn(async (_ctx: unknown, findings: unknown[]) =>
    (findings as Array<{ id: string; label: string; action: string }>).map(f => ({ ...f, result: 'done' })),
  ),
}));
vi.mock('@features/repair/apply/spacing.js', () => ({
  applySpacing: vi.fn(async (_ctx: unknown, findings: unknown[]) =>
    (findings as Array<{ id: string; label: string; action: string }>).map(f => ({ ...f, result: 'done' })),
  ),
}));
vi.mock('@features/repair/apply/ooxml-apply.js', () => ({
  applyOoxmlSurgery: vi.fn(async () => ({
    findings: [],
    metricsAfter: { fileSizeAfterBytes: 900, rsidCountAfter: 0 },
  })),
}));
vi.mock('@features/repair/apply/tracked.js', () => ({
  applyTracked: vi.fn(async (_ctx: unknown, findings: unknown[]) =>
    (findings as Array<{ id: string; label: string; action: string }>).map(f => ({ ...f, result: 'done' })),
  ),
}));

// ── Import under test (must come after vi.mock calls) ────────────────────────
import { scan } from '@features/repair/scan.js';
import { apply } from '@features/repair/apply.js';
import {
  getNormalisationWarnings,
  getCommercialOnCourtWarning,
} from '@features/repair/guards.js';
import { detectMetadata } from '@features/repair/detect/metadata.js';
import { detectStructure } from '@features/repair/detect/structure.js';
import { detectTracked } from '@features/repair/detect/tracked.js';
import { detectOoxml } from '@features/repair/detect/ooxml-scan.js';
import { applyMetadata } from '@features/repair/apply/metadata.js';
import { applyStandardisation } from '@features/repair/apply/standardise.js';
import { applySpacing } from '@features/repair/apply/spacing.js';
import { applyOoxmlSurgery } from '@features/repair/apply/ooxml-apply.js';
import { applyTracked } from '@features/repair/apply/tracked.js';
import type { RepairFinding } from '@features/repair/types.js';
import type { StyleProfile } from '@features/repair/profiles.js';
import type { DocumentHeuristics } from '@features/repair/guards.js';

// ── Helpers ───────────────────────────────────────────────────────────────────

function makeFinding(id: string, action: string, overrides: Partial<RepairFinding> = {}): RepairFinding {
  return {
    id,
    category: 'metadata',
    severity: 'confidentiality',
    label: `Finding ${id}`,
    destructive: false,
    requiresConfirm: false,
    selectedByDefault: true,
    action,
    result: null,
    ...overrides,
  };
}

/** Minimal StyleProfile stub — only the fields the orchestrators access. */
const stubProfile: StyleProfile = {
  id: 'sg-court-filing',
  name: 'SG Court Filing (PD 2021)',
  jurisdiction: 'SG',
  kind: 'court',
  locked: true,
  page: {
    margins: { top: 99.21, bottom: 99.21, left: 99.21, right: 99.21 },
    size: 'A4',
    orientation: 'portrait',
  },
  body: {
    font: { name: 'Times New Roman', size: 12, color: '#000000', bold: false, italic: false },
    spacing: { lineSpacingRule: 'multiple', lineSpacing: 2.0, spaceBefore: 0, spaceAfter: 0 },
    indent: { firstLine: 0, left: 0, right: 0 },
  },
  styles: {},
  rules: { normalisationOptIn: true, protectOrientationSectionBreaks: true },
};

const commercialProfile: StyleProfile = {
  ...stubProfile,
  id: 'firm-commercial',
  name: 'Firm Commercial',
  jurisdiction: 'GENERIC',
  kind: 'commercial',
};

const TIMESTAMP = '2026-06-08T10:00:00Z';

// ── scan() tests ──────────────────────────────────────────────────────────────

describe('scan()', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    // Restore default stub return values after clearAllMocks().
    vi.mocked(detectMetadata).mockResolvedValue([]);
    vi.mocked(detectStructure).mockResolvedValue([]);
    vi.mocked(detectTracked).mockResolvedValue(null);
    vi.mocked(detectOoxml).mockResolvedValue({
      findings: [],
      metrics: { fileSizeBeforeBytes: 1024, rsidCountBefore: 0 },
      bundle: { _kind: 'stub' } as unknown as Awaited<ReturnType<typeof detectOoxml>>['bundle'],
    });
  });

  it('returns a RepairReport with phase:"scan" and all four finding categories', async () => {
    const metaFinding = makeFinding('metadata.builtinProperties', 'clearBuiltinProps', { category: 'metadata' });
    const structFinding = makeFinding('styles.directFormatting', 'clearDirectFormatting', { category: 'styles' });
    const trackedFinding = makeFinding('tracked.changes', 'acceptAll', { category: 'tracked' });
    const ooxmlFinding = makeFinding('rsid.bloat', 'stripRsid', { category: 'rsid' });

    vi.mocked(detectMetadata).mockResolvedValue([metaFinding]);
    vi.mocked(detectStructure).mockResolvedValue([structFinding]);
    vi.mocked(detectTracked).mockResolvedValue(trackedFinding);
    vi.mocked(detectOoxml).mockResolvedValue({
      findings: [ooxmlFinding],
      metrics: { fileSizeBeforeBytes: 2048, rsidCountBefore: 42 },
      bundle: { _kind: 'stub' } as unknown as Awaited<ReturnType<typeof detectOoxml>>['bundle'],
    });

    const state = await scan('Test.docx', stubProfile, TIMESTAMP);
    const { report } = state;

    expect(report.phase).toBe('scan');
    expect(report.documentName).toBe('Test.docx');
    expect(report.generatedAt).toBe(TIMESTAMP);
    expect(report.profile.id).toBe('sg-court-filing');

    const ids = report.findings.map(f => f.id);
    expect(ids).toContain('metadata.builtinProperties');
    expect(ids).toContain('styles.directFormatting');
    expect(ids).toContain('tracked.changes');
    expect(ids).toContain('rsid.bloat');

    expect(report.metrics.fileSizeBeforeBytes).toBe(2048);
    expect(report.metrics.rsidCountBefore).toBe(42);
    expect(report.metrics.fileSizeAfterBytes).toBeNull();
    expect(report.metrics.rsidCountAfter).toBeNull();
    expect(report.errors).toHaveLength(0);
  });

  it('with empty detectors still returns a valid report with empty findings', async () => {
    // All mocks return empty (defaults already set in beforeEach).
    const state = await scan('Empty.docx', stubProfile, TIMESTAMP);

    expect(state.report.phase).toBe('scan');
    expect(state.report.findings).toHaveLength(0);
    expect(state.report.errors).toHaveLength(0);
    expect(state.bundle).toBeDefined();
  });

  it('invokes detectMetadata, detectStructure, detectTracked inside wordRun, and detectOoxml separately', async () => {
    await scan('Doc.docx', stubProfile, TIMESTAMP);

    // All typed detectors must have been called.
    expect(detectMetadata).toHaveBeenCalledTimes(1);
    expect(detectStructure).toHaveBeenCalledTimes(1);
    expect(detectTracked).toHaveBeenCalledTimes(1);
    // OOXML scanner must have been called.
    expect(detectOoxml).toHaveBeenCalledTimes(1);
  });

  it('includes the tracked finding when detectTracked returns non-null', async () => {
    const trackedFinding = makeFinding('tracked.changes', 'acceptAll', { category: 'tracked' });
    vi.mocked(detectTracked).mockResolvedValue(trackedFinding);

    const { report } = await scan('Doc.docx', stubProfile, TIMESTAMP);
    expect(report.findings.some(f => f.id === 'tracked.changes')).toBe(true);
  });

  it('excludes tracked when detectTracked returns null', async () => {
    vi.mocked(detectTracked).mockResolvedValue(null);

    const { report } = await scan('Doc.docx', stubProfile, TIMESTAMP);
    expect(report.findings.every(f => f.id !== 'tracked.changes')).toBe(true);
  });
});

// ── apply() tests ─────────────────────────────────────────────────────────────

describe('apply()', () => {
  const stubBundle = { _kind: 'stub' } as unknown as Parameters<typeof apply>[1];

  beforeEach(() => {
    vi.clearAllMocks();
    // Restore apply mock defaults.
    vi.mocked(applyMetadata).mockImplementation(async (_ctx, findings) =>
      (findings as RepairFinding[]).map(f => ({ ...f, result: 'done' as const })),
    );
    vi.mocked(applyStandardisation).mockImplementation(async (_ctx, findings) =>
      (findings as RepairFinding[]).map(f => ({ ...f, result: 'done' as const })),
    );
    vi.mocked(applySpacing).mockImplementation(async (_ctx, findings) =>
      (findings as RepairFinding[]).map(f => ({ ...f, result: 'done' as const })),
    );
    vi.mocked(applyTracked).mockImplementation(async (_ctx, findings) =>
      (findings as RepairFinding[]).map(f => ({ ...f, result: 'done' as const })),
    );
    vi.mocked(applyOoxmlSurgery).mockResolvedValue({
      findings: [],
      metricsAfter: { fileSizeAfterBytes: 900, rsidCountAfter: 0 },
    });
  });

  function makeScanReport(findings: RepairFinding[]) {
    return {
      documentName: 'Test.docx',
      generatedAt: TIMESTAMP,
      profile: { id: stubProfile.id, name: stubProfile.name },
      phase: 'scan' as const,
      findings,
      metrics: { fileSizeBeforeBytes: 2048, fileSizeAfterBytes: null, rsidCountBefore: 10, rsidCountAfter: null },
      errors: [],
    };
  }

  it('calls all apply functions with the correct filtered findings', async () => {
    const metaF  = makeFinding('metadata.builtinProperties', 'clearBuiltinProps', { category: 'metadata' });
    const stdF   = makeFinding('styles.directFormatting', 'clearDirectFormatting', { category: 'styles' });
    const spaceF = makeFinding('spacing.anomaly', 'normaliseSpacing', { category: 'spacing' });
    const trackF = makeFinding('tracked.changes', 'acceptAll', { category: 'tracked' });
    const rsidF  = makeFinding('rsid.bloat', 'stripRsid', { category: 'rsid', severity: 'info' });

    vi.mocked(applyOoxmlSurgery).mockResolvedValue({
      findings: [{ ...rsidF, result: 'done' }],
      metricsAfter: { fileSizeAfterBytes: 900, rsidCountAfter: 0 },
    });

    const scanReport = makeScanReport([metaF, stdF, spaceF, trackF, rsidF]);
    const selectedIds = [metaF.id, stdF.id, spaceF.id, trackF.id, rsidF.id];

    await apply(scanReport, stubBundle, stubProfile, {
      selectedFindingIds: selectedIds,
      confirmedSectionBreakIds: [],
    }, TIMESTAMP);

    // Metadata apply received the metadata finding.
    expect(applyMetadata).toHaveBeenCalledWith(
      expect.anything(),
      expect.arrayContaining([expect.objectContaining({ id: 'metadata.builtinProperties' })]),
    );

    // Standardisation apply received the styles finding.
    expect(applyStandardisation).toHaveBeenCalledWith(
      expect.anything(),
      expect.arrayContaining([expect.objectContaining({ id: 'styles.directFormatting' })]),
      stubProfile,
      expect.any(Object),
    );

    // Spacing apply received the spacing finding.
    expect(applySpacing).toHaveBeenCalledWith(
      expect.anything(),
      expect.arrayContaining([expect.objectContaining({ id: 'spacing.anomaly' })]),
      stubProfile,
      expect.any(Object),
    );

    // Tracked apply received the tracked finding.
    expect(applyTracked).toHaveBeenCalledWith(
      expect.anything(),
      expect.arrayContaining([expect.objectContaining({ id: 'tracked.changes' })]),
      expect.any(Object),
    );

    // OOXML surgery received the rsid finding plus both confirm channels.
    expect(applyOoxmlSurgery).toHaveBeenCalledWith(
      stubBundle,
      expect.arrayContaining([expect.objectContaining({ id: 'rsid.bloat' })]),
      [],
      [],
    );
  });

  it('produces a phase:"apply" report with result populated on each selected finding', async () => {
    const metaF  = makeFinding('metadata.builtinProperties', 'clearBuiltinProps', { category: 'metadata' });
    const rsidF  = makeFinding('rsid.bloat', 'stripRsid', { category: 'rsid', severity: 'info' });

    vi.mocked(applyOoxmlSurgery).mockResolvedValue({
      findings: [{ ...rsidF, result: 'done' }],
      metricsAfter: { fileSizeAfterBytes: 800, rsidCountAfter: 0 },
    });

    const scanReport = makeScanReport([metaF, rsidF]);
    const result = await apply(
      scanReport,
      stubBundle,
      stubProfile,
      { selectedFindingIds: [metaF.id, rsidF.id], confirmedSectionBreakIds: [] },
      TIMESTAMP,
    );

    expect(result.phase).toBe('apply');
    expect(result.metrics.fileSizeAfterBytes).toBe(800);
    expect(result.metrics.rsidCountAfter).toBe(0);

    const metaResult = result.findings.find(f => f.id === 'metadata.builtinProperties');
    expect(metaResult?.result).toBe('done');

    const rsidResult = result.findings.find(f => f.id === 'rsid.bloat');
    expect(rsidResult?.result).toBe('done');
  });

  it('not-selected findings get result:"skipped"', async () => {
    const selectedF = makeFinding('metadata.builtinProperties', 'clearBuiltinProps', { category: 'metadata' });
    const notSelectedF = makeFinding('spacing.anomaly', 'normaliseSpacing', { category: 'spacing' });

    const scanReport = makeScanReport([selectedF, notSelectedF]);
    const result = await apply(
      scanReport,
      stubBundle,
      stubProfile,
      { selectedFindingIds: [selectedF.id], confirmedSectionBreakIds: [] },
      TIMESTAMP,
    );

    const skipped = result.findings.find(f => f.id === 'spacing.anomaly');
    expect(skipped?.result).toBe('skipped');
  });

  it('error findings populate the report errors array', async () => {
    const metaF = makeFinding('metadata.builtinProperties', 'clearBuiltinProps', { category: 'metadata' });

    vi.mocked(applyMetadata).mockResolvedValue([{ ...metaF, result: 'error' }]);

    const scanReport = makeScanReport([metaF]);
    const result = await apply(
      scanReport,
      stubBundle,
      stubProfile,
      { selectedFindingIds: [metaF.id], confirmedSectionBreakIds: [] },
      TIMESTAMP,
    );

    expect(result.errors.length).toBeGreaterThan(0);
    expect(result.errors[0]?.findingId).toBe('metadata.builtinProperties');
    expect(result.errors[0]?.code).toBe('GeneralException');
  });
});

// ── getNormalisationWarnings() tests ─────────────────────────────────────────

describe('getNormalisationWarnings()', () => {
  const courtHeuristics: DocumentHeuristics = {
    detectedLineSpacing: 2.0,
    detectedMarginPt: 99.21,   // 35mm in pt — SG standard
    detectedFontSize: 12,
    detectedFontName: 'Times New Roman',
  };

  it('court profile + double-spaced TNR12 35mm → uncheckIds includes spacing.anomaly', () => {
    const { uncheckIds, warningMessage } = getNormalisationWarnings(stubProfile, courtHeuristics);

    expect(uncheckIds).toContain('spacing.anomaly');
    expect(warningMessage).not.toBeNull();
    expect(warningMessage).toMatch(/already match court formatting/i);
  });

  it('court profile + doc that does not match → empty uncheckIds and null warning', () => {
    const otherHeuristics: DocumentHeuristics = {
      detectedLineSpacing: 1.0,    // single-spaced
      detectedMarginPt: 36.0,      // narrow margins
      detectedFontSize: 11,
      detectedFontName: 'Calibri', // wrong font
    };
    const { uncheckIds, warningMessage } = getNormalisationWarnings(stubProfile, otherHeuristics);

    expect(uncheckIds).toHaveLength(0);
    expect(warningMessage).toBeNull();
  });

  it('non-court profile → always returns empty uncheckIds', () => {
    const { uncheckIds, warningMessage } = getNormalisationWarnings(commercialProfile, courtHeuristics);

    expect(uncheckIds).toHaveLength(0);
    expect(warningMessage).toBeNull();
  });

  it('court profile + lineSpacing exactly at threshold (1.9) → treated as court', () => {
    const borderHeuristics: DocumentHeuristics = {
      ...courtHeuristics,
      detectedLineSpacing: 1.9,
    };
    const { uncheckIds } = getNormalisationWarnings(stubProfile, borderHeuristics);
    expect(uncheckIds).toContain('spacing.anomaly');
  });

  it('null heuristic values → no false positive', () => {
    const nullHeuristics: DocumentHeuristics = {
      detectedLineSpacing: null,
      detectedMarginPt: null,
      detectedFontSize: null,
      detectedFontName: null,
    };
    const { uncheckIds, warningMessage } = getNormalisationWarnings(stubProfile, nullHeuristics);
    expect(uncheckIds).toHaveLength(0);
    expect(warningMessage).toBeNull();
  });
});

// ── getCommercialOnCourtWarning() tests ───────────────────────────────────────

describe('getCommercialOnCourtWarning()', () => {
  const courtHeuristics: DocumentHeuristics = {
    detectedLineSpacing: 2.0,
    detectedMarginPt: 99.21,
    detectedFontSize: 12,
    detectedFontName: 'Times New Roman',
  };

  it('commercial profile on court-looking doc → returns warning string', () => {
    const warning = getCommercialOnCourtWarning(commercialProfile, courtHeuristics);

    expect(warning).not.toBeNull();
    expect(typeof warning).toBe('string');
    expect(warning).toMatch(/court/i);
  });

  it('commercial profile on non-court doc → null', () => {
    const plainHeuristics: DocumentHeuristics = {
      detectedLineSpacing: 1.0,
      detectedMarginPt: 36.0,
      detectedFontSize: 11,
      detectedFontName: 'Calibri',
    };
    const warning = getCommercialOnCourtWarning(commercialProfile, plainHeuristics);
    expect(warning).toBeNull();
  });

  it('court profile on court-looking doc → null (guard only fires for commercial)', () => {
    const warning = getCommercialOnCourtWarning(stubProfile, courtHeuristics);
    expect(warning).toBeNull();
  });

  it('commercial profile + null heuristics → null (no false positive)', () => {
    const nullHeuristics: DocumentHeuristics = {
      detectedLineSpacing: null,
      detectedMarginPt: null,
      detectedFontSize: null,
      detectedFontName: null,
    };
    const warning = getCommercialOnCourtWarning(commercialProfile, nullHeuristics);
    expect(warning).toBeNull();
  });
});
