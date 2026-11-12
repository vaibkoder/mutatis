import { describe, it, expect, vi, beforeEach } from 'vitest';
import { PizZipBundle } from '@shared/ooxml/index.js';

// ---------------------------------------------------------------------------
// Mock ooxmlService before importing the module under test.
// vi.mock is hoisted so this executes before the imports below.
// ---------------------------------------------------------------------------

vi.mock('@shared/ooxml/index.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@shared/ooxml/index.js')>();
  return {
    ...actual,
    ooxmlService: {
      ...actual.ooxmlService,
      writeBack: vi.fn().mockResolvedValue(undefined),
    },
  };
});

import { ooxmlService } from '@shared/ooxml/index.js';
import { applyOoxmlSurgery } from '../../src/features/repair/apply/ooxml-apply.js';
import type { RepairFinding } from '../../src/features/repair/types.js';

// ---------------------------------------------------------------------------
// XML fixture helpers
// ---------------------------------------------------------------------------

/** Minimal document.xml wrapper. */
function makeDocXml(bodyContent: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>${bodyContent}</w:body>
</w:document>`;
}

/**
 * Build a document.xml with `rsidCount` paragraphs, each carrying one
 * w:rsidR attribute.
 */
function makeDocXmlWithRsids(rsidCount: number): string {
  const paras = Array.from({ length: rsidCount }, (_, i) =>
    `<w:p w:rsidR="AA${String(i).padStart(6, '0')}"><w:r><w:t>para ${i}</w:t></w:r></w:p>`,
  ).join('\n');
  return makeDocXml(paras);
}

/**
 * Build a settings.xml fragment containing a <w:rsids> block with N child
 * <w:rsid> entries.
 */
function makeSettingsWithRsids(count: number): string {
  const entries = Array.from(
    { length: count },
    (_, i) => `<w:rsid w:val="AA${String(i).padStart(6, '0')}"/>`,
  ).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:rsids>${entries}</w:rsids>
</w:settings>`;
}

/**
 * Minimal document.xml with one in-paragraph sectPr at the given index
 * position (zero-based).  There is always a body-level sectPr at the end.
 */
function makeDocXmlWithInParaSectPr(
  inParaSectPrXml: string,
  extra: string = '',
): string {
  return makeDocXml(`
    <w:p>
      <w:pPr>${inParaSectPrXml}</w:pPr>
      <w:r><w:t>page 1</w:t></w:r>
    </w:p>
    ${extra}
    <w:p><w:r><w:t>page 2</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
  `);
}

const PORTRAIT_SECT_PR = `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>`;
const LANDSCAPE_SECT_PR = `<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr>`;

// ---------------------------------------------------------------------------
// Finding factory helpers
// ---------------------------------------------------------------------------

function makeRsidFinding(overrides: Partial<RepairFinding> = {}): RepairFinding {
  return {
    id: 'rsid.bloat',
    category: 'rsid',
    severity: 'stabilisation',
    label: 'rsid bloat',
    count: 10,
    destructive: false,
    requiresConfirm: false,
    selectedByDefault: true,
    action: 'stripRsid',
    ...overrides,
  };
}

function makeSectionBreakFinding(
  index: number,
  overrides: Partial<RepairFinding> = {},
): RepairFinding {
  return {
    id: `sectionBreaks.ooxml.${index}`,
    category: 'sectionBreaks',
    severity: 'info',
    label: `Section break ${index + 1}`,
    count: 1,
    destructive: true,
    requiresConfirm: true,
    protected: false,
    selectedByDefault: false,
    action: 'removeSectionBreak',
    ...overrides,
  };
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('applyOoxmlSurgery()', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  // -------------------------------------------------------------------------
  // TC1 — rsid strip removes all w:rsid* attributes
  // -------------------------------------------------------------------------

  it('TC1 — rsid strip: bundle with 10 rsid attributes → 0 after apply, rsidCountAfter === 0', async () => {
    const bundle = new PizZipBundle();
    bundle.setPart('word/document.xml', makeDocXmlWithRsids(10));

    const findings = [makeRsidFinding({ count: 10 })];

    const result = await applyOoxmlSurgery(bundle, findings, []);

    const docXmlAfter = bundle.part('word/document.xml')!;
    const remainingRsids = (docXmlAfter.match(/\bw:rsid\w*\s*=/g) ?? []).length;

    expect(remainingRsids).toBe(0);
    expect(result.metricsAfter.rsidCountAfter).toBe(0);
  });

  // -------------------------------------------------------------------------
  // TC2 — <w:rsids> block removed from settings.xml
  // -------------------------------------------------------------------------

  it('TC2 — rsid strip: <w:rsids> block is removed from settings.xml', async () => {
    const bundle = new PizZipBundle();
    bundle.setPart('word/document.xml', makeDocXmlWithRsids(5));
    bundle.setPart('word/settings.xml', makeSettingsWithRsids(5));

    const findings = [makeRsidFinding({ count: 10 })];

    await applyOoxmlSurgery(bundle, findings, []);

    const settingsAfter = bundle.part('word/settings.xml')!;
    expect(settingsAfter).not.toContain('<w:rsids');
    expect(settingsAfter).not.toContain('</w:rsids>');
    // The settings element itself should survive
    expect(settingsAfter).toContain('<w:settings');
  });

  // -------------------------------------------------------------------------
  // TC3 — missing settings.xml → rsid strip completes without throwing
  // -------------------------------------------------------------------------

  it('TC3 — missing settings.xml: rsid strip still completes without throwing', async () => {
    const bundle = new PizZipBundle();
    bundle.setPart('word/document.xml', makeDocXmlWithRsids(5));
    // Deliberately do NOT add word/settings.xml

    const findings = [makeRsidFinding({ count: 5 })];

    await expect(applyOoxmlSurgery(bundle, findings, [])).resolves.toBeDefined();

    const result = await applyOoxmlSurgery(bundle, findings, []);
    expect(result.metricsAfter.rsidCountAfter).toBe(0);
  });

  // -------------------------------------------------------------------------
  // TC4 — Section break removal: confirmed in-para sectPr removed; body's
  //        final sectPr untouched
  // -------------------------------------------------------------------------

  it('TC4 — confirmed section break removal: in-para sectPr removed, body sectPr untouched', async () => {
    const bundle = new PizZipBundle();
    bundle.setPart('word/document.xml', makeDocXmlWithInParaSectPr(PORTRAIT_SECT_PR));

    const findings = [makeSectionBreakFinding(0)];

    await applyOoxmlSurgery(bundle, findings, ['sectionBreaks.ooxml.0']);

    const docXmlAfter = bundle.part('word/document.xml')!;

    // Count sectPr occurrences: body sectPr should still exist; in-para one gone
    const sectPrMatches = docXmlAfter.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g) ?? [];
    expect(sectPrMatches).toHaveLength(1); // only the body-level one

    // Verify the body sectPr (with pgSz) is still present
    expect(docXmlAfter).toContain('<w:pgSz w:w="11906" w:h="16838"/>');
  });

  // -------------------------------------------------------------------------
  // TC5 — Protected break requires the SEPARATE protected-confirm channel
  //        (hard rule, spec §5.7). Confirming only via confirmedSectionBreakIds
  //        is NOT enough; it must also be in confirmedProtectedSectionBreakIds.
  // -------------------------------------------------------------------------

  it('TC5a — protected break confirmed ONLY via confirmedSectionBreakIds → NOT removed', async () => {
    const bundle = new PizZipBundle();
    bundle.setPart('word/document.xml', makeDocXmlWithInParaSectPr(LANDSCAPE_SECT_PR));

    const findings = [makeSectionBreakFinding(0, { protected: true, severity: 'stabilisation' })];

    // In the non-protected set, but NOT in the protected set → must be skipped.
    const result = await applyOoxmlSurgery(bundle, findings, ['sectionBreaks.ooxml.0'], []);

    const docXmlAfter = bundle.part('word/document.xml')!;
    const sectPrMatches = docXmlAfter.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g) ?? [];

    // Orientation break must survive — both sectPr blocks still present.
    expect(sectPrMatches).toHaveLength(2);
    expect(docXmlAfter).toContain('w:orient="landscape"');
    const sbFinding = result.findings.find(f => f.id === 'sectionBreaks.ooxml.0');
    expect(sbFinding?.result).toBe('skipped');
  });

  it('TC5b — protected break confirmed via BOTH channels → removed', async () => {
    const bundle = new PizZipBundle();
    bundle.setPart('word/document.xml', makeDocXmlWithInParaSectPr(LANDSCAPE_SECT_PR));

    const findings = [makeSectionBreakFinding(0, { protected: true, severity: 'stabilisation' })];

    await applyOoxmlSurgery(
      bundle,
      findings,
      ['sectionBreaks.ooxml.0'],
      ['sectionBreaks.ooxml.0'],
    );

    const docXmlAfter = bundle.part('word/document.xml')!;
    const sectPrMatches = docXmlAfter.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g) ?? [];

    // Only the body-level sectPr should remain
    expect(sectPrMatches).toHaveLength(1);
    expect(docXmlAfter).not.toContain('w:orient="landscape"');
  });

  // -------------------------------------------------------------------------
  // TC6 — Section break NOT in confirmedSectionBreakIds → NOT removed,
  //        result:'skipped'
  // -------------------------------------------------------------------------

  it('TC6 — section break not confirmed → not removed, result is skipped', async () => {
    const bundle = new PizZipBundle();
    bundle.setPart('word/document.xml', makeDocXmlWithInParaSectPr(PORTRAIT_SECT_PR));

    const findings = [makeSectionBreakFinding(0)];

    // confirmedSectionBreakIds is empty — user did not confirm
    const result = await applyOoxmlSurgery(bundle, findings, []);

    const docXmlAfter = bundle.part('word/document.xml')!;
    // Both in-para sectPr and body sectPr should still be present
    const sectPrMatches = docXmlAfter.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g) ?? [];
    expect(sectPrMatches).toHaveLength(2);

    const sbFinding = result.findings.find(f => f.id === 'sectionBreaks.ooxml.0');
    expect(sbFinding?.result).toBe('skipped');
  });

  // -------------------------------------------------------------------------
  // TC7 — ooxmlService.writeBack called with onArtifactDetected:'fallbackToNewDoc'
  // -------------------------------------------------------------------------

  it('TC7 — writeBack called with onArtifactDetected: fallbackToNewDoc', async () => {
    const bundle = new PizZipBundle();
    bundle.setPart('word/document.xml', makeDocXmlWithRsids(3));

    const findings = [makeRsidFinding({ count: 3 })];

    await applyOoxmlSurgery(bundle, findings, []);

    expect(vi.mocked(ooxmlService.writeBack)).toHaveBeenCalledOnce();
    expect(vi.mocked(ooxmlService.writeBack)).toHaveBeenCalledWith(
      bundle,
      { onArtifactDetected: 'fallbackToNewDoc' },
    );
  });

  // -------------------------------------------------------------------------
  // TC8 — fileSizeAfterBytes measured from bundle.toUint8Array().length
  // -------------------------------------------------------------------------

  it('TC8 — fileSizeAfterBytes reflects bundle size (measured before writeBack)', async () => {
    const bundle = new PizZipBundle();
    bundle.setPart('word/document.xml', makeDocXmlWithRsids(5));

    const findings = [makeRsidFinding({ count: 5 })];

    const result = await applyOoxmlSurgery(bundle, findings, []);

    // After rsid strip the bundle is in its final state; measure it ourselves
    const expectedSize = bundle.toUint8Array().length;

    expect(result.metricsAfter.fileSizeAfterBytes).toBeGreaterThan(0);
    // The value returned must match what toUint8Array() produces at that state.
    // Because writeBack is mocked (no-op), the bundle is unchanged after the call,
    // so re-measuring gives the same number.
    expect(result.metricsAfter.fileSizeAfterBytes).toBe(expectedSize);
  });

  // -------------------------------------------------------------------------
  // Additional: finding result:'done' set on confirmed rsid finding
  // -------------------------------------------------------------------------

  it('rsid finding gets result:done after strip', async () => {
    const bundle = new PizZipBundle();
    bundle.setPart('word/document.xml', makeDocXmlWithRsids(5));

    const findings = [makeRsidFinding()];
    const result = await applyOoxmlSurgery(bundle, findings, []);

    const rsidFinding = result.findings.find(f => f.id === 'rsid.bloat');
    expect(rsidFinding?.result).toBe('done');
  });

  // -------------------------------------------------------------------------
  // Additional: finding result:'done' set on confirmed section-break finding
  // -------------------------------------------------------------------------

  it('confirmed section break finding gets result:done', async () => {
    const bundle = new PizZipBundle();
    bundle.setPart('word/document.xml', makeDocXmlWithInParaSectPr(PORTRAIT_SECT_PR));

    const findings = [makeSectionBreakFinding(0)];
    const result = await applyOoxmlSurgery(bundle, findings, ['sectionBreaks.ooxml.0']);

    const sbFinding = result.findings.find(f => f.id === 'sectionBreaks.ooxml.0');
    expect(sbFinding?.result).toBe('done');
  });
});
