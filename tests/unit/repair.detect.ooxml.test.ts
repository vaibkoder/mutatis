import { describe, it, expect, vi, beforeEach } from 'vitest';
import { PizZipBundle } from '@shared/ooxml/index.js';

// ---------------------------------------------------------------------------
// Mock ooxmlService before importing the module under test.
// vi.mock is hoisted, so this executes before the import below.
// ---------------------------------------------------------------------------

vi.mock('@shared/ooxml/index.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@shared/ooxml/index.js')>();
  return {
    ...actual,
    ooxmlService: {
      readDocx: vi.fn(),
      writeBack: vi.fn(),
    },
  };
});

import { ooxmlService } from '@shared/ooxml/index.js';
import { detectOoxml } from '../../src/features/repair/detect/ooxml-scan.js';

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

/** Build a paragraph element with N w:rsidR attributes spread across it. */
function makeParasWithRsids(count: number): string {
  const attr = 'w:rsidR="AABBCC00"';
  const rsidAttrs = Array.from({ length: count }, () => attr).join(' ');
  return `<w:p ${rsidAttrs}><w:r><w:t>text</w:t></w:r></w:p>`;
}

/**
 * Build a document with N paragraphs, each contributing one rsid attribute.
 * This makes rsidCount === paraCount, so the ratio heuristic stays below 3.
 */
function makeDocXmlWithRsids(rsidCount: number): string {
  const paras = Array.from({ length: rsidCount }, (_, i) =>
    `<w:p w:rsidR="AA${String(i).padStart(6, '0')}"><w:r><w:t>text ${i}</w:t></w:r></w:p>`
  ).join('\n');
  return makeDocXml(paras);
}

/** Document with a single in-paragraph sectPr (same geometry as body sectPr). */
function makeDocXmlWithInParaSectPr(inParaSectPrXml: string): string {
  return makeDocXml(`
    <w:p>
      <w:pPr>${inParaSectPrXml}</w:pPr>
      <w:r><w:t>page 1</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>page 2</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
  `);
}

/** Returns a sectPr block with portrait geometry (matching body default). */
const PORTRAIT_SECT_PR = `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>`;

/** Returns a sectPr block with landscape orientation. */
const LANDSCAPE_SECT_PR = `<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr>`;

/** Utility: return the mocked readDocx function with correct type. */
function mockReadDocx(bundle: PizZipBundle): void {
  vi.mocked(ooxmlService.readDocx).mockResolvedValue(bundle);
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('detectOoxml()', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  // -------------------------------------------------------------------------
  // rsid tests
  // -------------------------------------------------------------------------

  it('TC1 — 250 rsid attributes → rsid.bloat finding with severity stabilisation', async () => {
    // 250 rsid attrs on a single paragraph → ratio >> 3 → bloat
    const docXml = makeDocXml(makeParasWithRsids(250));
    const bundle = new PizZipBundle();
    bundle.setPart('word/document.xml', docXml);
    mockReadDocx(bundle);

    const result = await detectOoxml();

    const rsidFinding = result.findings.find(f => f.id === 'rsid.bloat');
    expect(rsidFinding).toBeDefined();
    expect(rsidFinding!.severity).toBe('stabilisation');
    expect(rsidFinding!.selectedByDefault).toBe(true);
    expect(rsidFinding!.category).toBe('rsid');
    expect(rsidFinding!.action).toBe('stripRsid');
    expect(rsidFinding!.destructive).toBe(false);
    expect(rsidFinding!.requiresConfirm).toBe(false);
    expect(rsidFinding!.count).toBeGreaterThanOrEqual(250);
  });

  it('TC2 — 5 rsid attributes → rsid finding present with severity info', async () => {
    // 5 rsids across 5 paragraphs → ratio === 1 → not bloat
    const docXml = makeDocXmlWithRsids(5);
    const bundle = new PizZipBundle();
    bundle.setPart('word/document.xml', docXml);
    mockReadDocx(bundle);

    const result = await detectOoxml();

    const rsidFinding = result.findings.find(f => f.id === 'rsid.bloat');
    expect(rsidFinding).toBeDefined();
    expect(rsidFinding!.severity).toBe('info');
    expect(rsidFinding!.selectedByDefault).toBe(true);
  });

  it('TC3 — 0 rsid attributes → no rsid finding emitted', async () => {
    const docXml = makeDocXml('<w:p><w:r><w:t>no rsids here</w:t></w:r></w:p>');
    const bundle = new PizZipBundle();
    bundle.setPart('word/document.xml', docXml);
    mockReadDocx(bundle);

    const result = await detectOoxml();

    const rsidFinding = result.findings.find(f => f.id === 'rsid.bloat');
    expect(rsidFinding).toBeUndefined();
  });

  // -------------------------------------------------------------------------
  // Section break tests
  // -------------------------------------------------------------------------

  it('TC4 — one in-paragraph sectPr with same geometry → 1 section break finding, protected:false', async () => {
    // In-para sectPr identical geometry to body sectPr
    const docXml = makeDocXmlWithInParaSectPr(PORTRAIT_SECT_PR);
    const bundle = new PizZipBundle();
    bundle.setPart('word/document.xml', docXml);
    mockReadDocx(bundle);

    const result = await detectOoxml();

    const sbFindings = result.findings.filter(f => f.category === 'sectionBreaks');
    expect(sbFindings).toHaveLength(1);
    expect(sbFindings[0]!.id).toBe('sectionBreaks.ooxml.0');
    expect(sbFindings[0]!.protected).toBe(false);
    expect(sbFindings[0]!.severity).toBe('info');
    expect(sbFindings[0]!.destructive).toBe(true);
    expect(sbFindings[0]!.requiresConfirm).toBe(true);
    expect(sbFindings[0]!.selectedByDefault).toBe(false);
    expect(sbFindings[0]!.action).toBe('removeSectionBreak');
  });

  it('TC5 — in-paragraph sectPr with landscape orientation → finding with protected:true, severity stabilisation', async () => {
    // Body sectPr is portrait; in-para sectPr is landscape → orientation change
    const docXml = makeDocXmlWithInParaSectPr(LANDSCAPE_SECT_PR);
    const bundle = new PizZipBundle();
    bundle.setPart('word/document.xml', docXml);
    mockReadDocx(bundle);

    const result = await detectOoxml();

    const sbFindings = result.findings.filter(f => f.category === 'sectionBreaks');
    expect(sbFindings).toHaveLength(1);
    expect(sbFindings[0]!.protected).toBe(true);
    expect(sbFindings[0]!.severity).toBe('stabilisation');
    expect(sbFindings[0]!.label).toContain('orientation-changing');
  });

  // -------------------------------------------------------------------------
  // Robustness: absent settings.xml
  // -------------------------------------------------------------------------

  it('TC6 — absent settings.xml → scan completes without throwing', async () => {
    // Bundle with document.xml only; no settings.xml
    const docXml = makeDocXmlWithRsids(10);
    const bundle = new PizZipBundle();
    bundle.setPart('word/document.xml', docXml);
    // Deliberately do NOT add word/settings.xml
    mockReadDocx(bundle);

    await expect(detectOoxml()).resolves.toBeDefined();

    const result = await detectOoxml();
    // rsid finding should still be present from document.xml alone
    const rsidFinding = result.findings.find(f => f.id === 'rsid.bloat');
    expect(rsidFinding).toBeDefined();
    expect(rsidFinding!.count).toBeGreaterThan(0);
  });

  // -------------------------------------------------------------------------
  // Metrics
  // -------------------------------------------------------------------------

  it('returns fileSizeBeforeBytes > 0 and rsidCountBefore in metrics', async () => {
    const docXml = makeDocXmlWithRsids(5);
    const bundle = new PizZipBundle();
    bundle.setPart('word/document.xml', docXml);
    mockReadDocx(bundle);

    const result = await detectOoxml();

    expect(result.metrics.fileSizeBeforeBytes).toBeGreaterThan(0);
    expect(result.metrics.rsidCountBefore).toBe(5);
  });

  it('returns the bundle so downstream callers can avoid a double-read', async () => {
    const docXml = makeDocXml('<w:p><w:r><w:t>hello</w:t></w:r></w:p>');
    const bundle = new PizZipBundle();
    bundle.setPart('word/document.xml', docXml);
    mockReadDocx(bundle);

    const result = await detectOoxml();

    expect(result.bundle).toBeDefined();
    expect(result.bundle.part('word/document.xml')).toContain('hello');
    // readDocx should have been called exactly once
    expect(vi.mocked(ooxmlService.readDocx)).toHaveBeenCalledTimes(1);
  });
});
