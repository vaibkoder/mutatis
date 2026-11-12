// src/features/repair/detect/ooxml-scan.ts
import { ooxmlService } from '@shared/ooxml/index.js';
import type { DocxBundle } from '@shared/ooxml/index.js';
import type { RepairFinding, RepairMetrics } from '../types.js';

export interface OoxmlScanResult {
  findings: RepairFinding[];
  metrics: Pick<RepairMetrics, 'fileSizeBeforeBytes' | 'rsidCountBefore'>;
  /** The bundle for downstream apply; callers must not read the file again to avoid double-read */
  bundle: DocxBundle;
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

/** Count all occurrences of a regex in a string (global flag required). */
function countMatches(re: RegExp, text: string): number {
  return (text.match(re) ?? []).length;
}

/**
 * Extract the text of every <w:sectPr …>…</w:sectPr> block in document.xml.
 * Returns the blocks in document order.
 */
function extractSectPrBlocks(docXml: string): string[] {
  const re = /<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/g;
  return docXml.match(re) ?? [];
}

/** Determine whether a sectPr block contains a landscape-oriented page. */
function isLandscape(sectPrBlock: string): boolean {
  return /w:orient\s*=\s*"landscape"/i.test(sectPrBlock);
}

/** Derive a human-readable geometry description for a sectPr block. */
function geometryDescription(sectPrBlock: string): string {
  if (isLandscape(sectPrBlock)) return 'landscape orientation';
  const pgSzMatch = sectPrBlock.match(/<w:pgSz\b([^/]*\/|[^>]*>)/);
  if (pgSzMatch) {
    const wMatch = pgSzMatch[0].match(/w:w\s*=\s*"(\d+)"/);
    const hMatch = pgSzMatch[0].match(/w:h\s*=\s*"(\d+)"/);
    if (wMatch && hMatch) return `page size ${wMatch[1]}×${hMatch[1]} twips`;
  }
  return 'page geometry change';
}

/** Return true when two sectPr blocks have matching page-size/orientation. */
function sameGeometry(a: string, b: string): boolean {
  const pgSzRe = /<w:pgSz\b[^/]*(\/|[^>]*>)/;
  const aMatch = a.match(pgSzRe)?.[0] ?? '';
  const bMatch = b.match(pgSzRe)?.[0] ?? '';
  if (aMatch === '' && bMatch === '') return true;
  return aMatch === bMatch;
}

// ---------------------------------------------------------------------------
// Rsid scan
// ---------------------------------------------------------------------------

function scanRsid(docXml: string, settingsXml: string | undefined): RepairFinding[] {
  // Count w:rsid* attributes in document.xml
  const attrCount = countMatches(/\bw:rsid\w*\s*=/g, docXml);

  // Count <w:rsid> elements in settings.xml (may be absent)
  const settingsCount = settingsXml !== undefined
    ? countMatches(/<w:rsid\b/g, settingsXml)
    : 0;

  const rsidCount = attrCount + settingsCount;

  if (rsidCount === 0) return [];

  const paraCount = Math.max(1, countMatches(/<w:p[ >]/g, docXml));
  const isBloat = rsidCount > 200 || rsidCount / paraCount > 3;
  const severity: RepairFinding['severity'] = isBloat ? 'stabilisation' : 'info';

  const label = rsidCount > 200
    ? `rsid bloat: ${rsidCount} revision IDs detected`
    : `rsid attributes present (${rsidCount})`;

  const detail = isBloat
    ? 'Stripping revision IDs will reduce file size and remove revision history metadata.'
    : undefined;

  const finding: RepairFinding = {
    id: 'rsid.bloat',
    category: 'rsid',
    severity,
    label,
    ...(detail !== undefined ? { detail } : {}),
    count: rsidCount,
    destructive: false,
    requiresConfirm: false,
    selectedByDefault: true,
    action: 'stripRsid',
  };

  return [finding];
}

// ---------------------------------------------------------------------------
// Section-break scan
// ---------------------------------------------------------------------------

function scanSectionBreaks(docXml: string): RepairFinding[] {
  const blocks = extractSectPrBlocks(docXml);

  // The final <w:sectPr> in the document is the body-level one — not a section break.
  // All earlier ones are in-paragraph section breaks.
  if (blocks.length <= 1) return [];

  const inParaBlocks = blocks.slice(0, blocks.length - 1);
  const bodyBlock = blocks[blocks.length - 1]!;

  const findings: RepairFinding[] = inParaBlocks.map((block, i) => {
    const prevBlock = i === 0 ? bodyBlock : inParaBlocks[i - 1]!;
    const orientationChanges = isLandscape(block) !== isLandscape(prevBlock);
    const geometryChanges = !sameGeometry(block, prevBlock);

    let severity: RepairFinding['severity'];
    let isProtected: boolean;

    if (orientationChanges) {
      severity = 'stabilisation';
      isProtected = true;
    } else if (geometryChanges) {
      severity = 'stabilisation';
      isProtected = false;
    } else {
      severity = 'info';
      isProtected = false;
    }

    const label = `Section break ${i + 1}${isProtected ? ' (orientation-changing — landscape/portrait)' : ''}`;
    const detail = orientationChanges
      ? `Orientation differs from adjacent section: ${geometryDescription(block)}`
      : geometryChanges
        ? `Page geometry differs from adjacent section: ${geometryDescription(block)}`
        : 'Section break with same page geometry as adjacent section';

    const finding: RepairFinding = {
      id: `sectionBreaks.ooxml.${i}`,
      category: 'sectionBreaks',
      severity,
      label,
      detail,
      count: 1,
      destructive: true,
      requiresConfirm: true,
      protected: isProtected,
      selectedByDefault: false,
      action: 'removeSectionBreak',
    };

    return finding;
  });

  return findings;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

export async function detectOoxml(): Promise<OoxmlScanResult> {
  const bundle = await ooxmlService.readDocx();

  // File-size measurement
  // PizZipBundle.toUint8Array() re-serialises — use it as-is; the cost is
  // acceptable at scan time and avoids a separate byte-count pass.
  const fileSizeBeforeBytes = (bundle as { toUint8Array?(): Uint8Array }).toUint8Array?.().length ?? 0;

  const docXml = bundle.part('word/document.xml') ?? '';
  const settingsXml = bundle.hasPart('word/settings.xml')
    ? bundle.part('word/settings.xml')
    : undefined;

  // Count rsids for the metrics record (document.xml attributes only, matching
  // what the apply phase will strip; settings.xml element count is additive).
  const rsidCountBefore = countMatches(/\bw:rsid\w*\s*=/g, docXml)
    + (settingsXml !== undefined ? countMatches(/<w:rsid\b/g, settingsXml) : 0);

  const findings: RepairFinding[] = [
    ...scanRsid(docXml, settingsXml),
    ...scanSectionBreaks(docXml),
  ];

  return {
    findings,
    metrics: { fileSizeBeforeBytes, rsidCountBefore },
    bundle,
  };
}
