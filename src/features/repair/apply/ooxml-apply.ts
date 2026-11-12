// src/features/repair/apply/ooxml-apply.ts
import { ooxmlService } from '@shared/ooxml/index.js';
import type { DocxBundle } from '@shared/ooxml/index.js';
import type { RepairFinding, RepairMetrics } from '../types.js';

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

export interface OoxmlApplyResult {
  findings: RepairFinding[];
  metricsAfter: Pick<RepairMetrics, 'fileSizeAfterBytes' | 'rsidCountAfter'>;
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

/**
 * Strip all w:rsid* attributes from document.xml and remove the <w:rsids>
 * block from settings.xml (if present).
 *
 * Returns the rsidCountAfter measured from the document.xml string that
 * remains in the bundle after stripping.
 */
function stripRsid(bundle: DocxBundle): number {
  // Strip attributes from document.xml
  bundle.stripAttributes('word/document.xml', ['w:rsid*']);

  // Remove <w:rsids> element from settings.xml (may be absent)
  if (bundle.hasPart('word/settings.xml')) {
    const settings = bundle.part('word/settings.xml')!;
    const cleaned = settings
      .replace(/<w:rsids\b[^>]*>[\s\S]*?<\/w:rsids>/g, '')
      .replace(/<w:rsids\s*\/>/g, '');
    bundle.setPart('word/settings.xml', cleaned);
  }

  // Re-measure (should be 0 after a successful strip)
  const docXml = bundle.part('word/document.xml') ?? '';
  return (docXml.match(/\bw:rsid\w*\s*=/g) ?? []).length;
}

/**
 * Remove a single in-paragraph <w:sectPr> from document.xml by zero-based index.
 *
 * Strategy (string-based, preserves fidelity):
 *   1. Collect all <w:sectPr …>…</w:sectPr> occurrences.
 *   2. Skip the last one (the body-level sectPr — must be preserved).
 *   3. Replace the Nth in-paragraph match by removing it from the XML string.
 *
 * The paragraph element itself is NOT removed — only the <w:sectPr> child
 * that lives inside <w:pPr>.
 */
function removeSectionBreakAtIndex(docXml: string, index: number): string {
  const sectPrRe = /<w:sectPr\b[\s\S]*?<\/w:sectPr>/g;
  const allMatches: Array<{ match: string; index: number }> = [];

  let m: RegExpExecArray | null;
  while ((m = sectPrRe.exec(docXml)) !== null) {
    allMatches.push({ match: m[0], index: m.index });
  }

  // The last entry is the body-level sectPr — never touch it.
  // In-paragraph sectPrs are everything before that.
  const inParaMatches = allMatches.slice(0, allMatches.length - 1);

  if (index < 0 || index >= inParaMatches.length) {
    // Nothing to remove; return unchanged.
    return docXml;
  }

  const target = inParaMatches[index]!;

  // Remove the target match from the string.
  return (
    docXml.slice(0, target.index) +
    docXml.slice(target.index + target.match.length)
  );
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Apply OOXML-level surgery to the provided bundle:
 *   - Strip w:rsid* attributes (action: 'stripRsid')
 *   - Remove confirmed in-paragraph section breaks (action: 'removeSectionBreak')
 *
 * The bundle is mutated in place. `ooxmlService.writeBack` is called once,
 * after all mutations, with `onArtifactDetected: 'fallbackToNewDoc'`.
 *
 * @param bundle                  The DocxBundle from the T7 detectOoxml pass.
 *                                Must NOT be re-read from the document here.
 * @param findings                The RepairFinding list from the scan phase.
 * @param confirmedSectionBreakIds           Finding ids of NON-protected section
 *                                         breaks the user confirmed for removal.
 * @param confirmedProtectedSectionBreakIds Finding ids of ORIENTATION-CHANGING
 *                                         (protected:true) section breaks the
 *                                         user explicitly confirmed via the
 *                                         per-break orientation-change modal.
 *                                         Kept as a separate channel so the
 *                                         apply layer can enforce the hard rule:
 *                                         orientation breaks are NEVER removed
 *                                         unless this separate set contains their
 *                                         id. (spec §5.7 hard rule / ADR-0006)
 */
export async function applyOoxmlSurgery(
  bundle: DocxBundle,
  findings: RepairFinding[],
  confirmedSectionBreakIds: string[],
  confirmedProtectedSectionBreakIds: string[] = [],
): Promise<OoxmlApplyResult> {
  const updatedFindings: RepairFinding[] = findings.map(f => ({ ...f }));
  let rsidCountAfter: number | null = null;

  // -------------------------------------------------------------------------
  // Pass 1: rsid strip
  // -------------------------------------------------------------------------
  for (const finding of updatedFindings) {
    if (finding.action !== 'stripRsid') continue;
    rsidCountAfter = stripRsid(bundle);
    finding.result = 'done';
  }

  // -------------------------------------------------------------------------
  // Pass 2: section break removal
  // -------------------------------------------------------------------------
  // Process confirmed removals in id order (which corresponds to document
  // order since ids are 'sectionBreaks.ooxml.N').
  //
  // IMPORTANT: because removing one sectPr shifts the indices of subsequent
  // ones, we must process from highest index to lowest to keep earlier
  // indices stable.
  const confirmedSet = new Set(confirmedSectionBreakIds);
  const confirmedProtectedSet = new Set(confirmedProtectedSectionBreakIds);

  const confirmedSbFindings = updatedFindings
    .filter(f => {
      if (f.action !== 'removeSectionBreak') return false;
      if (!confirmedSet.has(f.id)) return false;
      // Hard rule (spec §5.7): orientation-changing breaks MUST also appear
      // in the separate confirmedProtectedSectionBreakIds channel.
      if (f.protected === true && !confirmedProtectedSet.has(f.id)) return false;
      return true;
    })
    .sort((a, b) => {
      // Sort descending by numeric suffix so we remove from the end first.
      const indexA = parseInt(a.id.split('.').pop() ?? '0', 10);
      const indexB = parseInt(b.id.split('.').pop() ?? '0', 10);
      return indexB - indexA;
    });

  const removedIds = new Set<string>();
  for (const finding of confirmedSbFindings) {
    const indexSuffix = finding.id.split('.').pop() ?? '0';
    const targetIndex = parseInt(indexSuffix, 10);

    const docXml = bundle.part('word/document.xml') ?? '';
    const updated = removeSectionBreakAtIndex(docXml, targetIndex);
    bundle.setPart('word/document.xml', updated);
    finding.result = 'done';
    removedIds.add(finding.id);
  }

  // Mark any section-break finding that was NOT actually removed as skipped.
  // This includes both unconfirmed breaks and protected breaks that lacked the
  // separate protected-confirmation channel (the hard-rule block above).
  for (const finding of updatedFindings) {
    if (finding.action === 'removeSectionBreak' && !removedIds.has(finding.id)) {
      finding.result = 'skipped';
    }
  }

  // -------------------------------------------------------------------------
  // Measure post-apply metrics BEFORE writing back (bundle represents final state)
  // -------------------------------------------------------------------------
  const fileSizeAfterBytes =
    (bundle as { toUint8Array?(): Uint8Array }).toUint8Array?.().length ?? 0;

  // If no rsid strip was performed, measure current state.
  if (rsidCountAfter === null) {
    const docXml = bundle.part('word/document.xml') ?? '';
    rsidCountAfter = (docXml.match(/\bw:rsid\w*\s*=/g) ?? []).length;
  }

  // -------------------------------------------------------------------------
  // Write back
  // -------------------------------------------------------------------------
  await ooxmlService.writeBack(bundle, { onArtifactDetected: 'fallbackToNewDoc' });

  return {
    findings: updatedFindings,
    metricsAfter: {
      fileSizeAfterBytes,
      rsidCountAfter,
    },
  };
}
