/**
 * Cross-Reference + Static Numbering driver (spec §5, §8).
 *
 * Orchestrates both features in a single doc-model pass:
 *   1. Build/cache SectionIndex (T2 builder)
 *   2. Scan prose for references → ReferenceMatch[] (T4 scanner)
 *   3. Resolve matches against index (T5 resolve)
 *   4. Read REF fields + bookmarks feature-locally (T5 fields — async, ADR-0014)
 *   5. Detect static-numbered paragraphs layers (a)+(b) (T6)
 *   6. [optional] OOXML probe layer (c) for residuals (T7 — async)
 *   7. Cross-populate isStaticNumbered on index nodes (§5.2.4)
 *   8. Emit Feature A findings (xref) + Feature B findings (staticNumbering)
 *
 * clearThenEmit: removes both 'cross-reference' and 'static-numbering' prefixes
 * before emitting fresh findings (spec §8.1, §12.3).
 *
 * Feature A and B findings are NEVER merged: featureId discriminator enforced
 * throughout (spec §5.3, §8.4).
 */

import type { ParaNode } from '@shared/doc-model/index.js';
import type { JurisdictionProfile } from '@shared/config/index.js';
import type { Finding, FindingLocation } from '@shared/findings/index.js';
import { addFindings, removeByPrefix } from '@shared/findings/index.js';
import { buildSectionIndex } from '@shared/section-index/index.js';
import type { SectionIndex } from '@shared/section-index/types.js';
import { scanForReferences } from './scanner.js';
import { resolveAll } from './resolve.js';
import { detectStaticNumbering, crossPopulateStaticNumbered } from './static-numbering.js';
import { makeLocator, makeXrefFinding, makeStaticNumFinding, XREF_CATEGORIES } from './locate.js';
import type { ReferenceMatch, StaticNumberingResult } from './types.js';

export interface CrossReferenceAnalysis {
  index: SectionIndex;
  xrefFindings: Finding[];
  staticNumFindings: Finding[];
  /** Banner: fields API was unavailable on the host (ADR-0014 fallback). */
  fieldsApiUnavailable: boolean;
  /** Count of OOXML probes run (for diagnostics). */
  ooxmlProbeCount: number;
}

/**
 * Pure analysis — no findings-store mutation.
 * The `ctx` parameter is optional: if provided, field/bookmark reads and OOXML
 * probes run; if null (e.g. unit tests without Office JS), the features degrade
 * gracefully per ADR-0014.
 */
export async function analyseCrossReference(
  paras: ReadonlyArray<ParaNode>,
  profile: JurisdictionProfile,
  ctx: Word.RequestContext | null,
): Promise<CrossReferenceAnalysis> {
  const bodyParas = paras.filter(p => p.container === 'body');
  const locate = makeLocator();

  // Step 1: Build Section Index.
  const index = buildSectionIndex(bodyParas, profile);

  // Step 2: Scan prose for reference patterns.
  const matches = scanForReferences(bodyParas, profile, index);

  // Step 3: Resolve matches against the index (pure, index-only).
  resolveAll(matches, index, profile);

  // Step 4: Feature-local field/bookmark read (ADR-0014, async).
  // This runs AFTER index resolution and applies REF-field overrides ON TOP:
  // a broken REF field (target bookmark deleted) forces notFound/error even when
  // the cited key still exists in the index (spec §5.1.4, A13). We must NOT
  // re-resolve from the index afterwards or the broken-REF verdict is clobbered.
  let fieldsApiUnavailable = false;
  if (ctx) {
    const { readFieldsAndBookmarks } = await import('./fields.js');
    const result = await readFieldsAndBookmarks(matches, paras, ctx);
    fieldsApiUnavailable = !result.fieldsApiAvailable;
  } else {
    // No context: all matches remain staticBadge=true (ADR-0014 graceful degradation).
    fieldsApiUnavailable = true;
  }

  // Step 5: Feature B — detect static-numbered paragraphs (layers a+b).
  const staticResults = detectStaticNumbering(bodyParas, profile);

  // Step 6: OOXML probe layer (c) for residuals — only when ctx available.
  //
  // Layers (a)+(b) cannot tell a hand-typed number from a field-driven one when
  // there is no listItem — both look 'static'. The OOXML probe (spec §5.2.3)
  // disambiguates: a <w:t> run with no <w:numPr>/field → confirmed static; a
  // field/numPr → dynamic (suppress the false-positive finding). We therefore
  // probe the no-listItem 'static' results, capped at maxOoxmlProbes. Beyond the
  // cap they keep the 'static' verdict (still flagged — B8). A para object is
  // built per-id for the lazy getOoxml() range.
  const paraById = new Map(paras.map(p => [p.id, p]));
  let ooxmlProbeCount = 0;
  if (ctx) {
    const { probeOoxmlLayer } = await import('./static-numbering-ooxml.js');
    const residuals = staticResults
      .filter(r => r.verdict === 'static' && !r.hasListItem)
      .map(r => {
        const para = paraById.get(r.paragraphId);
        return para ? { para, result: r } : null;
      })
      .filter((x): x is { para: ParaNode; result: StaticNumberingResult } => x !== null);

    if (residuals.length > 0) {
      await probeOoxmlLayer(residuals, ctx, profile);
      ooxmlProbeCount = Math.min(residuals.length, profile.staticNumbering.maxOoxmlProbes);
    }
  }

  // Step 7: Cross-populate isStaticNumbered on index nodes.
  crossPopulateStaticNumbered(staticResults, index.byParaId);

  // Step 8: Build findings — Feature A (xref) and Feature B (staticNumbering) SEPARATE.
  const xrefFindings = buildXrefFindings(matches, locate);
  const staticNumFindings = buildStaticNumFindings(staticResults, paras, locate);

  return { index, xrefFindings, staticNumFindings, fieldsApiUnavailable, ooxmlProbeCount };
}

function buildXrefFindings(matches: ReferenceMatch[], locate: ReturnType<typeof makeLocator>): Finding[] {
  const findings: Finding[] = [];
  for (const m of matches) {
    if (m.classification === 'ok' && !m.staticBadge) continue;

    const anchor = m.anchor;
    const location: FindingLocation = locate(anchor, m.rawText);

    if (m.classification === 'notFound') {
      findings.push(makeXrefFinding(
        XREF_CATEGORIES.notFound,
        m.severity,
        `Not found: "${m.rawText}"`,
        location,
        anchor,
        m.fieldId
          ? 'REF field target bookmark deleted or renamed.'
          : 'This reference does not match any numbered unit in the document.',
      ));
    } else if (m.classification === 'ambiguous') {
      findings.push(makeXrefFinding(
        XREF_CATEGORIES.ambiguous,
        m.severity,
        `Ambiguous: "${m.rawText}"`,
        location,
        anchor,
        `Matches ${m.candidateTargets.length} units. Use Assign source to specify the intended target.`,
      ));
    } else if (m.classification === 'externalMaybe') {
      findings.push(makeXrefFinding(
        XREF_CATEGORIES.externalMaybe,
        'info',
        `May be in a separate file: "${m.rawText}"`,
        location,
        anchor,
        'The referenced unit was not found in this document. It may be in an attached or separate file.',
      ));
    }

    // Static-reference badge (OK or ambiguous refs that are plain text, not a REF field).
    if (m.staticBadge && m.classification === 'ok') {
      findings.push(makeXrefFinding(
        XREF_CATEGORIES.staticRefBadge,
        'info',
        `Manually typed reference: "${m.rawText}"`,
        location,
        anchor,
        'This reference is plain text, not a Word REF field — Word will not keep it in sync if clause numbers change.',
      ));
    }
  }
  return findings;
}

function buildStaticNumFindings(
  results: StaticNumberingResult[],
  paras: ReadonlyArray<ParaNode>,
  locate: ReturnType<typeof makeLocator>,
): Finding[] {
  const findings: Finding[] = [];
  for (const r of results) {
    if (r.verdict !== 'static' && r.verdict !== 'indeterminate') continue;

    const para = paras.find(p => p.id === r.paragraphId);
    if (!para) continue;

    // Anchor on the leading clause number itself so navigation highlights it
    // precisely. A zero-length (paragraph-level) anchor would fall back to
    // searching the whole paragraph text, which fails for headings whose number
    // is followed by a tab + long text (Word's literal search won't match it).
    const lead = r.leadingNumberText ?? '';
    const leadIdx = lead ? para.text.indexOf(lead) : -1;
    const anchor = leadIdx >= 0
      ? { paragraphId: r.paragraphId, startOffset: leadIdx, endOffset: leadIdx + lead.length }
      : { paragraphId: r.paragraphId, startOffset: 0, endOffset: 0 };
    const snippet = para.text.slice(0, 60) + (para.text.length > 60 ? '…' : '');
    const location: FindingLocation = locate(anchor, snippet);

    const isDuplicate = r.hasListItem;
    const title = isDuplicate
      ? `Manually typed number duplicates the automatic number: "${r.leadingNumberText ?? ''}"`
      : `Manually typed clause number: "${r.leadingNumberText ?? ''}"`;
    const detail = isDuplicate
      ? 'This paragraph has both an automatic list number and a duplicate typed number. Remove the typed number.'
      : 'This number was typed by hand, not generated by Word\'s list engine — it won\'t renumber automatically.';

    findings.push(makeStaticNumFinding(r.severity, title, location, r.paragraphId, detail));
  }
  return findings;
}

/**
 * Clear previous cross-reference + static-numbering findings and emit fresh results.
 * Both prefixes are always cleared together (spec §8.1).
 *
 * The `ctx` parameter is required for full analysis; pass null in tests to get
 * the ADR-0014 degraded mode (static badges only, no field-backed detection).
 */
export async function clearThenEmit(
  paras: ReadonlyArray<ParaNode>,
  profile: JurisdictionProfile,
  ctx: Word.RequestContext | null,
): Promise<CrossReferenceAnalysis> {
  removeByPrefix('cross-reference');
  removeByPrefix('static-numbering');
  const analysis = await analyseCrossReference(paras, profile, ctx);
  addFindings([...analysis.xrefFindings, ...analysis.staticNumFindings]);
  return analysis;
}
