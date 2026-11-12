/**
 * Feature B — layer (c) OOXML probe (spec §5.2.3).
 *
 * For paragraphs where layers (a)+(b) were inconclusive (field-driven number
 * that isn't a list item, or a number inside a field/SDT), calls
 * paragraph.getOoxml() and inspects leading content:
 *   - Leading number is <w:t> with no <w:numPr> and no field → static ('tRun')
 *   - Leading number from <w:numPr> or field element → dynamic ('listOrField')
 *   - Mixed → 'mixed', verdict = indeterminate
 *
 * CAPPED at staticNumbering.maxOoxmlProbes (default 200) to keep large documents
 * responsive. Paragraphs beyond the cap → indeterminate (low severity).
 *
 * ADR-0006: detection only — no OOXML mutation whatsoever.
 */

import type { ParaNode } from '@shared/doc-model/types.js';
import type { JurisdictionProfile } from '@shared/config/types.js';
import type { StaticNumberingResult } from './types.js';

export type OoxmlVerdict = 'tRun' | 'listOrField' | 'mixed';

/**
 * Run the OOXML probe for a set of residual-ambiguous paragraphs.
 * Mutates `results` in place, updating verdict + ooxmlVerdict for each affected para.
 *
 * `residualParas` are the ParaNodes that need probing (caller selects them from
 * the full set based on which results are still 'indeterminate' or need layer-c).
 *
 * Called inside a Word.run block.
 */
export async function probeOoxmlLayer(
  residualParas: Array<{ para: ParaNode; result: StaticNumberingResult }>,
  ctx: Word.RequestContext,
  profile: JurisdictionProfile,
): Promise<void> {
  const cap = profile.staticNumbering.maxOoxmlProbes;
  const toProbList = residualParas.slice(0, cap);
  // Beyond the cap: KEEP the layer-(a)+(b) verdict (static, still flagged) — we
  // simply forgo the extra OOXML confirmation (spec §10.7 / B8). We do NOT
  // downgrade to indeterminate, which would silently drop a real static finding.

  // For the capped set, load OOXML.
  const ooxmlByParaId = new Map<string, string>();
  const loadPromises: Array<{ paraId: string; ooxml: { value?: string } }> = [];

  for (const { para } of toProbList) {
    try {
      const paraRange = para.getRange(ctx);
      const ooxmlProp = paraRange.getOoxml();
      loadPromises.push({ paraId: para.id, ooxml: ooxmlProp as unknown as { value?: string } });
    } catch {
      // Skip paragraphs that can't be loaded.
    }
  }

  try {
    await ctx.sync();
  } catch {
    // If sync fails, leave results as-is.
    return;
  }

  for (const { paraId, ooxml } of loadPromises) {
    const xml = ooxml.value ?? '';
    ooxmlByParaId.set(paraId, xml);
  }

  // Inspect OOXML for each probed paragraph.
  for (const { para, result } of toProbList) {
    const xml = ooxmlByParaId.get(para.id) ?? '';
    if (!xml) {
      result.verdict = 'indeterminate';
      result.ooxmlVerdict = 'mixed';
      continue;
    }
    const verdict = inspectOoxml(xml);
    result.ooxmlVerdict = verdict;
    if (verdict === 'tRun') {
      result.verdict = 'static';
      result.confidence = 'high';
    } else if (verdict === 'listOrField') {
      result.verdict = 'dynamic';
    } else {
      result.verdict = 'indeterminate';
      result.confidence = 'medium';
      result.severity = 'info';
    }
  }
}

/**
 * Inspect the leading content of a paragraph's OOXML string to decide whether the
 * leading NUMBER is static (a plain `<w:t>` run) or dynamic (list/field driven).
 *
 * Heuristic (v1), based on which element governs the leading content:
 *   - `<w:numPr>` present → paragraph list numbering controls the number → 'listOrField'.
 *   - Otherwise compare positions: if a field marker (`<w:fldSimple>`/`<w:instrText>`)
 *     occurs BEFORE the first text run, the leading number sits inside the field →
 *     'listOrField'. If the first `<w:t>` occurs before any field (or there is no
 *     field at all), the leading number is a plain run → 'tRun'. A later field in
 *     the body (e.g. a cross-reference) does not make the NUMBER dynamic.
 *   - If none of these markers are found → 'mixed' (genuinely indeterminate).
 */
function inspectOoxml(xml: string): OoxmlVerdict {
  if (xml.includes('<w:numPr')) return 'listOrField';

  const firstText = firstIndexOf(xml, ['<w:t>', '<w:t ']);
  const firstField = firstIndexOf(xml, ['<w:fldSimple', '<w:instrText']);

  if (firstField === -1 && firstText === -1) return 'mixed';
  if (firstField === -1) return 'tRun';            // text run, no field
  if (firstText === -1) return 'listOrField';      // field, no standalone text
  return firstField < firstText ? 'listOrField' : 'tRun';
}

/** Smallest index at which any of the needles occurs, or -1 if none. */
function firstIndexOf(haystack: string, needles: string[]): number {
  let best = -1;
  for (const n of needles) {
    const i = haystack.indexOf(n);
    if (i !== -1 && (best === -1 || i < best)) best = i;
  }
  return best;
}
