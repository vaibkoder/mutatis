/**
 * Feature B — Static Numbering Detector, layers (a) and (b).
 * spec §5.2.1 (pattern match) + §5.2.2 (list-membership check).
 *
 * Layer (c) OOXML probe lives in static-numbering-ooxml.ts (T7).
 *
 * ADR-0006: detect-and-flag only — NO mutation, NO OOXML list surgery.
 */

import type { ParaNode } from '@shared/doc-model/types.js';
import type { JurisdictionProfile, StaticSeverityLabel } from '@shared/config/types.js';
import type { Severity } from '@shared/findings/types.js';
import type { StaticNumberingResult, NumberingVerdict } from './types.js';

/**
 * Map a conceptual static-numbering severity label (spec §5.2.5) to the shared
 * Finding `Severity` union. The Findings model has only error/warning/info.
 */
export function toFindingSeverity(label: StaticSeverityLabel): Severity {
  return label === 'medium' ? 'warning' : 'info';
}

// spec §5.2.1 leading-token regexes. Conservative against prose false positives
// (see builder.ts): dotted decimals allow a trailing space; bare integers require
// `.`, `)`, or tab so "5 days after…" / "1976 Act" / "30 June 2026" are NOT flagged.
//
// Exported as the single canonical vocabulary so the Numbering-upgrade token-span
// detector (feature 06 §11.1, ADR-0023) reuses these patterns rather than authoring
// a 4th copy. They are anchored at start and stateless (no /g).
export const DECIMAL_DOTTED_RE = /^\s*(\d+(?:\.\d+)+)[.)\t ]/;
export const DECIMAL_BARE_RE = /^\s*(\d+)[.)\t]/;
export const LETTER_RE = /^\s*\(?([a-z])\)[ \t]/;
export const ROMAN_RE = /^\s*\(?(i{1,3}|iv|vi{0,3}|ix|xi{0,2}|xii)\)[ \t]/i;
export const UPPER_RE = /^\s*\(?([A-Z])\)[ \t]/;

/** Returns the leading number text if the paragraph text matches any static-numbering pattern. */
export function detectLeadingNumber(text: string): string | null {
  const t = text.trim();
  const dm = DECIMAL_DOTTED_RE.exec(t) ?? DECIMAL_BARE_RE.exec(t);
  if (dm && dm[1]) return dm[1];
  const lm = LETTER_RE.exec(t);
  if (lm && lm[1]) return `(${lm[1]})`;
  const rm = ROMAN_RE.exec(t);
  if (rm && rm[1]) return `(${rm[1].toLowerCase()})`;
  const um = UPPER_RE.exec(t);
  if (um && um[1]) return `(${um[1]})`;
  return null;
}

/**
 * Run Feature B layers (a) + (b) over body paragraphs.
 *
 * Returns one StaticNumberingResult per paragraph that has a leading number token,
 * with `verdict` set to:
 *   - `'dynamic'`       — layer (a) matched AND list item present (normal case, no finding)
 *   - `'static'`        — layer (a) matched AND no list item → hand-typed
 *   - `'static'` (dup)  — layer (a) matched AND list item present but texts differ → duplicate-on-list
 *   - `'noNumber'`      — no leading number detected
 *   - `'indeterminate'` — list item present but leading number absent, OR inconclusive layer (b)
 *     (these cases go to layer c in static-numbering-ooxml.ts)
 *
 * Paragraphs excluded by TOC style are skipped. Non-body containers are skipped.
 */
export function detectStaticNumbering(
  paras: ReadonlyArray<ParaNode>,
  profile: JurisdictionProfile,
): StaticNumberingResult[] {
  const tocStyles = new Set(profile.xref.styles.tocStyles);
  const headingStyles = new Set(profile.xref.styles.headingStyles);
  const headingSeverity = profile.staticNumbering.headingSeverity;
  const bodySeverity = profile.staticNumbering.bodySeverity;
  const results: StaticNumberingResult[] = [];

  for (const para of paras) {
    if (para.container !== 'body') continue;
    if (tocStyles.has(para.styleName)) continue;

    const leadingNumberText = detectLeadingNumber(para.text ?? '');
    const hasListItem = (para.listString?.length ?? 0) > 0;
    const listStringIfAny: string | null = para.listString !== undefined ? para.listString : null;

    // noNumber: no leading token
    if (!leadingNumberText) {
      // Only produce a result if there's something to report; skip noNumber for cleanliness.
      continue;
    }

    const isHeading = headingStyles.has(para.styleName);
    let verdict: NumberingVerdict;
    let confidence: 'high' | 'medium';
    let severityLabel: StaticSeverityLabel;

    if (hasListItem) {
      // Layer (b): has list item.
      // Check for duplicate-on-list: both a listString-rendered number and a hand-typed leading token.
      // If the leading text matches the listString (ignoring format), it's a duplicate → flag.
      const cleanList = listStringIfAny?.replace(/[().]/g, '').trim() ?? '';
      const cleanLeading = leadingNumberText.replace(/[().]/g, '').trim();
      if (cleanList && cleanList.toLowerCase() !== cleanLeading.toLowerCase()) {
        // Leading token doesn't match list number — might be prose, not a duplicate numbering.
        // Dynamic — no finding needed; fall through to next para.
        verdict = 'dynamic';
        confidence = 'high';
        severityLabel = 'low';
      } else {
        // Leading token matches list number: duplicate-on-list (medium, unusual).
        verdict = 'static';
        confidence = 'high';
        severityLabel = 'medium';
      }
    } else {
      // No list item + leading number → static.
      verdict = 'static';
      confidence = 'high';
      severityLabel = isHeading ? headingSeverity : bodySeverity;
    }

    // Emit a result only for static/indeterminate paragraphs.
    if (verdict === 'dynamic') continue;

    results.push({
      paragraphId: para.id,
      leadingNumberText,
      hasListItem,
      listStringIfAny,
      verdict,
      confidence,
      severity: toFindingSeverity(severityLabel),
    });
  }

  return results;
}

/**
 * Cross-populate `isStaticNumbered` on Section Index nodes for all 'static' verdicts.
 * Called after buildSectionIndex and detectStaticNumbering (spec §5.2.4).
 *
 * Mutates `byParaId` entries in place — safe because SectionNode is a mutable object
 * inside the frozen ReadonlyMap.
 */
export function crossPopulateStaticNumbered(
  results: StaticNumberingResult[],
  byParaId: ReadonlyMap<string, { isStaticNumbered: boolean }>,
): void {
  for (const r of results) {
    if (r.verdict === 'static') {
      const node = byParaId.get(r.paragraphId);
      if (node) node.isStaticNumbered = true;
    }
  }
}
