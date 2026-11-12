/**
 * Party-Role finding construction. Reuses the defined-terms `findingId` +
 * `makeLocator` (the shared anchor→Range resolver) — only the `featureId` and
 * category set differ. (feature 09 — ADR-0016.)
 */
import type { Finding, FindingLocation } from '@shared/findings/index.js';
// Import the pure `locate.js` submodule directly, not the defined-terms
// barrel (`index.js`) — the barrel also re-exports that feature's
// store-touching `clearThenEmit`, which would drag `@shared/findings`'s
// store module (and `@preact/signals`) into any bundle reachable from here.
import { findingId } from '@features/defined-terms/locate.js';
import type { TextAnchor } from './types.js';

export const PR_CATEGORIES = {
  undefinedRole: 'party-role/undefined-role',
  inconsistentLabel: 'party-role/inconsistent-label',
} as const;

/** Build a Finding with `featureId: 'party-role'`. */
export function makePartyFinding(
  category: string,
  severity: Finding['severity'],
  title: string,
  location: FindingLocation,
  anchor: TextAnchor,
  detail?: string,
  occurrences?: FindingLocation[],
): Finding {
  const finding: Finding = {
    id: findingId(category, anchor),
    category,
    featureId: 'party-role',
    severity,
    title,
    location,
  };
  if (detail !== undefined) finding.detail = detail;
  if (occurrences && occurrences.length > 1) finding.occurrences = occurrences;
  return finding;
}

/** Numeric document-order rank from a paragraph id like "para-12" / "body-para-12". */
export function paraRank(paragraphId: string): number {
  const m = /(\d+)\s*$/.exec(paragraphId);
  return m ? Number(m[1]) : 0;
}

/** Sort key for anchors by (paragraph order, offset). */
export function anchorOrder(a: TextAnchor, b: TextAnchor): number {
  const pr = paraRank(a.paragraphId) - paraRank(b.paragraphId);
  return pr !== 0 ? pr : a.startOffset - b.startOffset;
}
