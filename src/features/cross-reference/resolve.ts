/**
 * Feature A — resolution (spec §5.1.3).
 *
 * Resolves ReferenceMatch objects against the SectionIndex:
 *   - [] → Not Found (unless externalMaybe applies)
 *   - [one] → OK
 *   - [many] → Ambiguous
 *   - exact key absent but prefix children present → Ambiguous (implied-parent, spec §5.1.3 step 4)
 *   - outer unit missing + externalMaybe heuristic → externalMaybe info
 *
 * Does NOT call any Office JS — purely operates on already-loaded data.
 */

import type { SectionIndex } from '@shared/section-index/types.js';
import type { JurisdictionProfile } from '@shared/config/types.js';
import type { ReferenceMatch } from './types.js';

/**
 * Resolve a single reference match.
 *
 * For compound references (expandedKeys length > 1) each key is resolved;
 * the overall classification is the worst result. A range like "Schedules 2 to 5"
 * produces one Finding per broken key but a single overall match.
 *
 * Mutates `match` in place (classification, candidateTargets, severity, staticBadge).
 */
export function resolveMatch(
  match: ReferenceMatch,
  index: SectionIndex,
  profile: JurisdictionProfile,
): void {
  const keys = match.expandedKeys.length > 0 ? match.expandedKeys : [match.normalizedKey];

  // Handle "bare paragraph N of that Schedule" placeholder
  if (match.normalizedKey.startsWith('schedule:?/')) {
    // Can't resolve without context; treat as not-found for now
    match.classification = 'notFound';
    match.severity = 'warning';
    return;
  }

  const allResults = keys.map(k => ({
    key: k,
    nodes: index.resolve(k),
    prefixAmbiguous: !index.has(k) && hasPrefixChildren(index, k),
  }));

  // Check for compound both-halves validation (schedule:N/para:M)
  // Both the outer schedule and inner paragraph must exist.
  for (const r of allResults) {
    if (r.key.includes('/para:')) {
      const outerKey = r.key.split('/para:')[0] ?? r.key;
      if (!index.has(outerKey)) {
        // Outer unit missing: try externalMaybe
        if (shouldBeExternalMaybe(outerKey, profile)) {
          match.classification = 'externalMaybe';
          match.severity = 'info';
          return;
        }
        match.classification = 'notFound';
        match.severity = 'warning';
        return;
      }
    }
  }

  // Any of the keys is not-found → overall not-found (unless externalMaybe)
  const notFoundKeys = allResults.filter(r => r.nodes.length === 0 && !r.prefixAmbiguous);
  const ambiguousKeys = allResults.filter(r => r.nodes.length > 1 || r.prefixAmbiguous);
  const okKeys = allResults.filter(r => r.nodes.length === 1 && !r.prefixAmbiguous);

  if (notFoundKeys.length > 0) {
    const nfEntry = notFoundKeys[0];
    if (!nfEntry) return;
    const key = nfEntry.key;
    if (shouldBeExternalMaybe(key, profile) && okKeys.length === 0) {
      match.classification = 'externalMaybe';
      match.severity = 'info';
      return;
    }
    match.classification = 'notFound';
    match.severity = 'warning';
    return;
  }

  if (ambiguousKeys.length > 0) {
    match.classification = 'ambiguous';
    match.severity = 'warning';
    // Gather candidate targets from prefix children or direct ambiguous entries
    for (const r of ambiguousKeys) {
      if (r.prefixAmbiguous) {
        match.candidateTargets.push(...getPrefixChildren(index, r.key));
      } else {
        match.candidateTargets.push(...r.nodes);
      }
    }
    return;
  }

  // All keys resolved uniquely
  match.classification = 'ok';
  match.severity = match.staticBadge ? 'info' : 'info';
}

/**
 * Check whether the index has children with keys that start with `parentKey + ':'` or `parentKey + '('`.
 * Detects implied-parent ambiguity: "Clause 5" absent but "clause:5(a)", "clause:5(b)" present.
 */
function hasPrefixChildren(index: SectionIndex, key: string): boolean {
  return getPrefixChildren(index, key).length > 0;
}

function getPrefixChildren(index: SectionIndex, key: string): import('@shared/section-index/types.js').SectionNode[] {
  const prefix = key + '(';
  const prefix2 = key + '.';
  const results: import('@shared/section-index/types.js').SectionNode[] = [];
  for (const [k, nodes] of index.byKey) {
    if (k.startsWith(prefix) || k.startsWith(prefix2)) {
      results.push(...nodes);
    }
  }
  return results;
}

/**
 * Heuristic: should a missing outer unit be reported as "may be in a separate file"?
 * True only for TOP-LEVEL schedule/annex/appendix keys (not compound para refs).
 * A compound key like "schedule:2/para:3" where the outer schedule exists but the
 * inner paragraph doesn't is a definite Not Found, not externalMaybe.
 */
function shouldBeExternalMaybe(key: string, _profile: JurisdictionProfile): boolean {
  if (key.includes('/para:')) return false;
  return key.startsWith('schedule:') || key.startsWith('annex:') || key.startsWith('appendix:');
}

/**
 * Resolve all reference matches in-place.
 */
export function resolveAll(
  matches: ReferenceMatch[],
  index: SectionIndex,
  profile: JurisdictionProfile,
): void {
  for (const match of matches) {
    resolveMatch(match, index, profile);
  }
}
