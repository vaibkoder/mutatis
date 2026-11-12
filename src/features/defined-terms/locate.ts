import type { Finding, FindingLocation } from '@shared/findings/index.js';
import type { TextAnchor } from './types.js';

/** A factory that turns a TextAnchor into a host-neutral FindingLocation. */
export type Locator = (anchor: TextAnchor, label: string) => FindingLocation;

/**
 * Build a Locator. The anchor's char span becomes a host-neutral location token;
 * the active host resolves it to a navigable range at click time
 * (`HostPort.reveal`), so no `Word.Range` is captured here. (ADR-0022.)
 */
export function makeLocator(): Locator {
  return (anchor, label) => ({
    paraId: anchor.paragraphId,
    charStart: anchor.startOffset,
    charEnd: anchor.endOffset,
    label,
  });
}

/** Stable finding id from a category + anchor. */
export function findingId(category: string, anchor: TextAnchor): string {
  return `${category}/${anchor.paragraphId}-${anchor.startOffset}-${anchor.endOffset}`;
}

export const CATEGORIES = {
  usedNotDefined: 'defined-terms/used-not-defined',
  definedNotUsed: 'defined-terms/defined-not-used',
  definedMultiple: 'defined-terms/defined-multiple',
  capitalisation: 'defined-terms/capitalisation',
} as const;

/** Build a Finding with the defined-terms featureId set. */
export function makeFinding(
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
    featureId: 'defined-terms',
    severity,
    title,
    location,
  };
  if (detail !== undefined) finding.detail = detail;
  if (occurrences && occurrences.length > 1) finding.occurrences = occurrences;
  return finding;
}
