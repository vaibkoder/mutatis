import type { Finding, FindingLocation, Severity } from '@shared/findings/index.js';
import type { TextAnchor } from '@shared/doc-model/anchor.js';

export type Locator = (anchor: TextAnchor, label: string) => FindingLocation;

/**
 * Build a Locator. The anchor's char span becomes a host-neutral location token
 * resolved by the active host at click time (`HostPort.reveal`). (ADR-0022.)
 */
export function makeLocator(): Locator {
  return (anchor, label) => ({
    paraId: anchor.paragraphId,
    charStart: anchor.startOffset,
    charEnd: anchor.endOffset,
    label,
  });
}

export function findingId(category: string, anchor: TextAnchor): string {
  return `${category}/${anchor.paragraphId}-${anchor.startOffset}-${anchor.endOffset}`;
}

/** Finding id for a paragraph-level finding (Feature B). */
export function paraFindingId(category: string, paraId: string): string {
  return `${category}/${paraId}`;
}

/** Category constants — Feature A (`featureId: 'xref'`). */
export const XREF_CATEGORIES = {
  notFound: 'cross-reference/not-found',
  ambiguous: 'cross-reference/ambiguous',
  staticRefBadge: 'cross-reference/static',
  externalMaybe: 'cross-reference/external-maybe',
} as const;

/** Category constants — Feature B (`featureId: 'staticNumbering'`). */
export const STATIC_NUM_CATEGORIES = {
  manuallyTyped: 'static-numbering/manually-typed',
} as const;

export function makeXrefFinding(
  category: string,
  severity: Severity,
  title: string,
  location: FindingLocation,
  anchor: TextAnchor,
  detail?: string,
): Finding {
  const f: Finding = {
    id: findingId(category, anchor),
    category,
    featureId: 'xref',
    severity,
    title,
    location,
  };
  if (detail !== undefined) f.detail = detail;
  return f;
}

export function makeStaticNumFinding(
  severity: Severity,
  title: string,
  location: FindingLocation,
  paraId: string,
  detail?: string,
): Finding {
  const f: Finding = {
    id: paraFindingId(STATIC_NUM_CATEGORIES.manuallyTyped, paraId),
    category: STATIC_NUM_CATEGORIES.manuallyTyped,
    featureId: 'staticNumbering',
    severity,
    title,
    location,
  };
  if (detail !== undefined) f.detail = detail;
  return f;
}
