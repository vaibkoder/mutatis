/**
 * Feature 06 — Numbering: edge-case detection and warnings (M7).
 *
 * Provides standalone detectors for the §8 edge cases that the driver and
 * apply engine surface as banners or pre-op confirmations.
 *
 * All detectors are non-throwing and best-effort — they return a boolean/string
 * result; callers decide what to do with it. No document mutations here.
 */

import type { SectionIndex } from '@shared/section-index/index.js';

// ---------------------------------------------------------------------------
// §8.4 — Multi-section detection (pure — uses the section index)
// ---------------------------------------------------------------------------

/**
 * Returns true when the section index spans more than one document section
 * (i.e., the index contains both operative clauses AND schedules/annexes as
 * separate kind groups, suggesting Word section breaks between them).
 *
 * Used as a fast pre-check before going to the Office API.
 * The authoritative check uses detectMultipleSections() in apply.ts.
 */
export function hasMultipleSectionKinds(index: SectionIndex): boolean {
  const kinds = new Set(index.nodes.map(n => n.kind));
  return kinds.size > 1;
}

// ---------------------------------------------------------------------------
// §8.6 — Tracked-changes detection
// ---------------------------------------------------------------------------

/**
 * Detect whether Track Changes is currently on for the document.
 * Returns true if tracked, false if off, null if the API is unavailable.
 */
export async function detectTrackChanges(ctx: Word.RequestContext): Promise<boolean | null> {
  try {
    const settings = ctx.document.changeTrackingMode;
    ctx.load(ctx.document, 'changeTrackingMode');
    await ctx.sync();
    const mode = settings as unknown as string;
    return mode !== 'Off' && mode !== 'off' && mode !== 'None' && mode !== 'none';
  } catch {
    return null; // API unavailable — cannot determine
  }
}

/**
 * Message to show when Track Changes is detected before a level change (§8.6).
 * The operation still proceeds — this is a warn-before, not a block.
 */
export const TRACK_CHANGES_WARNING =
  'Track Changes is on — level changes will be recorded as revisions.';

// ---------------------------------------------------------------------------
// §8.7 — Style-linked list detection (best-effort, Office JS only)
// ---------------------------------------------------------------------------

/**
 * Warning message shown when a style-linked list is detected.
 * The list is NOT modified — only direct-format lists are touched.
 */
export const STYLE_LINKED_WARNING =
  'This list is linked to a paragraph style; it was not modified to avoid breaking your template.';

// ---------------------------------------------------------------------------
// §8.2 — Footnote scope message
// ---------------------------------------------------------------------------

/**
 * Message shown when the user tries to apply/promote inside a footnote context.
 */
export const FOOTNOTE_DISABLED_MESSAGE =
  'Numbering apply and promote/demote are disabled inside footnotes.';

// ---------------------------------------------------------------------------
// Banner strings (§7.6)
// ---------------------------------------------------------------------------

/** Compose the Word Online partial-support banner from a list of degraded levels. */
export function wordOnlineBanner(degradedLevels: number[]): string {
  if (degradedLevels.length === 0) return '';
  return `Word Online: scheme applied with reduced support — levels ${degradedLevels.join(', ')} may not display correctly.`;
}

/** Multi-section restart warning (§8.4). */
export const MULTI_SECTION_WARNING =
  'This document has multiple sections; top-level clause numbering may restart at a section break — verify manually.';

/** Static-numbering skip message (§8.3 / ADR-0006). */
export function staticSkippedMessage(count: number): string {
  return `${count} paragraph(s) appear to use manually typed numbers and were not converted (see the Proof tab's Manually-typed clause numbers section).`;
}
