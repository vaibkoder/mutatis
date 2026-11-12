/**
 * Feature 06 — Numbering: cross-reference warn + highlight review (M6).
 *
 * Spec §5.3: after any apply or promote/demote, search for legal reference
 * patterns and highlight candidates so the user can manually verify them.
 *
 * IMPORTANT:
 *  - We NEVER edit reference text (spec §5.3 / ❌ cut).
 *  - We only apply/remove a specific highlight colour (our "review sentinel").
 *  - Clear only removes our own sentinel colour, never the user's highlights.
 */

import type { JurisdictionProfile } from '@shared/config/index.js';
import type { ReviewState, ReviewItem } from './types.js';
import { getNumberingConfig } from '@shared/config/index.js';

/**
 * Sentinel highlight — must be one of Word's 15 named palette colours.
 * Arbitrary hex strings throw "argument not valid" in Word Online.
 */
const DEFAULT_HIGHLIGHT = 'Yellow';

// ---------------------------------------------------------------------------
// Pattern generation (pure — testable)
// ---------------------------------------------------------------------------

/**
 * Build the set of Word wildcard patterns for the cross-reference review.
 * Patterns are derived from the config unit nouns so they are always
 * jurisdiction-correct (Clause/paragraph/Schedule/Part/recital).
 *
 * Word wildcard syntax: `@` = 1-or-more of preceding; `\(…\)` = literal parens.
 * Patterns are intentionally broad — the user reviews each candidate.
 */
export function buildReviewPatterns(profile: JurisdictionProfile): string[] {
  const primary = profile.xref.units.primaryBody;           // 'Clause'
  const scheduleWord = profile.xref.units.annexurePrimary;  // 'Schedule'
  const subUnit = profile.xref.units.scheduleSubUnit;       // 'paragraph'
  const otherDivs = profile.xref.units.otherDivisions;     // ['Annex','Annexure','Appendix','Part']

  // Use lowerCase for the first letter so it matches both "clause 4.2" and "Clause 4.2".
  const clauseLower = primary.charAt(0).toLowerCase() + primary.slice(1);

  return [
    // "clause 4", "clause 4.2", "clause 4.2(a)" — compound decimal
    `${clauseLower} [0-9]@.[0-9]@`,
    // "clause 4(a)" — top-level with lettered limb
    `${clauseLower} [0-9]@\\([a-z]@\\)`,
    // "paragraph 1", "paragraph 1.1" (Schedule sub-units)
    `${subUnit} [0-9]@`,
    // "Schedule 1", "Schedule 2"
    `${scheduleWord} [0-9]@`,
    // "Part 1", "Part 2" (Part is in otherDivisions)
    ...(otherDivs.includes('Part') ? ['Part [0-9]@'] : []),
    // Standalone lettered limbs: "(a)", "(b)", "(c)"
    '\\([a-z]\\)',
    // Standalone romanette limbs: "(i)", "(ii)", "(iii)"
    '\\([ivxlc]@\\)',
    // Recital references: "recital (A)", "Recital (B)"
    '[Rr]ecital \\([A-Z]\\)',
  ];
}

// ---------------------------------------------------------------------------
// Highlight scan (requires Word context)
// ---------------------------------------------------------------------------

/**
 * Run the cross-reference review after any apply/promote/demote.
 *
 * Clears any prior review highlight first (scoped to our sentinel colour only),
 * then re-searches and applies the new highlight.
 *
 * Returns the updated ReviewState for the UI panel.
 */
export async function reviewCrossReferences(
  body: Word.Body,
  profile: JurisdictionProfile,
  ctx: Word.RequestContext,
): Promise<ReviewState> {
  const config = getNumberingConfig(profile);
  const highlightColour = config.reviewHighlightColour ?? DEFAULT_HIGHLIGHT;

  // 1. Clear prior review highlights (scoped — only remove our sentinel).
  await clearReviewHighlights(body, highlightColour, ctx);

  const patterns = buildReviewPatterns(profile);
  const items: ReviewItem[] = [];

  // 2. Search and highlight each pattern.
  for (const pattern of patterns) {
    let results: Word.RangeCollection;
    try {
      results = body.search(pattern, { matchWildcards: true });
      ctx.load(results, 'items');
      await ctx.sync();
    } catch {
      continue; // unsupported pattern on this host — skip
    }

    // Load text and font in a second pass; guard independently so a load
    // failure on one pattern doesn't abort the whole review.
    try {
      ctx.load(results, 'items/font,items/text');
      await ctx.sync();
    } catch {
      continue;
    }

    for (const range of results.items) {
      try {
        // highlightColour must be a Word named colour (e.g. 'Yellow'), not
        // a hex string — arbitrary hex throws "argument not valid" in Word Online.
        range.font.highlightColor = highlightColour;
        const text = range.text;
        items.push({
          text,
          context: text,
          paraId: '',       // para id not readily available here — panel uses text match
          startOffset: 0,
          endOffset: text.length,
        });
      } catch { /* skip this range */ }
    }
  }

  // Commit all highlight writes.
  try { await ctx.sync(); } catch { /* best-effort */ }

  return {
    count: items.length,
    items,
    active: true,
  };
}

// ---------------------------------------------------------------------------
// Scoped clear (removes only our review highlight colour)
// ---------------------------------------------------------------------------

/**
 * Remove the review highlight colour from all ranges that have it.
 * Does NOT touch any other highlight colour the user may have applied.
 * (spec §7.5 / §5.3 "scoped Clear" behaviour.)
 */
export async function clearReviewHighlights(
  body: Word.Body,
  highlightColour: string,
  ctx: Word.RequestContext,
): Promise<void> {
  try {
    // Search for any text with our specific sentinel highlight.
    // Word wildcard "?" matches any single character; broad search to find all.
    const all = body.search('?@', { matchWildcards: true });
    ctx.load(all, 'items/font');
    await ctx.sync();

    for (const range of all.items) {
      if (range.font.highlightColor === highlightColour) {
        range.font.highlightColor = 'None';
      }
    }
    await ctx.sync();
  } catch {
    /* best-effort clear; non-fatal */
  }
}
