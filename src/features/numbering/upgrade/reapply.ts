/**
 * Feature 06 — Numbering upgrade (ADR-0023): human-driven re-apply (M11, Step 2).
 *
 * After Step-1 removal strips hand-typed prefixes, the cleared paragraphs belong
 * to NO list. This engine lets the user turn a selection of cleared paragraphs
 * into a real Word list: it starts a fresh list on the first eligible paragraph,
 * attaches the rest, applies the chosen scheme via the SHIPPED apply engine, and
 * runs the §5.3 cross-reference review.
 *
 * Reuse before build: the list-formatting and the
 * cross-reference review are the exact same primitives the normal Apply path uses
 * (applySchemeToList + reviewCrossReferences) — no new numbering logic. The only
 * new step is materialising a list from previously-unlisted paragraphs, which the
 * shipped apply path can't do because it only enumerates EXISTING lists.
 *
 * Typed Office JS API only — NO OOXML injection (Confidentiality + ADR-0023).
 */

import type { NumberingScheme } from '../types.js';
import type { ApplyReport, ReviewState } from '../types.js';
import type { JurisdictionProfile } from '@shared/config/types.js';
import { applySchemeToList, detectMultipleSections } from '../apply.js';
import { reviewCrossReferences } from '../review.js';

export interface ReapplyResult {
  report: ApplyReport;
  review: ReviewState;
  /** Paragraphs that joined the new list. */
  listed: number;
}

/**
 * Build a new list from the eligible paragraphs in the current Word selection,
 * apply the scheme, and run the cross-reference review.
 *
 * Reads selected paragraphs live from ctx.document.getSelection() (no stale
 * DocModel proxies crossing the Word.run boundary — mirrors applyToCurrentSelection).
 *
 * Returns null when no eligible paragraph is found (nothing to list).
 */
export async function startListAndApplyOnSelection(
  scheme: NumberingScheme,
  profile: JurisdictionProfile,
  ctx: Word.RequestContext,
): Promise<ReapplyResult | null> {
  const selection = ctx.document.getSelection();
  const paragraphs = selection.paragraphs;
  ctx.load(paragraphs, 'items');
  await ctx.sync();

  // --- Filter to eligible paragraphs (skip table cells and already-listed). ----
  // Guard parity with the shipped live apply path: applyToCurrentSelection()
  // guards table cells + already-listed paragraphs only. Footnote exclusion
  // (constraint 6 / §8.2) lives in the DocModel-prefilter path (classifyParagraphs);
  // there is no reliable per-paragraph footnote check on a live Word.Paragraph at
  // the WordApi 1.5 floor, and footnote bodies don't host clause-numbered lists in
  // practice — so we match the live path rather than invent an unreliable guard.
  const eligible: Word.Paragraph[] = [];
  let tableParasSkipped = 0;

  for (const paragraph of paragraphs.items) {
    // Skip table-cell paragraphs (§8.1, best-effort — not all hosts support it).
    try {
      const cell = paragraph.parentTableCellOrNullObject;
      ctx.load(cell, 'isNullObject');
      await ctx.sync();
      if (!cell.isNullObject) { tableParasSkipped += 1; continue; }
    } catch { /* host doesn't support cell check — proceed */ }

    // Skip paragraphs that already belong to a list — Step 2 only lists CLEARED
    // (unlisted) paragraphs; never disturb an existing list item here.
    const listItem = paragraph.listItemOrNullObject;
    ctx.load(listItem, 'isNullObject');
    await ctx.sync();
    if (!listItem.isNullObject) continue;

    eligible.push(paragraph);
  }

  if (eligible.length === 0) return null;

  // --- Materialise a single new list from the eligible paragraphs. -------------
  const [first, ...rest] = eligible;
  const list = first!.startNewList();
  ctx.load(list, 'id');
  await ctx.sync();

  for (const paragraph of rest) {
    // Attach at the top level; the user re-levels afterwards via promote/demote
    // (constraint 4 — we never auto-resequence or infer nesting depth).
    paragraph.attachToList(list.id, 0);
  }
  await ctx.sync();

  // --- Apply the scheme via the shipped engine (per-level tolerant, §3.1). -----
  const degradedLevels = await applySchemeToList(list, scheme, ctx);
  const hasMultipleSections = await detectMultipleSections(ctx);

  const report: ApplyReport = {
    listsAffected: 1,
    degradedLevels,
    tableParasSkipped,
    footnoteParasSkipped: 0,
    staticParasSkipped: 0,
    hasMultipleSections,
    styleLinkedSkipped: false,
  };

  // --- §5.3 cross-reference review (same primitive as normal Apply). -----------
  const review = await reviewCrossReferences(ctx.document.body, profile, ctx);

  return { report, review, listed: eligible.length };
}
