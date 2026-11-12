/**
 * Feature 06 — Numbering upgrade (ADR-0023): assisted static-number removal (M11, Step 1).
 *
 * Deletes hand-typed leading clause-number prefixes over the current Word
 * selection as NATIVE TRACKED CHANGES. This is a typed-API text delete only —
 * NO OOXML `abstractNum`/`num` injection (that conversion stays cut, ADR-0023).
 *
 * The six binding constraints from ADR-0023 honoured here:
 *  1. Tracked changes mandatory — we force `changeTrackingMode = trackAll` for
 *     the batch and RESTORE the prior mode afterwards. If the host refuses to
 *     enable tracking, we perform NO deletes and report `disabled` so the driver
 *     can grey the step out.
 *  2. Batch guard — we only delete where a leading token is detected
 *     (computeRemovalSpan returns null → skip; hanging paragraphs untouched).
 *  3. Whitespace rule — the removal span (incl. trailing run) comes from the
 *     single global override (token-span.ts).
 *  5. Leading-occurrence-only — search is scoped to each paragraph and we delete
 *     `items[0]` (the lowest-offset match), never a later in-body occurrence.
 *
 * The deletion mechanism is a tracked-change-aware `range.insertText('', 'Replace')`
 * on the leading slice — the same proven primitive as cross-reference/assign-source.
 */

import { computeRemovalSpan } from './token-span.js';
import type { WhitespaceOverride } from './token-span.js';

/** Per-batch outcome of an assisted removal pass. */
export interface RemovalReport {
  /** Host refused to enable Track Changes — NO deletes were performed (constraint 1). */
  disabled: boolean;
  /** Paragraphs examined in the current selection. */
  total: number;
  /** Paragraphs whose leading prefix was deleted as a tracked change. */
  deleted: number;
  /** Paragraphs skipped because no leading token was detected (batch guard, constraint 2). */
  skipped: number;
  /**
   * Paragraphs that had a detected token but whose slice could not be located by
   * search (e.g. tab/NBSP-delimited prefixes the host search won't match). Safe
   * no-op — nothing deleted, surfaced so the driver can advise a manual edit.
   */
  notFound: number;
  /**
   * The leading tokens detected across the selection, in document order — fed to
   * the sequence-anomaly detector so duplicate/gap readings can be surfaced
   * (constraint 4: never auto-resequence). Tokens are collected whether or not the
   * delete succeeded (a notFound paragraph still carries a real number).
   */
  tokens: string[];
}

const EMPTY_REPORT: RemovalReport = {
  disabled: false, total: 0, deleted: 0, skipped: 0, notFound: 0, tokens: [],
};

/**
 * Remove leading hand-typed numbers from the paragraphs intersecting the current
 * Word selection, recorded as tracked changes. Reads selected paragraphs live from
 * `ctx.document.getSelection()` so no stale proxies cross the Word.run boundary
 * (mirrors applyToCurrentSelection).
 */
export async function removeLeadingNumbers(
  override: WhitespaceOverride,
  ctx: Word.RequestContext,
): Promise<RemovalReport> {
  // --- Constraint 1: force Track Changes, remembering the prior mode. ----------
  ctx.load(ctx.document, 'changeTrackingMode');
  await ctx.sync();
  const priorMode = ctx.document.changeTrackingMode;

  try {
    ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    await ctx.sync();
  } catch {
    // Host cannot enable tracking — perform no deletes (constraint 1).
    return { ...EMPTY_REPORT, disabled: true };
  }

  try {
    const selection = ctx.document.getSelection();
    const paragraphs = selection.paragraphs;
    ctx.load(paragraphs, 'items');
    await ctx.sync();

    const items = paragraphs.items;
    for (const paragraph of items) {
      paragraph.load('text');
    }
    await ctx.sync();

    let deleted = 0;
    let skipped = 0;
    let notFound = 0;
    const tokens: string[] = [];

    for (const paragraph of items) {
      const span = computeRemovalSpan(paragraph.text, override);
      if (!span) {
        // Constraint 2: no leading token — leave the hanging paragraph alone.
        skipped += 1;
        continue;
      }

      // Record the token in document order for anomaly surfacing (constraint 4).
      tokens.push(span.token);

      const needle = paragraph.text.slice(span.start, span.end);
      // Constraint 5: search is scoped to THIS paragraph; items[0] is the
      // leading (lowest-offset) match, never a later in-body occurrence.
      const found = paragraph.getRange().search(needle, { matchCase: true });
      found.load('items');
      await ctx.sync();

      const hit = found.items[0];
      if (!hit) {
        // Delimiter the host search can't match (tab/NBSP) — safe no-op.
        notFound += 1;
        continue;
      }

      // Tracked-change-aware delete of the leading slice.
      hit.insertText('', 'Replace');
      await ctx.sync();
      deleted += 1;
    }

    return { disabled: false, total: items.length, deleted, skipped, notFound, tokens };
  } finally {
    // Constraint 1: restore the user's prior change-tracking mode. The revisions
    // we recorded persist regardless of the restored mode.
    try {
      ctx.document.changeTrackingMode = priorMode;
      await ctx.sync();
    } catch {
      /* best-effort restore — never mask the primary result */
    }
  }
}
