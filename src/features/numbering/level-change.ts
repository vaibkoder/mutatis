/**
 * Feature 06 — Numbering: promote/demote (M3).
 *
 * Spec §5.2:
 *   - Changes listItem.level ±1, clamped to 0..8.
 *   - For non-list paragraphs adjacent to a list, uses attachToList.
 *   - Skips table-cell paragraphs (§8.1).
 *   - Provides button boundary state (disabled when all selected items are at a boundary).
 *   - Warns if Track Changes is on (caller responsibility — see warnings.ts).
 */

import type { ParaNode } from '@shared/doc-model/types.js';

export const MIN_LEVEL = 0;
export const MAX_LEVEL = 8;

// ---------------------------------------------------------------------------
// Boundary state (pure — drives UI button enable/disable)
// ---------------------------------------------------------------------------

/**
 * Returns whether promote and demote are available given a set of
 * list-level values. Both may be false if the selection is empty.
 */
export function boundaryState(levels: readonly number[]): { canPromote: boolean; canDemote: boolean } {
  if (levels.length === 0) return { canPromote: false, canDemote: false };
  const canPromote = levels.some(l => l > MIN_LEVEL);
  const canDemote = levels.some(l => l < MAX_LEVEL);
  return { canPromote, canDemote };
}

/** Clamp a level to the valid [0, 8] range. */
export function clampLevel(level: number): number {
  return Math.max(MIN_LEVEL, Math.min(MAX_LEVEL, level));
}

// ---------------------------------------------------------------------------
// Promote/demote (requires Word context)
// ---------------------------------------------------------------------------

/**
 * Promote (delta = -1) or demote (delta = +1) the selected paragraphs.
 * Spec §5.2 algorithm:
 *   - Table-cell paragraphs: skip.
 *   - Paragraphs already in a list: write listItem.level = clamped(level ± delta).
 *     Word auto-recalculates surrounding numbers.
 *   - Non-list paragraphs adjacent to a list: attachToList at the target level.
 *   - Non-list paragraphs not adjacent to a list: skip (nothing to promote/demote).
 *
 * @returns the number of paragraphs whose level was changed.
 */
export async function changeLevel(
  paras: readonly ParaNode[],
  delta: -1 | 1,
  ctx: Word.RequestContext,
): Promise<number> {
  let changed = 0;

  for (const para of paras) {
    // Skip table-cell paragraphs (§8.1).
    if (para.container === 'tableCell') continue;

    const range = para.getRange(ctx);
    const paragraphs = range.paragraphs;
    ctx.load(paragraphs, 'items');
    await ctx.sync();

    for (const paragraph of paragraphs.items) {
      const listItem = paragraph.listItemOrNullObject;
      ctx.load(listItem, 'isNullObject,level');
      await ctx.sync();

      if (!listItem.isNullObject) {
        const newLevel = clampLevel(listItem.level + delta);
        if (newLevel !== listItem.level) {
          listItem.level = newLevel;
          await ctx.sync();
          changed++;
        }
      } else {
        // Non-list paragraph — try to attach to an adjacent list.
        const adjacentList = await findAdjacentList(paragraph, ctx);
        if (adjacentList) {
          const inferredLevel = clampLevel(0 + delta);
          paragraph.attachToList(adjacentList.id, inferredLevel);
          await ctx.sync();
          changed++;
        }
        // If no adjacent list, skip silently (nothing to promote/demote).
      }
    }
  }

  return changed;
}

// ---------------------------------------------------------------------------
// Promote/demote on current Word selection (live — within a single Word.run)
// ---------------------------------------------------------------------------

/**
 * Promote (delta = -1) or demote (delta = +1) paragraphs in the *current*
 * Word selection. Uses ctx.document.getSelection() so no stale proxy objects
 * cross Word.run boundaries. changeLevel(paras, ...) remains for unit tests.
 */
export async function changeLevelOnSelection(
  delta: -1 | 1,
  ctx: Word.RequestContext,
): Promise<number> {
  const selection = ctx.document.getSelection();
  const paragraphs = selection.paragraphs;
  ctx.load(paragraphs, 'items');
  await ctx.sync();

  let changed = 0;

  for (const paragraph of paragraphs.items) {
    // Skip table-cell paragraphs (best-effort).
    try {
      const cell = paragraph.parentTableCellOrNullObject;
      ctx.load(cell, 'isNullObject');
      await ctx.sync();
      if (!cell.isNullObject) continue;
    } catch { /* host doesn't support cell check — proceed */ }

    const listItem = paragraph.listItemOrNullObject;
    ctx.load(listItem, 'isNullObject,level');
    await ctx.sync();

    if (!listItem.isNullObject) {
      const newLevel = clampLevel(listItem.level + delta);
      if (newLevel !== listItem.level) {
        listItem.level = newLevel;
        await ctx.sync();
        changed++;
      }
    } else {
      // Non-list paragraph — try to attach to an adjacent list.
      const adjacentList = await findAdjacentList(paragraph, ctx);
      if (adjacentList) {
        const inferredLevel = clampLevel(0 + delta);
        paragraph.attachToList(adjacentList.id, inferredLevel);
        await ctx.sync();
        changed++;
      }
    }
  }

  return changed;
}

// ---------------------------------------------------------------------------
// Adjacent-list detection (helper)
// ---------------------------------------------------------------------------

/**
 * Find a list that the paragraph is adjacent to (immediately before or after).
 * Used to attach a non-list paragraph into a nearby list's scheme.
 * Returns null if no adjacent list is found.
 */
async function findAdjacentList(
  paragraph: Word.Paragraph,
  ctx: Word.RequestContext,
): Promise<Word.List | null> {
  try {
    // Check the paragraph immediately after.
    const next = paragraph.getNextOrNullObject();
    ctx.load(next, 'isNullObject');
    await ctx.sync();
    if (!next.isNullObject) {
      const nextListItem = next.listItemOrNullObject;
      ctx.load(nextListItem, 'isNullObject');
      await ctx.sync();
      if (!nextListItem.isNullObject) {
        const list = next.listOrNullObject;
        ctx.load(list, 'isNullObject,id');
        await ctx.sync();
        if (!list.isNullObject) return list;
      }
    }

    // Fall back to the paragraph immediately before.
    const prev = paragraph.getPreviousOrNullObject();
    ctx.load(prev, 'isNullObject');
    await ctx.sync();
    if (!prev.isNullObject) {
      const prevListItem = prev.listItemOrNullObject;
      ctx.load(prevListItem, 'isNullObject');
      await ctx.sync();
      if (!prevListItem.isNullObject) {
        const list = prev.listOrNullObject;
        ctx.load(list, 'isNullObject,id');
        await ctx.sync();
        if (!list.isNullObject) return list;
      }
    }
  } catch {
    /* non-fatal — adjacent detection is best-effort */
  }

  return null;
}
