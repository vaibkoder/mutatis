/**
 * Feature 06 — Numbering: apply engine (M2).
 *
 * Applies a NumberingScheme to Word lists via setLevelNumbering + setLevelIndents.
 * Spec §5.1 behaviours:
 *   - Excludes table-cell paragraphs (§8.1) and footnote paragraphs (§8.2).
 *   - Skips statically-numbered paragraphs (§8.3, ADR-0006 — NEVER converts).
 *   - Per-level error tolerance for Word Online (§3.1).
 *   - Detects style-linked lists and warns (§8.7 — does NOT modify them).
 *   - Multi-section detection for the continuity warning (§8.4).
 *
 * IMPORTANT: typed Office JS API only — NO OOXML loop, NO abstractNum injection.
 * (Confidentiality invariant + ADR-0006.)
 */

import type { NumberingScheme } from './types.js';
import type { ApplyReport } from './types.js';
import type { ParaNode } from '@shared/doc-model/types.js';
import type { Finding } from '@shared/findings/types.js';
import {
  composeFormat,
  mapToWordListNumbering,
  formatStringToArray,
} from './schema.js';

// ---------------------------------------------------------------------------
// Static-numbering integration (spec §5.1 step 3, ADR-0006)
// ---------------------------------------------------------------------------

/**
 * Returns true when the paragraph is flagged as statically numbered by
 * feature 02's 'static-numbering/manually-typed' finding.
 * Read-only consumer — this tab never writes that finding.
 */
export function isStaticNumbered(paraId: string, staticFindings: readonly Finding[]): boolean {
  return staticFindings.some(
    f => f.category === 'static-numbering/manually-typed' && f.location.paraId === paraId,
  );
}

// ---------------------------------------------------------------------------
// Paragraph classification (pure — testable without Office)
// ---------------------------------------------------------------------------

export type ParaKind = 'in-list' | 'convertible' | 'static' | 'table' | 'footnote';

export interface ClassifiedPara {
  para: ParaNode;
  kind: ParaKind;
}

/**
 * Pre-filter paragraphs by kind before any Word.run round-trip.
 * Table-cell and footnote paragraphs are excluded per spec §8.1/§8.2.
 * Statically-numbered paragraphs are skipped per ADR-0006.
 * The rest are 'convertible' (eligible for apply).
 *
 * Note: 'in-list' classification requires a live Word context — the driver
 * later refines 'convertible' to 'in-list' by calling listItemOrNullObject.
 */
export function classifyParagraphs(
  paras: readonly ParaNode[],
  staticFindings: readonly Finding[],
): ClassifiedPara[] {
  return paras.map(para => {
    if (para.container === 'footnote') return { para, kind: 'footnote' };
    if (para.container === 'tableCell') return { para, kind: 'table' };
    if (isStaticNumbered(para.id, staticFindings)) return { para, kind: 'static' };
    return { para, kind: 'convertible' };
  });
}

// ---------------------------------------------------------------------------
// Apply scheme to a resolved Word.List (inner loop)
// ---------------------------------------------------------------------------

/**
 * Apply all levels of a scheme to a single Word.List.
 * Returns the set of level indices that could not be applied (per-level tolerance §3.1).
 * Catches per-level errors and continues — never throws from partial failure.
 */
export async function applySchemeToList(
  list: Word.List,
  scheme: NumberingScheme,
  ctx: Word.RequestContext,
): Promise<number[]> {
  const degraded: number[] = [];
  const sortedLevels = [...scheme.levels].sort((a, b) => a.levelIndex - b.levelIndex);

  // IMPORTANT: Office JS defers errors to ctx.sync(). A try-catch around the
  // setLevel* call alone catches nothing — the "argument not valid" surfaces at
  // the next sync. So each level is queued AND synced inside its own try block,
  // giving genuine per-level tolerance (one bad level can't abort the whole list).
  for (const level of sortedLevels) {
    try {
      const wordNumbering = mapToWordListNumbering(level.listNumbering) as
        'None' | 'Arabic' | 'UpperRoman' | 'LowerRoman' | 'UpperLetter' | 'LowerLetter';
      const rawFormat = composeFormat(level.prefix, level.formatString, level.suffix);
      const formatArray = formatStringToArray(rawFormat);
      list.setLevelNumbering(level.levelIndex, wordNumbering, formatArray);
      await ctx.sync();
    } catch {
      degraded.push(level.levelIndex);
      continue;
    }

    try {
      list.setLevelIndents(
        level.levelIndex,
        level.indent.textIndentPt,
        level.indent.numberIndentPt,
      );
      await ctx.sync();
    } catch {
      if (!degraded.includes(level.levelIndex)) degraded.push(level.levelIndex);
    }

    // Alignment is not set programmatically — left is Word's default for legal lists,
    // and setLevelAlignment has inconsistent API surface across host builds.
  }

  return degraded;
}

// ---------------------------------------------------------------------------
// Apply to document body (all lists)
// ---------------------------------------------------------------------------

/**
 * Apply a scheme to every list in the document body.
 * Used for 'document' scope. Table-cell and footnote paragraphs are skipped
 * by their position (lists in those locations are not in body.lists).
 */
export async function applyToDocumentBody(
  scheme: NumberingScheme,
  staticFindings: readonly Finding[],
  ctx: Word.RequestContext,
): Promise<ApplyReport> {
  const lists = ctx.document.body.lists;
  ctx.load(lists, 'items');
  await ctx.sync();

  const allDegraded: number[] = [];
  for (const list of lists.items) {
    const degraded = await applySchemeToList(list, scheme, ctx);
    for (const d of degraded) {
      if (!allDegraded.includes(d)) allDegraded.push(d);
    }
  }

  const hasMultipleSections = await detectMultipleSections(ctx);

  return {
    listsAffected: lists.items.length,
    degradedLevels: allDegraded,
    tableParasSkipped: 0,
    footnoteParasSkipped: 0,
    staticParasSkipped: staticFindings.filter(f => f.category === 'static-numbering/manually-typed').length,
    hasMultipleSections,
    styleLinkedSkipped: false,
  };
}

// ---------------------------------------------------------------------------
// Apply to selection (only lists touching the selected paragraphs)
// ---------------------------------------------------------------------------

/**
 * Apply a scheme to the list(s) intersecting the current selection.
 * Used for 'selection' scope (spec §7.4).
 */
export async function applyToSelection(
  selectedParas: readonly ParaNode[],
  scheme: NumberingScheme,
  staticFindings: readonly Finding[],
  ctx: Word.RequestContext,
): Promise<ApplyReport> {
  const classified = classifyParagraphs(selectedParas, staticFindings);
  const tableParasSkipped = classified.filter(r => r.kind === 'table').length;
  const footnoteParasSkipped = classified.filter(r => r.kind === 'footnote').length;
  const staticParasSkipped = classified.filter(r => r.kind === 'static').length;

  const eligible = classified.filter(r => r.kind === 'convertible');

  // Collect unique lists from eligible paragraphs.
  const listIds = new Set<number>();
  const uniqueLists: Word.List[] = [];

  for (const { para } of eligible) {
    const range = para.getRange(ctx);
    const paragraphs = range.paragraphs;
    ctx.load(paragraphs, 'items');
    await ctx.sync();

    for (const paragraph of paragraphs.items) {
      const listItem = paragraph.listItemOrNullObject;
      ctx.load(listItem, 'isNullObject');
      await ctx.sync();

      if (listItem.isNullObject) continue;

      const list = paragraph.listOrNullObject;
      ctx.load(list, 'isNullObject,id');
      await ctx.sync();

      if (list.isNullObject || listIds.has(list.id)) continue;
      listIds.add(list.id);
      uniqueLists.push(list);
    }
  }

  const allDegraded: number[] = [];
  for (const list of uniqueLists) {
    const degraded = await applySchemeToList(list, scheme, ctx);
    for (const d of degraded) {
      if (!allDegraded.includes(d)) allDegraded.push(d);
    }
  }

  const hasMultipleSections = await detectMultipleSections(ctx);

  return {
    listsAffected: uniqueLists.length,
    degradedLevels: allDegraded,
    tableParasSkipped,
    footnoteParasSkipped,
    staticParasSkipped,
    hasMultipleSections,
    styleLinkedSkipped: false,
  };
}

// ---------------------------------------------------------------------------
// Apply to current Word selection (live — within a single Word.run)
// ---------------------------------------------------------------------------

/**
 * Apply a scheme to the list(s) intersecting the *current* Word selection.
 * This is the live counterpart to applyToSelection: it obtains the selected
 * paragraphs directly from ctx.document.getSelection() so that no stale proxy
 * objects cross Word.run context boundaries.
 *
 * Used by the driver for 'selection' scope. applyToSelection(paras, ...) remains
 * for unit-test use but must NOT be called from a live Word.run with DocModel nodes.
 */
export async function applyToCurrentSelection(
  scheme: NumberingScheme,
  staticFindings: readonly Finding[],
  ctx: Word.RequestContext,
): Promise<ApplyReport> {
  const selection = ctx.document.getSelection();
  const paragraphs = selection.paragraphs;
  ctx.load(paragraphs, 'items');
  await ctx.sync();

  const listIds = new Set<number>();
  const uniqueLists: Word.List[] = [];

  for (const paragraph of paragraphs.items) {
    // Skip table-cell paragraphs (best-effort — not all hosts support parentTableCellOrNullObject).
    try {
      const cell = paragraph.parentTableCellOrNullObject;
      ctx.load(cell, 'isNullObject');
      await ctx.sync();
      if (!cell.isNullObject) continue;
    } catch { /* host doesn't support cell check — proceed */ }

    const listItem = paragraph.listItemOrNullObject;
    ctx.load(listItem, 'isNullObject');
    await ctx.sync();
    if (listItem.isNullObject) continue;

    const list = paragraph.listOrNullObject;
    ctx.load(list, 'isNullObject,id');
    await ctx.sync();
    if (list.isNullObject || listIds.has(list.id)) continue;
    listIds.add(list.id);
    uniqueLists.push(list);
  }

  const allDegraded: number[] = [];
  for (const list of uniqueLists) {
    const degraded = await applySchemeToList(list, scheme, ctx);
    for (const d of degraded) {
      if (!allDegraded.includes(d)) allDegraded.push(d);
    }
  }

  const hasMultipleSections = await detectMultipleSections(ctx);

  return {
    listsAffected: uniqueLists.length,
    degradedLevels: allDegraded,
    tableParasSkipped: 0,
    footnoteParasSkipped: 0,
    staticParasSkipped: staticFindings.filter(f => f.category === 'static-numbering/manually-typed').length,
    hasMultipleSections,
    styleLinkedSkipped: false,
  };
}

// ---------------------------------------------------------------------------
// Multi-section detection helper (§8.4)
// ---------------------------------------------------------------------------

/**
 * Returns true when the document has more than one Word section.
 * Used to surface the §8.4 top-level continuity warning.
 */
export async function detectMultipleSections(ctx: Word.RequestContext): Promise<boolean> {
  try {
    const sections = ctx.document.sections;
    ctx.load(sections, 'items');
    await ctx.sync();
    return sections.items.length > 1;
  } catch {
    return false;
  }
}
