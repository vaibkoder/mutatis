/**
 * Assign-source workflow (spec §8.3).
 *
 * v1: static text replacement only via range.insertText(label, 'Replace').
 * Seam for v2 refField strategy is present but not implemented (spec §8.3 §11 v2 seam).
 *
 * The replacement label uses the CONFIGURED unit noun from the Section Index
 * entry's displayLabel — NEVER hardcodes "Section" or "Exhibit" (spec §7 hard rule).
 */

import type { ParaNode } from '@shared/doc-model/types.js';
import type { SectionNode } from '@shared/section-index/types.js';
import type { TextAnchor } from '@shared/doc-model/anchor.js';
import { resolveAnchor } from '@shared/doc-model/index.js';

/** v1 only. v2 seam: 'refField' would insert a live REF field + bookmark. */
export type ReplaceStrategy = 'text'; // | 'refField'  ← v2 seam, not implemented

export interface AssignSourceOptions {
  replaceStrategy?: ReplaceStrategy;
}

/**
 * Apply an Assign-source selection: replace the matched text range with the
 * selected target's displayLabel.
 *
 * Must be called inside a Word.run block (or provide your own ctx).
 *
 * The replacement is a tracked-change-aware `insertText('Replace')` — honours
 * whatever tracked-changes mode the document is currently in.
 *
 * @param ctx     Word.RequestContext (live)
 * @param paras   Snapshot paragraphs (for resolveAnchor)
 * @param anchor  TextAnchor of the matched reference span
 * @param target  The chosen SectionNode from the picker
 * @param opts    Strategy options (v1: 'text' only)
 */
export async function applyAssignSource(
  ctx: Word.RequestContext,
  paras: ReadonlyArray<ParaNode>,
  anchor: TextAnchor,
  target: SectionNode,
  opts: AssignSourceOptions = {},
): Promise<void> {
  const strategy = opts.replaceStrategy ?? 'text';

  if (strategy === 'text') {
    const range = resolveAnchor(ctx, paras, anchor);
    // First sync materialises the deferred search hit; resolveAnchor returns a
    // proxy that only resolves to the matched span after the collection loads.
    await ctx.sync();
    // Use the configured displayLabel (e.g. "Clause 4.3") — never "Section"/"Exhibit".
    range.insertText(target.displayLabel, 'Replace');
    await ctx.sync();
  }
  // v2 seam: else if (strategy === 'refField') { ... }
}

/**
 * Filter SectionIndex nodes to those compatible with a given unit type.
 * Used to pre-filter the picker's candidate list.
 */
export function filterCompatibleTargets(
  nodes: ReadonlyArray<SectionNode>,
  unitType: string,
): SectionNode[] {
  // If unit type is 'clause', prefer clauses first; schedules for 'schedule'; etc.
  const preferred = nodes.filter(n => n.kind === unitType);
  const others = nodes.filter(n => n.kind !== unitType);
  return [...preferred, ...others];
}
