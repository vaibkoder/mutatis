/**
 * Fixed-slot repeating fill — C6.
 *
 * Fills rows 1..K into pre-placed slot content controls, hides K+1..N
 * non-destructively.  K > N is an honest cap error — no cloning ever.
 *
 * Spec §C.4.3 (Repeating Section Content Control / OOXML
 * cloning are explicitly cut — ADR-0005, ARCHITECTURE.md row 15).
 */

import { wordRun } from '@shared/office/index.js';
import { hideConditionalNonDestructive, restoreConditional } from './conditional.js';
import type { RepeatingGroup } from './types.js';
import { SLOT_COUNT_MAX } from './types.js';
import { formatValueForInsert } from './format.js';

// ---------------------------------------------------------------------------
// Row-to-slot fill
// ---------------------------------------------------------------------------

export interface RepeatingFillRow {
  /** Values keyed by SlotTagPattern.field. */
  [field: string]: string;
}

export interface RepeatingFillResult {
  ok: boolean;
  /** The cap error message when K > N. */
  capError?: string;
  /** Tags successfully written. */
  written: string[];
  /** Tags not found in the document (drift). */
  notFound: string[];
  /** Slots hidden (K+1..N). */
  hiddenSlots: string[];
  /** Any errors during hide. */
  errors: string[];
}

/**
 * Fill a fixed-slot repeating group.
 *
 * 1. If rows.length > group.slotCount → cap error, stop (no cloning).
 * 2. Fill rows 1..K into the tagged content controls.
 * 3. Hide slots K+1..N non-destructively (re-show on re-fill).
 */
export async function fillRepeatingGroup(
  group: RepeatingGroup,
  rows: RepeatingFillRow[],
): Promise<RepeatingFillResult> {
  const K = rows.length;
  const N = group.slotCount;

  if (K > N) {
    const noun = N === 1 ? 'slot' : 'slots';
    return {
      ok: false,
      capError:
        `This template supports up to ${N} ${group.group.toLowerCase()} ${noun} ` +
        `(fixed-slot cap: max ${SLOT_COUNT_MAX}). ` +
        `You have provided ${K} rows. Please reduce the list or use a template with more slots. ` +
        `No additional rows have been written.`,
      written: [],
      notFound: [],
      hiddenSlots: [],
      errors: [],
    };
  }

  const written: string[] = [];
  const notFound: string[] = [];

  // Step 1: Restore ALL slots (1..N), not just the soon-to-be-hidden tail.
  //
  // A slot that was hidden+cached on a previous fill must have its
  // hiddenSections cache entry cleared before it is re-filled — otherwise the
  // cache wrongly marks a now-visible, freshly-filled slot as "hidden", and a
  // later hide/restore corrupts it (stale-cache bug: re-push 2→4→2 left
  // DIRECTOR_3/4 cached while visibly filled). restoreConditional is a no-op
  // (no Office round-trip) for any slot with no cache entry, so restoring the
  // active head 1..K is cheap and keeps the cache consistent with the document.
  for (let i = 1; i <= N; i++) {
    for (const slotTag of group.slotTags) {
      const tag = slotTag.tagPattern.replace('{i}', String(i));
      await restoreConditional(tag);
    }
  }

  // Step 2: Write rows 1..K in a single Word.run with one sync per chunk.
  // Each slot may have multiple fields — build the full tag→value map first,
  // then write in one pass.
  const fills: Array<{ tag: string; value: string }> = [];

  for (let i = 1; i <= K; i++) {
    const row = rows[i - 1] ?? {};
    for (const slotTag of group.slotTags) {
      const tag = slotTag.tagPattern.replace('{i}', String(i));
      const rawValue = row[slotTag.field] ?? '';
      fills.push({ tag, value: formatValueForInsert('text', rawValue) });
    }
  }

  if (fills.length > 0) {
    await wordRun(async ctx => {
      const controls = ctx.document.contentControls;
      controls.load('tag');
      await ctx.sync();

      const tagToControl = new Map<string, Word.ContentControl>();
      for (const cc of controls.items) {
        if (cc.tag) tagToControl.set(cc.tag, cc);
      }

      for (const { tag, value } of fills) {
        const cc = tagToControl.get(tag);
        if (!cc) {
          notFound.push(tag);
          continue;
        }
        cc.insertText(value, 'Replace');
        written.push(tag);
      }

      await ctx.sync();
    });
  }

  // Step 3: Hide unused slots K+1..N non-destructively.
  const hiddenSlots: string[] = [];
  const errors: string[] = [];

  if (group.hideUnusedSlots) {
    for (let i = K + 1; i <= N; i++) {
      for (const slotTag of group.slotTags) {
        const tag = slotTag.tagPattern.replace('{i}', String(i));
        const result = await hideConditionalNonDestructive(tag);
        if (result.ok) {
          hiddenSlots.push(tag);
        } else if (result.error) {
          errors.push(result.error);
        }
      }
    }
  }

  return { ok: true, written, notFound, hiddenSlots, errors };
}
