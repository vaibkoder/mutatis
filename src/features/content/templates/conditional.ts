/**
 * Conditional show/hide — C5.
 *
 * ESCALATE: this module touches Office JS getOoxml() / insertText() on content
 * controls — the most fragile interaction in this feature.
 *
 * Non-destructive hide (default):
 *   cache control.getOoxml() into docSettings.hiddenSections[tag]
 *   → insertText('', Replace) clearing the control (keeping the CC shell)
 *   → on re-show: load cache, insertOoxml(cached, Replace) to restore.
 *
 * Destructive hide (explicit confirm only):
 *   control.delete(keepContent = false)   ← irreversible in this session.
 *
 * Office JS fidelity caveats (reason this task is ESCALATE):
 *   1. getOoxml() returns the OOXML of the control's *content*, not the whole
 *      document.  The returned string can be empty if the control is already
 *      cleared.  We guard against caching empty strings.
 *   2. insertOoxml into a content control replaces existing content but
 *      preserves the control shell — tested pattern for restore.
 *   3. content-control-in-table: insertText with embedded newlines or
 *      insertOoxml with block-level content into a PlainText-in-table CC can
 *      corrupt the cell.  We warn at authoring time (C8 edge case) but cannot
 *      fully prevent it here.
 *   4. After destructive delete the CC is gone; we cannot auto-restore.
 *
 * HOST-ONLY verification needed:
 *   - getOoxml() round-trip fidelity on Word Online vs Desktop.
 *   - insertOoxml into a CC shell after insertText('', Replace) actually
 *     restores the original layout (tables, lists, etc.) — unit tests cover
 *     the logic but live verification is required on both platforms.
 */

import { wordRun } from '@shared/office/index.js';
import { getStorageProvider } from '@shared/storage/index.js';
import type {
  ShowWhenCondition, ConditionalSection, AnswersMap, HiddenSectionsCache,
} from './types.js';
import { HIDDEN_SECTIONS_KEY } from './types.js';

// ---------------------------------------------------------------------------
// C5.1 — Condition evaluation
// ---------------------------------------------------------------------------

/**
 * Evaluate a showWhen condition against the current answers map.
 * Returns true if the section should be shown.
 */
export function evaluateCondition(
  condition: ShowWhenCondition,
  answers: AnswersMap,
): boolean {
  if ('and' in condition) {
    return condition.and.every(c => evaluateCondition(c, answers));
  }
  if ('or' in condition) {
    return condition.or.some(c => evaluateCondition(c, answers));
  }
  if ('equals' in condition) {
    const actual = answers[condition.var];
    // Coerce booleans from string form for checkbox answers.
    if (typeof condition.equals === 'boolean') {
      const coerced =
        actual === true || actual === 'true' || actual === '1' ? true :
        actual === false || actual === 'false' || actual === '0' ? false :
        actual;
      return coerced === condition.equals;
    }
    return String(actual) === String(condition.equals);
  }
  if ('notEquals' in condition) {
    const actual = answers[condition.var];
    if (typeof condition.notEquals === 'boolean') {
      const coerced =
        actual === true || actual === 'true' || actual === '1' ? true :
        actual === false || actual === 'false' || actual === '0' ? false :
        actual;
      return coerced !== condition.notEquals;
    }
    return String(actual) !== String(condition.notEquals);
  }
  return true;
}

// ---------------------------------------------------------------------------
// C5.2 — Hidden sections cache (docSettings)
// ---------------------------------------------------------------------------

async function loadHiddenCache(): Promise<HiddenSectionsCache> {
  const store = getStorageProvider().docSettings;
  const cached = await store.get(HIDDEN_SECTIONS_KEY) as HiddenSectionsCache | undefined;
  return cached && typeof cached === 'object' ? cached : {};
}

async function saveHiddenCache(cache: HiddenSectionsCache): Promise<void> {
  const store = getStorageProvider().docSettings;
  await store.set(HIDDEN_SECTIONS_KEY, cache as unknown);
}

// ---------------------------------------------------------------------------
// C5.3 — Non-destructive hide
// ---------------------------------------------------------------------------

export interface HideResult {
  ok: boolean;
  error?: string;
  /** True if the section was already hidden (idempotent). */
  alreadyHidden?: boolean;
}

/**
 * Hide a conditional section non-destructively:
 *   1. getOoxml() on the content control → cache in docSettings.hiddenSections[tag]
 *   2. insertText('', Replace) to clear content (CC shell stays).
 *
 * HOST-ONLY: Verify getOoxml() / insertOoxml round-trip fidelity on target platforms.
 */
export async function hideConditionalNonDestructive(tag: string): Promise<HideResult> {
  const cache = await loadHiddenCache();

  // If already cached, the section is already hidden — idempotent.
  if (cache[tag]) {
    return { ok: true, alreadyHidden: true };
  }

  try {
    await wordRun(async ctx => {
      const controls = ctx.document.contentControls;
      controls.load('tag');
      await ctx.sync();

      const cc = controls.items.find(c => c.tag === tag);
      if (!cc) {
        throw new Error(`Content control with tag "${tag}" not found in document.`);
      }

      // Load OOXML of the control content.
      const ooxml = cc.getOoxml();
      await ctx.sync();

      const ooxmlValue = ooxml.value;

      if (ooxmlValue && ooxmlValue.trim()) {
        cache[tag] = ooxmlValue;
      }
      // If OOXML is empty, the content is already blank — still clear for safety.

      // Clear content, keep the CC shell as a placeholder.
      cc.insertText('', 'Replace');
      await ctx.sync();
    });

    await saveHiddenCache(cache);
    return { ok: true };
  } catch (err) {
    return {
      ok: false,
      error: `Failed to hide section "${tag}": ${(err as Error).message ?? String(err)}`,
    };
  }
}

// ---------------------------------------------------------------------------
// C5.4 — Non-destructive restore
// ---------------------------------------------------------------------------

/**
 * Restore a previously hidden section from the OOXML cache.
 * If no cache entry exists, the section is assumed not to have been hidden.
 *
 * HOST-ONLY: insertOoxml into a CC shell after insertText('', Replace)
 * must restore original layout — live verification required.
 */
export async function restoreConditional(tag: string): Promise<HideResult> {
  const cache = await loadHiddenCache();
  const cachedOoxml = cache[tag];

  if (!cachedOoxml) {
    // Section was not hidden non-destructively — nothing to restore.
    return { ok: true, alreadyHidden: false };
  }

  try {
    await wordRun(async ctx => {
      const controls = ctx.document.contentControls;
      controls.load('tag');
      await ctx.sync();

      const cc = controls.items.find(c => c.tag === tag);
      if (!cc) {
        throw new Error(`Content control with tag "${tag}" not found for restore.`);
      }

      // Restore OOXML content.
      cc.insertOoxml(cachedOoxml, 'Replace');
      await ctx.sync();
    });

    // Remove from cache after successful restore.
    delete cache[tag];
    await saveHiddenCache(cache);

    return { ok: true };
  } catch (err) {
    return {
      ok: false,
      error: `Failed to restore section "${tag}": ${(err as Error).message ?? String(err)}`,
    };
  }
}

// ---------------------------------------------------------------------------
// C5.5 — Destructive delete (explicit confirm required — irreversible)
// ---------------------------------------------------------------------------

export interface DestructiveDeleteResult {
  ok: boolean;
  error?: string;
}

/**
 * Irreversibly delete a conditional content control and its content.
 * MUST only be called after the user has explicitly confirmed the modal
 * warning that this action cannot be undone within the fill session.
 */
export async function deleteConditionalDestructive(tag: string): Promise<DestructiveDeleteResult> {
  try {
    await wordRun(async ctx => {
      const controls = ctx.document.contentControls;
      controls.load('tag');
      await ctx.sync();

      const cc = controls.items.find(c => c.tag === tag);
      if (!cc) {
        // Already deleted — treat as success.
        return;
      }
      // keepContent = false → remove both the CC shell and its content.
      cc.delete(false);
      await ctx.sync();
    });

    // Remove from hidden cache if present (won't be restorable anyway).
    const cache = await loadHiddenCache();
    if (cache[tag]) {
      delete cache[tag];
      await saveHiddenCache(cache);
    }

    return { ok: true };
  } catch (err) {
    return {
      ok: false,
      error: `Failed to delete section "${tag}": ${(err as Error).message ?? String(err)}`,
    };
  }
}

// ---------------------------------------------------------------------------
// C5.6 — Bulk conditional evaluation and apply
// ---------------------------------------------------------------------------

export interface ConditionalApplyResult {
  shown: string[];
  hidden: string[];
  errors: Array<{ tag: string; error: string }>;
}

/**
 * Evaluate all conditional sections against the current answers and
 * hide or restore each one accordingly.
 */
export async function applyConditionals(
  conditionals: ConditionalSection[],
  answers: AnswersMap,
): Promise<ConditionalApplyResult> {
  const shown: string[] = [];
  const hidden: string[] = [];
  const errors: Array<{ tag: string; error: string }> = [];

  for (const cond of conditionals) {
    const shouldShow = evaluateCondition(cond.showWhen, answers);

    if (shouldShow) {
      const result = await restoreConditional(cond.tag);
      if (result.ok) {
        shown.push(cond.tag);
      } else if (result.error) {
        errors.push({ tag: cond.tag, error: result.error });
      }
    } else {
      // Non-destructive hide only (destructive requires explicit UI confirmation).
      const result = await hideConditionalNonDestructive(cond.tag);
      if (result.ok) {
        hidden.push(cond.tag);
      } else if (result.error) {
        errors.push({ tag: cond.tag, error: result.error });
      }
    }
  }

  return { shown, hidden, errors };
}
