/**
 * Selection plumbing for Interactive Reference Lookup (feature 04).
 *
 * Responsibilities (spec §4.1–4.3, M0):
 *  1. Register `addHandlerAsync(DocumentSelectionChanged)` once after Office.onReady.
 *  2. 400 ms trailing debounce (config lookup.debounceMs) — typing yields zero lookups.
 *  3. Read the selected span directly, or grow a bare cursor to the enclosing
 *     paragraph text capped at sentenceWindowChars (~240 chars).
 *  4. Seq-guard cancellation: every lookup is versioned; stale results are discarded.
 *     We cannot cancel the Office round-trip, but only the latest result renders.
 *
 * ESCALATE — fragile Office event / round-trip sequencing + cancellation logic.
 */

import { getLookupConfig } from '@shared/config/index.js';
import { getActiveProfile } from '@shared/config/index.js';

// ── Seq-guard ─────────────────────────────────────────────────────────────────

let _lookupSeq = 0;

/** Increment and return the new sequence number for this lookup. */
export function nextSeq(): number {
  return ++_lookupSeq;
}

/** True if `seq` has been superseded by a newer lookup. */
export function isSuperseded(seq: number): boolean {
  return seq !== _lookupSeq;
}

/** Reset seq counter (testing only). */
export function _resetSeq(): void {
  _lookupSeq = 0;
}

// ── Debounce ──────────────────────────────────────────────────────────────────

let _debounceTimer: number | null = null;
let _onSettled: (() => void) | null = null;

/** Raw event handler — always debounces then calls _onSettled. */
function onSelectionChangedRaw(): void {
  if (_debounceTimer !== null) clearTimeout(_debounceTimer);
  const cfg = getLookupConfig(getActiveProfile());
  _debounceTimer = window.setTimeout(() => {
    _debounceTimer = null;
    _onSettled?.();
  }, cfg.debounceMs);
}

// ── Registration ──────────────────────────────────────────────────────────────

export type RegistrationStatus = 'ok' | 'failed';

/**
 * Register the `DocumentSelectionChanged` handler once.
 * Returns 'ok' on success, 'failed' when the Office async callback reports an error
 * (e.g., the host does not support the event or the context is gone).
 * The caller should disable the feature and surface a graceful message on 'failed'.
 */
export function registerSelectionHandler(
  onSettled: () => void,
): Promise<RegistrationStatus> {
  _onSettled = onSettled;
  return new Promise<RegistrationStatus>((resolve) => {
    try {
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        onSelectionChangedRaw,
        (res) => {
          if (res.status !== Office.AsyncResultStatus.Succeeded) {
            resolve('failed');
          } else {
            resolve('ok');
          }
        },
      );
    } catch {
      resolve('failed');
    }
  });
}

/** Tear down: clear any pending debounce timer. */
export function clearDebounce(): void {
  if (_debounceTimer !== null) {
    clearTimeout(_debounceTimer);
    _debounceTimer = null;
  }
}

// ── Selection read ────────────────────────────────────────────────────────────

export interface SelectionWindow {
  /** The text to classify. Non-empty selection text, or the paragraph sentence window. */
  text: string;
  /** True when this came from an explicit non-empty selection (vs bare cursor expansion). */
  isExplicit: boolean;
  /**
   * For a bare cursor, the caret's character offset within `text` (the sentence
   * window). Lets the classifier resolve the defined term the cursor sits
   * *inside* (spec §4.2 / T3). Undefined for explicit selections.
   */
  cursorOffset?: number;
}

/**
 * Read the current selection inside a Word.run context.
 *
 * - Non-empty selection → use selection text directly (`isExplicit: true`).
 * - Bare cursor (empty selection) → load the containing paragraph's text, then return
 *   a window of up to `sentenceWindowChars` chars. This lets the classifier see
 *   multi-token references ("Clause", "4.2", "of", "Schedule", "2") even when only
 *   one token is under the cursor.
 *
 * Two-sync path for bare cursor; one-sync for explicit selection (fast path).
 *
 * Callers should check `isSuperseded(seq)` after each await before acting on the result.
 */
export async function readSelectionWindow(
  ctx: Word.RequestContext,
): Promise<SelectionWindow> {
  const profile = getActiveProfile();
  const { sentenceWindowChars } = getLookupConfig(profile);

  const sel = ctx.document.getSelection();
  sel.load('text, isEmpty');
  await ctx.sync();

  const selText = sel.text?.trim() ?? '';
  if (selText.length > 0) {
    return { text: selText, isExplicit: true };
  }

  // Bare cursor: expand to the containing paragraph for multi-token context.
  // Use the first paragraph only (selection always in one paragraph for bare cursor).
  const paras = sel.paragraphs;
  paras.load('items');
  await ctx.sync();

  const paraItem = paras.items[0];
  if (!paraItem) return { text: '', isExplicit: false };

  paraItem.load('text');
  // Range from paragraph start to the caret — its text length is the caret's
  // offset within the paragraph, used to find the defined term under the cursor.
  const prefix = paraItem
    .getRange(Word.RangeLocation.start)
    .expandTo(sel.getRange(Word.RangeLocation.start));
  prefix.load('text');
  await ctx.sync();

  const rawPara = paraItem.text ?? '';
  // `paraText` is left-trimmed; shift the caret offset by the trimmed prefix so
  // it stays aligned with the window we return and classify against.
  const leadingWs = rawPara.length - rawPara.replace(/^\s+/, '').length;
  const paraText = rawPara.trim();
  const cursorOffset = Math.max(0, (prefix.text ?? '').length - leadingWs);
  // Cap to sentenceWindowChars to bound classifier latency regardless of para size.
  const window = paraText.length > sentenceWindowChars
    ? paraText.slice(0, sentenceWindowChars)
    : paraText;

  return { text: window, isExplicit: false, cursorOffset };
}
