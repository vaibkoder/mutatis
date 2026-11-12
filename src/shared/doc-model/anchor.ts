import type { ParaNode } from './types.js';

/** A character-offset anchor into a ParaNode's normalised text. */
export interface TextAnchor {
  /** Matches `ParaNode.id`. */
  paragraphId: string;
  /** Inclusive start offset into `ParaNode.text`. */
  startOffset: number;
  /** Exclusive end offset into `ParaNode.text`. */
  endOffset: number;
}

/** Word's document search is capped at 255 characters per query. */
const SEARCH_MAX = 255;

/**
 * Resolve a `TextAnchor` to a live, selectable `Word.Range` in the **caller's**
 * request context.
 *
 * Office.js binds object proxies to the `Word.run` that created them, so a
 * paragraph captured while the snapshot was loaded cannot be reused for
 * navigation in a later `Word.run` (it throws `InvalidObjectPath`). This
 * resolver therefore never touches the snapshot's captured proxies: it issues a
 * fresh document-wide `body.search()` for the exact offending substring and
 * selects the correct occurrence by its **global rank** — the number of
 * identical substrings that appear earlier in document order, computed from the
 * in-memory snapshot text. That count matches Word's document-order search
 * results, so `items[rank]` is the intended hit.
 *
 * The search is queued but its results only materialise after a `ctx.sync()`.
 * The returned value is a Proxy whose `select`/`scrollIntoView` read the loaded
 * hit lazily, so the navigation caller must:
 *   Word.run(async ctx => {
 *     const range = resolveAnchor(ctx, paras, anchor);
 *     await ctx.sync();   // materialise the search hits
 *     range.select();
 *     await ctx.sync();
 *   });
 *
 * AC-6: this remains the ONLY place navigation calls `search()`; extraction never does.
 */
export function resolveAnchor(
  ctx: Word.RequestContext,
  paras: ReadonlyArray<ParaNode>,
  anchor: TextAnchor,
): Word.Range {
  const idx = paras.findIndex(p => p.id === anchor.paragraphId);
  if (idx === -1) {
    throw new Error(`resolveAnchor: paragraph '${anchor.paragraphId}' not found in snapshot`);
  }
  const para = paras[idx]!;

  // Inert fallback target for the Proxy: a whole-body range, valid in any
  // context. Only used to forward non-navigation property access; select and
  // scrollIntoView are always overridden to delegate to the resolved hit.
  const bodyRange = ctx.document.body.getRange();

  // The exact offending substring. A zero-length anchor degrades to navigating
  // to the paragraph itself (searching on its leading text).
  const exact = para.text.slice(anchor.startOffset, anchor.endOffset);
  const needle = (exact.length > 0 ? exact : para.text).slice(0, SEARCH_MAX);
  if (!needle) return bodyRange; // empty paragraph — nothing to search for.

  const rank = exact.length > 0
    ? occurrenceRankExact(paras, idx, needle, anchor.startOffset)
    : occurrenceRankParagraph(paras, idx, needle);

  const hits = ctx.document.body.search(needle, { matchCase: true, matchWholeWord: false });
  hits.load('items');

  return makeDeferredHit(bodyRange, hits, rank);
}

/**
 * Global rank of an exact substring: identical substrings in earlier
 * (body-searchable) paragraphs, plus earlier occurrences in this paragraph
 * before `startOffset`. Footnote paragraphs are skipped because `body.search`
 * does not cover footnote text.
 */
function occurrenceRankExact(
  paras: ReadonlyArray<ParaNode>,
  idx: number,
  needle: string,
  startOffset: number,
): number {
  let rank = 0;
  for (let i = 0; i < idx; i++) {
    const p = paras[i]!;
    if (p.container === 'footnote') continue;
    rank += countOccurrences(p.text, needle);
  }
  rank += countOccurrencesBefore(paras[idx]!.text, needle, startOffset);
  return rank;
}

/** Global rank of a paragraph among earlier paragraphs whose leading text matches. */
function occurrenceRankParagraph(
  paras: ReadonlyArray<ParaNode>,
  idx: number,
  needle: string,
): number {
  let rank = 0;
  for (let i = 0; i < idx; i++) {
    const p = paras[i]!;
    if (p.container === 'footnote') continue;
    if (p.text.slice(0, SEARCH_MAX) === needle) rank++;
  }
  return rank;
}

/** Count all non-overlapping occurrences of `needle` in `haystack`. */
function countOccurrences(haystack: string, needle: string): number {
  if (!needle) return 0;
  let count = 0;
  let pos = 0;
  for (;;) {
    const i = haystack.indexOf(needle, pos);
    if (i === -1) break;
    count++;
    pos = i + needle.length;
  }
  return count;
}

/**
 * Count how many non-overlapping occurrences of `needle` start in
 * `haystack` strictly before `beforeIndex`.
 */
function countOccurrencesBefore(haystack: string, needle: string, beforeIndex: number): number {
  let count = 0;
  let pos = 0;
  while (pos < beforeIndex) {
    const idx = haystack.indexOf(needle, pos);
    if (idx === -1 || idx >= beforeIndex) break;
    count++;
    pos = idx + needle.length;
  }
  return count;
}

/**
 * Methods that must operate on the *resolved* span (the matched search hit),
 * not the inert fallback range: navigation and in-place text replacement.
 */
const HIT_METHODS = new Set(['select', 'scrollIntoView', 'insertText']);

/**
 * Build a range proxy whose hit-scoped methods (see HIT_METHODS) delegate to the
 * resolved search hit once the collection has loaded (after the caller's
 * `ctx.sync()`). If the exact text is no longer present (document edited since
 * the snapshot), those methods become no-ops rather than acting on the wrong
 * range. All other access forwards to the inert fallback.
 */
function makeDeferredHit(
  fallback: Word.Range,
  hits: Word.RangeCollection,
  rank: number,
): Word.Range {
  return new Proxy(fallback, {
    get(target, prop) {
      if (typeof prop === 'string' && HIT_METHODS.has(prop)) {
        return (...args: unknown[]) => {
          const items: Word.Range[] = (hits as unknown as { items?: Word.Range[] }).items ?? [];
          const resolved = items[rank];
          if (!resolved) return undefined; // text not found after edit — no-op.
          return (resolved as unknown as Record<string, (...a: unknown[]) => unknown>)[prop]?.(...args);
        };
      }
      const val = (target as unknown as Record<string, unknown>)[prop as string];
      return typeof val === 'function' ? val.bind(target) : val;
    },
  }) as Word.Range;
}
