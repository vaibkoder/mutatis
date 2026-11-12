/**
 * Fuzzy similarity fallback for the PDF path.
 *
 * Heuristic geometry reconstruction can split the SAME wording into slightly
 * different paragraph boundaries across two PDFs, so a paragraph that exact-hash
 * anchoring can't pair surfaces as a spurious delete + insert. This pass re-pairs
 * such near-identical, structurally-unanchored delete/insert ops into a single
 * `replace` (the subsequent word-diff stage then highlights the real change),
 * using a high-threshold normalised-Levenshtein similarity. No ML.
 *
 * Invariants:
 *  - NEVER suppresses (ADR-0008) — it only RE-TYPES del+ins into a replace; the
 *    change stays fully visible, just better presented. Nothing is dropped.
 *  - Conservative — only pairs above {@link FUZZY_SIMILARITY_THRESHOLD}; exact
 *    matches are already handled (equal/move), and dissimilar ops are left as
 *    delete/insert (we never guess).
 *  - Every fuzzy match is flagged `fuzzyAnchored` so the UI badges it low-confidence.
 *
 * This realises the `SIMILARITY_THRESHOLD` seam left in diff/move.ts, scoped to
 * the PDF path only — Word-vs-Word remains verbatim-hash moves.
 */
import type { DiffOp, ParaRecord } from '../model.js';

/** Minimum normalised similarity to treat a delete/insert pair as a replace. */
export const FUZZY_SIMILARITY_THRESHOLD = 0.85;

/** Strings longer than this are truncated for the (O(n·m)) similarity metric. */
const SIMILARITY_LEN_CAP = 600;

/**
 * Re-pair near-identical unanchored delete/insert ops into replaces. Returns a
 * new ops array (paired inserts removed; their deletes retyped to replace).
 */
export function fuzzyAnchorPdf(
  ops: DiffOp[],
  original: ParaRecord[],
  revised: ParaRecord[],
): DiffOp[] {
  const deletes = ops.filter((o) => o.type === 'delete' && o.origRange !== null);
  const inserts = ops.filter((o) => o.type === 'insert' && o.revRange !== null);
  if (deletes.length === 0 || inserts.length === 0) return ops;

  // Score every candidate pair; pair greedily by descending similarity.
  const candidates: Array<{ del: DiffOp; ins: DiffOp; sim: number }> = [];
  for (const del of deletes) {
    const dText = original[del.origRange!.start]?.normText ?? '';
    if (dText === '') continue;
    for (const ins of inserts) {
      const iText = revised[ins.revRange!.start]?.normText ?? '';
      if (iText === '') continue;
      const sim = similarity(dText, iText);
      if (sim >= FUZZY_SIMILARITY_THRESHOLD && sim < 1) {
        candidates.push({ del, ins, sim });
      }
    }
  }
  candidates.sort((a, b) => b.sim - a.sim);

  const usedDel = new Set<DiffOp>();
  const usedIns = new Set<DiffOp>();
  const removeIns = new Set<DiffOp>();
  for (const { del, ins } of candidates) {
    if (usedDel.has(del) || usedIns.has(ins)) continue;
    usedDel.add(del);
    usedIns.add(ins);
    // Retype the delete into a replace spanning both ranges; drop the insert.
    del.type = 'replace';
    del.revRange = ins.revRange;
    del.fuzzyAnchored = true;
    del.lowConfidence = true;
    removeIns.add(ins);
  }

  return removeIns.size === 0 ? ops : ops.filter((o) => !removeIns.has(o));
}

// ---------------------------------------------------------------------------
// Similarity (normalised Levenshtein, length-prefiltered)
// ---------------------------------------------------------------------------

/** 1 − levenshtein/maxLen on the (capped) strings. 1 = identical, 0 = disjoint. */
export function similarity(a: string, b: string): number {
  if (a === b) return 1;
  const s = a.length > SIMILARITY_LEN_CAP ? a.slice(0, SIMILARITY_LEN_CAP) : a;
  const t = b.length > SIMILARITY_LEN_CAP ? b.slice(0, SIMILARITY_LEN_CAP) : b;
  const maxLen = Math.max(s.length, t.length);
  if (maxLen === 0) return 1;
  // Cheap length prefilter: very different lengths can't clear a high threshold.
  if (Math.min(s.length, t.length) / maxLen < FUZZY_SIMILARITY_THRESHOLD) {
    // Still allow if one is a near-prefix of the other handled by Levenshtein,
    // but skip the expensive DP when the length gap alone rules it out.
    return Math.min(s.length, t.length) / maxLen;
  }
  return 1 - levenshtein(s, t) / maxLen;
}

/** Standard Levenshtein edit distance with a rolling two-row buffer. */
function levenshtein(s: string, t: string): number {
  const n = s.length;
  const m = t.length;
  let prev = new Array<number>(m + 1);
  let curr = new Array<number>(m + 1);
  for (let j = 0; j <= m; j++) prev[j] = j;
  for (let i = 1; i <= n; i++) {
    curr[0] = i;
    const sc = s.charCodeAt(i - 1);
    for (let j = 1; j <= m; j++) {
      const cost = sc === t.charCodeAt(j - 1) ? 0 : 1;
      curr[j] = Math.min(prev[j]! + 1, curr[j - 1]! + 1, prev[j - 1]! + cost);
    }
    [prev, curr] = [curr, prev];
  }
  return prev[m]!;
}
