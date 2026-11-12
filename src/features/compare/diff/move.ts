/**
 * Level-3 verbatim moved-paragraph detection.
 * Spec: spec/features/08-compare.md §5.4, §2.2.
 *
 * Post-processes the structuralDiff output: any paragraph hash present in BOTH
 * the delete set (Original-only) and the insert set (Revised-only) is a verbatim
 * move — removed from position X, re-appearing at position Y unchanged. Pair the
 * delete and insert ops, retype both as `move`, and cross-link via moveTargetId.
 *
 * Duplicated boilerplate (a hash appearing multiple times in either set) is
 * matched greedily by nearest index; residue is left as plain insert/delete
 * (we never guess). Pure — mutates op type/moveTargetId in place (per A4 spec).
 *
 * v1 is verbatim hash-equality ONLY (spec §2.2). The v2 fuzzy-move seam is left
 * as a comment below, not code.
 */
import type { ParaRecord, DiffOp } from '../model.js';

export function detectMoves(
  ops: DiffOp[],
  original: ParaRecord[],
  revised: ParaRecord[],
): DiffOp[] {
  // Group delete ops and insert ops by the hash of their referenced paragraph.
  const deletesByHash = new Map<string, DiffOp[]>();
  const insertsByHash = new Map<string, DiffOp[]>();

  for (const op of ops) {
    if (op.type === 'delete' && op.origRange) {
      const rec = original[op.origRange.start];
      // ADR-0027: table-unit records (text-vs-text Markdown) are never move
      // candidates — a relocated table reads as delete+insert (the run-level
      // moveFrom/moveTo markup applies to paragraphs, not <w:tbl>). Inert for
      // Word-vs-Word, which never produces table-unit records.
      if (rec?.table) continue;
      const h = rec?.hash;
      if (h !== undefined) pushInto(deletesByHash, h, op);
    } else if (op.type === 'insert' && op.revRange) {
      const rec = revised[op.revRange.start];
      if (rec?.table) continue;
      const h = rec?.hash;
      if (h !== undefined) pushInto(insertsByHash, h, op);
    }
  }

  // For each hash present in both maps, pair greedily by nearest index.
  for (const [hash, dels] of deletesByHash) {
    const inss = insertsByHash.get(hash);
    if (!inss || inss.length === 0) continue;

    // Build candidate (delOp, insOp) pairs ordered by |origIndex - revIndex|.
    const candidates: Array<{ del: DiffOp; ins: DiffOp; dist: number }> = [];
    for (const del of dels) {
      for (const ins of inss) {
        const oIdx = del.origRange?.start ?? 0;
        const rIdx = ins.revRange?.start ?? 0;
        candidates.push({ del, ins, dist: Math.abs(oIdx - rIdx) });
      }
    }
    candidates.sort((a, b) => a.dist - b.dist);

    const usedDel = new Set<DiffOp>();
    const usedIns = new Set<DiffOp>();
    for (const { del, ins } of candidates) {
      if (usedDel.has(del) || usedIns.has(ins)) continue;
      usedDel.add(del);
      usedIns.add(ins);

      // Retype both ops as `move` and cross-link. `structural` classification
      // is preserved (move is a structural op per §6 row 1).
      del.type = 'move';
      ins.type = 'move';
      del.moveTargetId = ins.id;
      ins.moveTargetId = del.id;
    }
    // Any del/ins not used (count mismatch in duplicated boilerplate) is left as
    // a plain insert/delete — residue, not guessed.
  }

  // Similarity-threshold matching for relocated-and-edited paragraphs is NOT
  // done here — Word-vs-Word stays verbatim hash equality only (spec §2.2). The PDF path (feature 12, ADR-0017) realises a SCOPED fuzzy
  // fallback in pdf/fuzzy-anchor.ts (FUZZY_SIMILARITY_THRESHOLD), applied only to
  // reconstructed-PDF paragraphs before word diff; it never runs for Word-vs-Word.

  return ops;
}

/** Append an op to the array stored at `key` in `map`, creating it if needed. */
function pushInto(map: Map<string, DiffOp[]>, key: string, op: DiffOp): void {
  const arr = map.get(key);
  if (arr) arr.push(op);
  else map.set(key, [op]);
}
