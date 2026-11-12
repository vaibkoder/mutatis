/**
 * Change-complexity threshold — demote a too-dissimilar `replace` pair to a
 * clean whole-paragraph delete + insert.
 * Spec: spec/features/08-compare.md §5.3; ADR-0026.
 *
 * Level-1 emits a `replace` for every paragraph pair it aligns with unequal
 * hashes (since ADR-0039 a gap is paired by similarity with a positional
 * fallback; its lock runs `demoteDissimilarReplaces` itself, and its candidates
 * share this module's Dice), and Level-2 then word-diffs it inline. That inline diff is right when
 * the two paragraphs are similar (a changed amount, a `shall → may` swap) but
 * produces unreadable "diff salad" when the pair is a near-total rewrite — a few
 * incidental shared words shown as `equal` islands in a sea of ins/del.
 *
 * This pass mirrors the conventional "if the match ratio is below
 * a threshold, show the whole paragraph as deleted and the new one as inserted"
 * rule (the legal-comparator readability convention). It runs BEFORE the word
 * diff so the demoted pair is never word-diffed.
 *
 * Never suppresses (ADR-0008): a demoted pair emits the ENTIRE original
 * paragraph as a delete and the ENTIRE revised paragraph as an insert — strictly
 * MORE revision markup than the inline diff, never less. Word-vs-Word only; the
 * PDF pipeline keeps its opposite-direction fuzzy-anchor re-pairing (ADR-0017).
 *
 * Pure: returns a new op list; does not mutate the input ops.
 */
import type { CompareConfig } from '@shared/config/index.js';

import type { ParaRecord, DiffOp } from '../model.js';

import { tokenize } from './atomic-tokens.js';

/**
 * Sørensen–Dice similarity of two paragraph texts over their content tokens
 * (the atomic-aware word tokenizer, whitespace tokens excluded). Returns a value
 * in [0,1]: 1 = identical token multiset, 0 = disjoint. Two empty texts → 1.
 */
export function paragraphSimilarity(a: string, b: string, config: CompareConfig): number {
  return bagDice(bagOf(contentTokens(a, config)), bagOf(contentTokens(b, config)));
}

/** A token multiset: how many tokens, and how many of each. */
export interface TokenBag {
  size: number;
  counts: Map<string, number>;
}

/** Count `tokens` into a {@link TokenBag}. */
export function bagOf(tokens: readonly string[]): TokenBag {
  const counts = new Map<string, number>();
  for (const t of tokens) counts.set(t, (counts.get(t) ?? 0) + 1);
  return { size: tokens.length, counts };
}

/**
 * Sørensen–Dice over two token multisets: `2·|A∩B| / (|A|+|B|)`, with two EMPTY
 * bags scoring 1 (identical). The ONE implementation of the arithmetic
 *: `paragraphSimilarity` here and `absorbs` in
 * `merge-split.ts` call it; `gap-pairing.ts` computes the same intersection over
 * interned ids and shares the formula through {@link diceFromOverlap}, and a test
 * pins that the two agree on every pair it scores.
 */
export function bagDice(a: TokenBag, b: TokenBag): number {
  return diceFromOverlap(sharedTokens(a, b), a.size, b.size);
}

/**
 * Dice from a multiset intersection size and the two multiset sizes — the
 * formula itself, shared by {@link bagDice} and by `gap-pairing.ts`, which
 * computes the same intersection over interned, sorted ids for speed. Two
 * empty multisets score 1.
 */
export function diceFromOverlap(overlap: number, sizeA: number, sizeB: number): number {
  const total = sizeA + sizeB;
  if (total === 0) return 1;
  return (2 * overlap) / total;
}

/**
 * Content tokens for the similarity metric: the atomic-aware word tokens,
 * lowercased and stripped of leading/trailing punctuation, with anything left
 * empty (pure whitespace / bare punctuation) dropped. Normalising the edges
 * matters because the word tokenizer keeps a trailing full stop on a word
 * (`Alpha.` ≠ `Alpha`); without it a small edit on a short sentence would score
 * near-zero similarity and demote spuriously. Atomic tokens (amounts, dates,
 * quoted terms) keep their distinguishing core, so a value change is still
 * counted as a mismatch.
 */
export function contentTokens(text: string, config: CompareConfig): string[] {
  const out: string[] = [];
  for (const raw of tokenize(text, config)) {
    const t = raw
      .toLowerCase()
      .replace(/^[^\p{L}\p{N}]+/u, '')
      .replace(/[^\p{L}\p{N}]+$/u, '');
    if (t.length > 0) out.push(t);
  }
  return out;
}

/**
 * Length of the longest UNBROKEN run of content tokens common to both sides.
 *
 * A rolling two-row DP; only the previous row is kept, so this is O(|a|·|b|)
 * time and O(|b|) space over paragraph-sized inputs. Deliberately measures a
 * CONTIGUOUS run and not an LCS: an LCS would happily stitch a shared word
 * here to a shared word forty tokens later, which is precisely the scattered
 * shape the demotion exists to catch.
 */
export function longestContiguousRun(a: readonly string[], b: readonly string[]): number {
  let best = 0;
  const prev = new Array<number>(b.length + 1).fill(0);
  for (let i = 1; i <= a.length; i++) {
    let diag = 0;
    for (let j = 1; j <= b.length; j++) {
      const above = prev[j]!;
      prev[j] = a[i - 1] === b[j - 1] ? diag + 1 : 0;
      if (prev[j]! > best) best = prev[j]!;
      diag = above;
    }
  }
  return best;
}

/**
 * Share of the SHORTER paragraph's content tokens that survive as one unbroken
 * run in the other. 1 = the shorter side sits verbatim inside the longer one.
 *
 * The *shorter* side is the denominator on purpose. A large deletion leaves a
 * short survivor wholly contained in a long original; scoring against the
 * longer side would report ~0.2 for a rendering that is one clean
 * strike-through, which is the very asymmetry Dice already gets wrong.
 */
export function contiguousShare(a: string, b: string, config: CompareConfig): number {
  const aTokens = contentTokens(a, config);
  const bTokens = contentTokens(b, config);
  const floor = Math.min(aTokens.length, bTokens.length);
  if (floor === 0) return 0;
  return longestContiguousRun(aTokens, bTokens) / floor;
}

/**
 * Join a record range's `normText` with `separator`. The ONE implementation of
 * this walk: `rangeDiffText` (ADR-0032) does it with a paragraph-mark sentinel
 * and this module with a space (reuse before build — two implementations of the
 * same thing is a defect). It lives here because `contentTokens` already does,
 * so `merge-split.ts` depends on this module and not the other way round.
 */
export function joinRange(
  records: readonly ParaRecord[],
  range: { start: number; end: number } | null,
  separator: string,
): string {
  if (range === null) return '';
  const parts: string[] = [];
  for (let i = range.start; i <= range.end; i++) {
    const rec = records[i];
    if (rec) parts.push(rec.normText);
  }
  return parts.join(separator);
}

/**
 * A record range's `normText` joined with a space. Whitespace and
 * bare-punctuation tokens are dropped by `contentTokens` anyway, so the
 * separator never affects the score — it only keeps two paragraphs' words from
 * fusing across the join. Shares `joinRange` with `rangeDiffText`, which does
 * the same walk with a paragraph-mark sentinel.
 */
function joinWithSpace(
  records: readonly ParaRecord[],
  range: { start: number; end: number } | null,
): string {
  return joinRange(records, range, ' ');
}

/**
 * ADR-0026's keep-or-demote decision for one paragraph pair: `true` when the
 * pair is similar enough to read inline, `false` when it is to be shown as a
 * whole-paragraph delete + insert.
 *
 * The ONE implementation of the gate, used by
 * `demoteDissimilarReplaces` for every pair it judges. ADR-0039's gap pairing
 * reaches it only THROUGH `demoteDissimilarReplaces` (its lock runs that pass
 * over the gap's positional ops rather than asking this predicate about a pair
 * on its own, which missed merges — 08f task 7), and its candidates use this
 * module's Dice arithmetic, so a pair it chooses clears this gate.
 *
 * A threshold ≤ 0 disables the demotion, so every pair survives.
 */
export function survivesDemotion(origText: string, revText: string, config: CompareConfig): boolean {
  if (config.replaceSimilarityThreshold <= 0) return true;
  const aTokens = contentTokens(origText, config);
  const bTokens = contentTokens(revText, config);
  const a = bagOf(aTokens);
  const b = bagOf(bTokens);
  const overlap = sharedTokens(a, b);
  if (diceFromOverlap(overlap, a.size, b.size) >= config.replaceSimilarityThreshold) {
    return true;
  }
  // Contiguity escape (ADR-0041). Dice is a SET measure, so it punishes
  // length asymmetry — and a large contiguous deletion is nothing but length
  // asymmetry. Cutting two sentences from a three-sentence clause scores
  // ~0.36 and demotes here, yet inline it is one strike-through with the rest
  // untouched. Measured over 4,000 generated pairs of each shape: clean
  // large deletions score 1.00 on this measure without exception, while
  // genuine rewrites reach 0.68 at the 99th percentile and 0.81 at worst, so
  // the default 0.8 keeps every clean cut and re-admits none of the salad.
  //
  // `contiguousShare` is O(|a|·|b|). An unbroken shared run can be no longer
  // than the shared multiset, so when that is already short of the bar the run
  // search cannot clear it and is skipped. Exact, not a heuristic: the skip
  // fires only where `contiguousShare` would return less than `overlap / floor`,
  // itself below the bar. It is what keeps a gap of long, dissimilar paragraphs
  // from costing seconds (08f task 7).
  const floor = Math.min(a.size, b.size);
  if (floor > 0 && overlap / floor < config.replaceContiguityThreshold) return false;
  return contiguousShare(origText, revText, config) >= config.replaceContiguityThreshold;
}

/** Multiset intersection size of two token bags. */
function sharedTokens(a: TokenBag, b: TokenBag): number {
  const [small, large] = a.counts.size <= b.counts.size ? [a.counts, b.counts] : [b.counts, a.counts];
  let overlap = 0;
  for (const [token, n] of small) overlap += Math.min(n, large.get(token) ?? 0);
  return overlap;
}

/**
 * Demote every `replace` op whose paragraph similarity is below
 * `config.replaceSimilarityThreshold` into a `delete` + `insert` pair. A
 * threshold ≤ 0 disables the pass (similarity is always ≥ 0). Non-replace ops,
 * and replaces at or above the threshold, pass through unchanged.
 */
export function demoteDissimilarReplaces(
  ops: DiffOp[],
  original: ParaRecord[],
  revised: ParaRecord[],
  config: CompareConfig,
  opts: { keepLastParagraph?: boolean } = {},
): DiffOp[] {
  const threshold = config.replaceSimilarityThreshold;
  if (threshold <= 0) return ops;

  const out: DiffOp[] = [];
  for (const op of ops) {
    if (op.type !== 'replace' || !op.origRange || !op.revRange) {
      out.push(op);
      continue;
    }
    const origPara = original[op.origRange.start];
    const revPara = revised[op.revRange.start];
    // ADR-0046: a Word table pair is revised cell by cell, and each cell's
    // content passes through this demotion on its own. Striking and
    // re-inserting a whole schedule for one changed amount is not readability.
    if (origPara === undefined || revPara === undefined || origPara.wordTable || revPara.wordTable) {
      out.push(op);
      continue;
    }
    // ADR-0046: inside a table cell, the LAST paragraph's mark can never carry
    // a revision (a cell must end with a paragraph — ADR-0035), so demoting a
    // pair that includes it leaves a blank paragraph in the cell whichever way
    // the reader resolves it. Kept inline: a cell value's change reads as the
    // new value then the old one in the one paragraph, as a cell change should.
    if (
      opts.keepLastParagraph === true &&
      (op.origRange.end === original.length - 1 || op.revRange.end === revised.length - 1)
    ) {
      out.push(op);
      continue;
    }

    // Merge/split ops span more than one paragraph on one side, so the pair is
    // scored on the JOINED text — otherwise a merge would be judged against
    // only the first of the paragraphs it replaces.
    const origText = joinWithSpace(original, op.origRange);
    const revText = joinWithSpace(revised, op.revRange);
    if (survivesDemotion(origText, revText, config)) {
      out.push(op);
      continue;
    }

    // Too dissimilar to read inline → clean whole-paragraph delete + insert.
    // Built fresh (not spread) so no inline-diff/move/cascade optional fields
    // leak onto the demoted ops. Both full paragraphs are preserved (delete
    // carries the original, insert the revised) — readability gain, never
    // suppression.
    out.push({
      id: `${op.id}-d`,
      type: 'delete',
      origRange: op.origRange,
      revRange: null,
      partId: origPara.partId,
      classification: 'structural',
      classes: ['structural'],
      author: op.author,
      date: op.date,
    });
    out.push({
      id: `${op.id}-i`,
      type: 'insert',
      origRange: null,
      revRange: op.revRange,
      partId: revPara.partId,
      classification: 'structural',
      classes: ['structural'],
      author: op.author,
      date: op.date,
    });
  }
  return out;
}
