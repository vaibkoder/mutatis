/**
 * Paragraph merge / split detection. ADR-0032.
 * Spec: spec/features/08-compare.md §5.1–5.3. Principles: P1/P2/P9.
 *
 * Level-1 matches paragraphs one-to-one, so when a drafter joins two paragraphs
 * into one the pipeline has no way to say so: paragraph 1 pairs with the merged
 * paragraph and becomes a huge `replace`, and paragraph 2 has no partner and is
 * struck whole. The same words are then shown once struck and once inserted —
 * the reader has to diff the redline to see that nothing happened to them (P2).
 * A split (one paragraph becoming two) is the mirror.
 *
 * This pass re-pairs those op triples into ONE `replace` spanning the whole
 * range on the many-side. No new op type is needed: `DiffOp.origRange` /
 * `revRange` are already ranges, and `recordsInRange` already walks them — a
 * merge is a `replace` whose `origRange` covers two paragraphs, a split one
 * whose `revRange` does.
 *
 * The paragraph mark itself is carried into the Level-2 word diff as a sentinel
 * token ({@link PARA_SEAM}), so its removal (or addition) is a first-class
 * `del`/`ins` the author turns into a tracked, rejectable `w:pPr/w:rPr/w:del`
 * (or `…/w:ins`) on the paragraph mark. Nothing is glued away invisibly.
 *
 * Detection rests on STRUCTURE only — token identity, positional containment at
 * one end, length ratios (ADR-0008: no semantic-equivalence scoring).
 *
 * Never suppresses (ADR-0008): the pass only re-labels which op owns which
 * records. It reconstructs both sides' full text before and after and returns
 * its input unchanged on any mismatch — the guard pattern from `diff/anchor.ts`.
 *
 * Pure: returns a new op list; does not mutate the input ops.
 */
import type { CompareConfig } from '@shared/config/index.js';

import type { ParaRecord, DiffOp } from '../model.js';

import { bagDice, bagOf, contentTokens, joinRange } from './replace-threshold.js';

/**
 * Sentinel standing for a paragraph mark inside a merged/split word diff. The
 * tokenizer emits it as its own single-character token (it is neither
 * whitespace nor a word character), which is exactly what lets the diff delete
 * or insert it as a unit.
 *
 * It is a Private Use Area code point, which makes a collision with document
 * text UNLIKELY — not impossible, as this comment previously claimed.
 * `normaliseText` does not strip PUA code points, and both broken-`ToUnicode`
 * PDF extraction and verbatim pasted text yield them. Two guards stand on that:
 * `isPlain` declines to re-pair any record whose own text contains the
 * sentinel, so it can only ever appear in an op claiming a single record; and
 * `spellOutSeams` only rewrites ops claiming more than one. Neither is
 * defence-in-depth — each is load-bearing, and both are tested.
 */
export const PARA_SEAM = '\uE000';

/**
 * Floor on how much of the absorbed paragraph must reappear, in order, at the
 * matching end of the surviving paragraph. High on purpose: gluing two
 * genuinely unrelated paragraphs together is a far worse defect than leaving a
 * merge rendered the old way.
 */
export const MERGE_CONTAINMENT_FLOOR = 0.85;

/**
 * Never re-pair on a fragment. A two- or three-word paragraph (a limb stub, a
 * heading, `"; and"`) matches too easily by chance, and the readability gain on
 * one is negligible anyway.
 */
export const MERGE_MIN_TOKENS = 4;

/**
 * Join a record range into the text the Level-2 word diff sees, with a
 * {@link PARA_SEAM} standing for each paragraph mark between two records. A
 * single-record range is the record's `normText` unchanged, so every existing
 * one-to-one `replace` produces a byte-identical word diff.
 */
export function rangeDiffText(
  records: readonly ParaRecord[],
  range: { start: number; end: number } | null,
): string {
  return joinRange(records, range, PARA_SEAM);
}

/**
 * Share of `part` that reappears, token-for-token and in order, at the given
 * end of `host`. Positional (not a bag-of-words overlap) so a paragraph that
 * merely shares vocabulary with its neighbour does not read as absorbed.
 */
function endContainment(
  host: readonly string[],
  part: readonly string[],
  at: 'head' | 'tail',
): number {
  if (part.length === 0) return 0;
  const slice = at === 'tail' ? host.slice(host.length - part.length) : host.slice(0, part.length);
  if (slice.length !== part.length) return 0;
  let hit = 0;
  for (const [i, tok] of part.entries()) if (slice[i] === tok) hit++;
  return hit / part.length;
}

/** The half of `host` NOT accounted for by an absorbed `part` at `at`. */
function remainder(host: readonly string[], part: readonly string[], at: 'head' | 'tail'): string[] {
  return at === 'tail' ? host.slice(0, host.length - part.length) : host.slice(part.length);
}

/**
 * Would `part` read as having been absorbed into `host`, leaving `keep` behind?
 * Purely structural: a long-enough absorbed paragraph, reproduced in order at
 * one end of the host, with the rest of the host still recognisably `keep`.
 */
function absorbs(
  hostText: string,
  partText: string,
  keepText: string,
  at: 'head' | 'tail',
  config: CompareConfig,
): boolean {
  const host = contentTokens(hostText, config);
  const part = contentTokens(partText, config);
  const keep = contentTokens(keepText, config);

  if (part.length < MERGE_MIN_TOKENS) return false;
  // There must be something left over: a host that is ENTIRELY the absorbed
  // paragraph is a delete of the other one, not a merge.
  if (host.length <= part.length) return false;
  if (endContainment(host, part, at) < MERGE_CONTAINMENT_FLOOR) return false;

  // …and what is left over has to still be the paragraph it was paired with,
  // otherwise this is a wholesale rewrite that happens to end on similar words.
  // Reuses the ADR-0026 threshold so the two gates cannot disagree.
  return bagDice(bagOf(remainder(host, part, at)), bagOf(keep)) >= config.replaceSimilarityThreshold;
}

/** Re-point an op at a wider range on one side. Built fresh — never mutated. */
function widen(
  op: DiffOp,
  origRange: { start: number; end: number },
  revRange: { start: number; end: number },
): DiffOp {
  return { ...op, origRange, revRange };
}

/** A record that must never be re-paired: a table unit is one alignment unit (ADR-0027). */
function isPlain(rec: ParaRecord | undefined): rec is ParaRecord {
  if (rec === undefined || rec.table !== undefined) return false;
  // A document may legitimately CONTAIN U+E000: `normaliseText` does not strip
  // Private Use Area code points, and PDF text extraction with a broken
  // `ToUnicode` map yields them, as does verbatim pasted text. Joining such a
  // record with `PARA_SEAM` makes the sentinel ambiguous — the author would cut
  // the paragraph at the document's own character, emitting an extra <w:p> with
  // NO revision on it (invisible and unrejectable) and dropping the character
  // with no <w:del>. `sideText`'s guard cannot see this: it compares the set of
  // records claimed, not their text. So refuse to re-pair on such a record and
  // leave the change rendered the ordinary way (ADR-0008).
  return !rec.normText.includes(PARA_SEAM);
}

/**
 * Re-pair `replace`+`delete` (merge) and `replace`+`insert` (split) op pairs
 * into a single range-spanning `replace`. Runs on the Level-1 ops, before the
 * ADR-0026 dissimilarity gate — so the gate scores the merge against the JOINED
 * original text, which is what it actually replaces.
 *
 * `mayAbsorb`, when given, may refuse a record as the absorbed orphan: the
 * pipeline passes ADR-0039's `GapPairing.mayAbsorb`, so a record headed for a
 * Level-3 move is not swallowed beside a pair it never followed before.
 */
export function detectMergesAndSplits(
  ops: DiffOp[],
  original: ParaRecord[],
  revised: ParaRecord[],
  config: CompareConfig,
  mayAbsorb?: (record: ParaRecord, side: 'original' | 'revised') => boolean,
): DiffOp[] {
  if (ops.length < 2) return ops;

  const before = { orig: sideText(ops, original, 'orig'), rev: sideText(ops, revised, 'rev') };
  const out: DiffOp[] = [];

  for (let i = 0; i < ops.length; i++) {
    const a = ops[i]!;
    const b = ops[i + 1];
    const paired = b === undefined ? undefined : tryPair(a, b, original, revised, config, mayAbsorb);
    if (paired !== undefined) {
      out.push(paired);
      i++; // both ops consumed
      continue;
    }
    out.push(a);
  }

  // Never-suppress guard (ADR-0008): re-pairing may only change WHICH op owns a
  // record, never which records are present on either side. Keep the input if
  // that ever fails.
  if (
    sideText(out, original, 'orig') !== before.orig ||
    sideText(out, revised, 'rev') !== before.rev
  ) {
    return ops;
  }
  return out;
}

/**
 * The two matched adjacency patterns, or undefined when this pair is not a
 * merge/split.
 *
 * Only `replace`+`delete` and `replace`+`insert` are matched. Positional
 * pairing (`emitPositional` in `diff/structural.ts`) pushes every `replace` in a
 * gap first, then the leftover `delete`s, then the leftover `insert`s, so there
 * an orphan is never the op immediately BEFORE its replace; the mirrored
 * orderings were removed as unreachable (ADR-0032, correction 1).
 *
 * **Since ADR-0039 that ordering IS reachable** — a similarity-aligned pair can
 * follow a positional surplus — and is deliberately still not matched. The gap
 * pairing LOCKS every positional pair this pass would have re-paired (it runs
 * this very function to find out), so a merge or split that rendered as one op
 * before still does. An orphan-before-replace shape now arises only where the
 * positional pairing would have shown the paragraph struck and re-typed; ADR-0039
 * records it as an accepted limit rather than re-opening ADR-0032's correction.
 * The matched orderings can also arise where they did not before, with a held
 * record (one headed for a Level-3 move) as the orphan — which `mayAbsorb`
 * refuses unless positional pairing absorbed it too (08f task 11).
 */
function tryPair(
  a: DiffOp,
  b: DiffOp,
  original: ParaRecord[],
  revised: ParaRecord[],
  config: CompareConfig,
  mayAbsorb?: (record: ParaRecord, side: 'original' | 'revised') => boolean,
): DiffOp | undefined {
  // ── merge: two original paragraphs → one revised ────────────────────────
  if (a.type === 'replace' && b.type === 'delete' && a.origRange && a.revRange && b.origRange) {
    const head = original[a.origRange.end];
    const absorbed = original[b.origRange.start];
    const host = revised[a.revRange.start];
    if (
      isPlain(head) && isPlain(absorbed) && isPlain(host) &&
      (mayAbsorb?.(absorbed, 'original') ?? true) &&
      b.origRange.start === a.origRange.end + 1 &&
      b.origRange.start === b.origRange.end &&
      absorbs(host.normText, absorbed.normText, head.normText, 'tail', config)
    ) {
      return widen(a, { start: a.origRange.start, end: b.origRange.end }, a.revRange);
    }
  }
  // ── split: one original paragraph → two revised (the mirror) ────────────
  if (a.type === 'replace' && b.type === 'insert' && a.origRange && a.revRange && b.revRange) {
    const head = revised[a.revRange.end];
    const added = revised[b.revRange.start];
    const host = original[a.origRange.start];
    if (
      isPlain(head) && isPlain(added) && isPlain(host) &&
      (mayAbsorb?.(added, 'revised') ?? true) &&
      b.revRange.start === a.revRange.end + 1 &&
      b.revRange.start === b.revRange.end &&
      absorbs(host.normText, added.normText, head.normText, 'tail', config)
    ) {
      return widen(a, a.origRange, { start: a.revRange.start, end: b.revRange.end });
    }
  }
  return undefined;
}

/** Reconstruct one side's full paragraph text from the op list, for the guard. */
function sideText(
  ops: readonly DiffOp[],
  records: readonly ParaRecord[],
  side: 'orig' | 'rev',
): string {
  const parts: string[] = [];
  for (const op of ops) {
    const range = side === 'orig' ? op.origRange : op.revRange;
    if (range === null) continue;
    for (let i = range.start; i <= range.end; i++) {
      const rec = records[i];
      if (rec) parts.push(rec.normText);
    }
  }
  return parts.join('\u0000');
}
