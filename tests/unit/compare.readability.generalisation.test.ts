/**
 * Feature 08 — Compare §5.3 · does the readability work GENERALISE?
 *
 * `compare.readability.test.ts` ratchets 30 curated cases — the ones a reviewer
 * actually looked at. That is the right gate for "does it read well", and the
 * wrong gate for "did we overfit to thirty sentences". This suite sweeps a
 * generated corpus of realistic clause edits that nobody chose, and asserts the
 * property the curated corpus is a sample of:
 *
 *   no rendering leaves the reader to diff a replace block themselves.
 *
 * It is deliberately blunt — three aggregate numbers — so it pins the shape of
 * the result without pinning any individual rendering, which would make every
 * future improvement a test edit.
 *
 * **This suite asserted a stronger property than the tool delivers, and passed
 * because its corpus was degenerate.** The generator multiplied a 31-bit state
 * by 1103515245; that product exceeds 2^53, so the arithmetic was lossy and the
 * "random" draw was not one. Over 4,000 pairs it produced **92 distinct
 * originals**, `rnd(8)` returned only 0, 1 or 3, and two of the six editing acts
 * were drawn about 0.4% of the time. Swept with a generator that works
 * (mulberry32, exact 32-bit `Math.imul`), the same code yields **3,385 distinct
 * originals** and the floor is breached 158 times. The numbers below are that
 * honest measurement, and the floor is now a disclosed ratchet rather than a
 * clean pass — see `KNOWN_OFFENDERS`.
 *
 * Measured discrimination, on the corrected corpus:
 *
 *   | chain                                       | mean redundancy | blocks ≥ 0.60 |
 *   | ------------------------------------------- | --------------- | ------------- |
 *   | as shipped                                  | 0.0388          | 66            |
 *   | without `cleanupPreservingRelocations`      | 0.0635          | 158           |
 *   | without `detectTransposition`               | 0.0388          | 66            |
 *   | without `refineReplaceBlocks`               | 0.1422          | 397           |
 *   | without both                                | 0.1693          | 495           |
 *
 * 149 of the 158 offenders this suite first recorded were slot relocations —
 * one span struck where it sat and re-inserted a few words later, rendered as
 * a block that strikes and re-inserts the untouched middle. `diff_main` gets
 * that shape right and `diff_cleanupSemantic` destroys it, so ADR-0034 applies
 * the cleanup everywhere EXCEPT across a verbatim relocation. That took the
 * ratchet from 158 to 66 with no change to the marks average and none to the
 * curated corpus.
 *
 * `detectTransposition` still changes nothing this suite measures — it matches
 * one shape in the COALESCED diff, which the fusion had already erased. (An
 * earlier version of this comment claimed ADR-0034 prevents that shape forming;
 * measured, the pass fires 0 times either way. No effect in either direction.)
 *
 * ── WHAT `KNOWN_OFFENDERS` IS, AND WHAT IT IS NOT ─────────────────────────
 *
 * It is a **regression tripwire under stress**. It is NOT an estimate of what a
 * reader will meet, and it must not be quoted as a user-facing defect count.
 * Every one of the 66 survivors was pulled out and read:
 *
 *   a phrase genuinely relocated, nothing else changed        4
 *   a relocation PLUS another edit inside the same span      62
 *     (24 parenthetical added · 19 slot substituted or
 *      dropped · 12 qualifier added · 3 shall->must · 4 other)
 *   marks: min 2, median 2, max 6
 *
 * Where the two sides differ in CONTENT and not only in order, one replaced
 * phrase is a defensible rendering rather than a defect — and median marks is
 * already 2, the smallest a replaced phrase can have. Worse, `edit`'s slot swap
 * (`kind === 3`) permutes adjacent fragments with no regard for grammar, so the
 * revised texts behind these are not English:
 *
 *   "shall not The Company hold harmless for any Liability the Vendor ..."
 *   "The Guarantor shall not the Vendor hold harmless howsoever arising ..."
 *
 * That is fine for its purpose — a scrambled clause is a hard case, and the
 * ratchet catches any pass that renders hard cases worse. It is not fine as a
 * quality metric. **The curated 30 in `tests/fixtures/readability-cases.ts` —
 * lawyer-plausible and reviewer-approved — is the realistic measure.** Do not
 * write code to lower this number; write code when a curated case or a real
 * user says the redline reads badly.
 */
import { describe, it, expect } from 'vitest';

import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import type { CompareConfig } from '@shared/config/index.js';
import type { WordDiff } from '@features/compare/model.js';
import { wordDiff } from '@features/compare/diff/index.js';

import { readerTokens, hasReadableContent } from '../fixtures/readability-metrics.js';

const CFG: CompareConfig = getCompareConfig(SG_PROFILE);

/** Pairs swept. ~1s at this size; the numbers below are stable well past it. */
const CASES = 4000;

/**
 * Blocks at or above the 0.60 floor as the tool stands today. This is a
 * DISCLOSED GAP being ratcheted down, not a budget to spend. It was 158 when
 * the sweep was first measured honestly; ADR-0034 (relocation-preserving
 * cleanup) took it to 66. What is left is not the single-span relocation that
 * pass recognises but wider permutations — two slots exchanged with no
 * untouched text between them, or one of them edited in transit — where
 * `diff_main` never produces a `del(X) eq eq ins(X)` window in the first place.
 */
const KNOWN_OFFENDERS = 66;

/**
 * The share of the shorter side of a struck-beside-inserted block that
 * reappears verbatim on the longer side — i.e. how much of the block is the
 * tool's work left undone (P2). Same definition the curated corpus uses, scored
 * on the word diff rather than on rendered paragraphs.
 */
function redundancy(diffs: readonly WordDiff[], config: CompareConfig): number {
  let worst = 0;
  for (let i = 0; i + 1 < diffs.length; i++) {
    const a = diffs[i]!;
    const b = diffs[i + 1]!;
    if (!((a.op === 'del' && b.op === 'ins') || (a.op === 'ins' && b.op === 'del'))) continue;
    const words = (t: string): string[] => readerTokens(t, config).filter((x) => /\S/.test(x));
    const del = words(a.op === 'del' ? a.text : b.text);
    const ins = words(a.op === 'del' ? b.text : a.text);
    const [short, long] = del.length <= ins.length ? [del, ins] : [ins, del];
    if (short.length === 0) continue;
    // A block whose shorter side carries no letter or digit cannot be "read
    // twice": `hit / short.length` scores a lone struck `.` beside an insertion
    // ending in one as 1.00, the maximum, on nothing. 162 of the 390 blocks the
    // corrected corpus first flagged were this artefact, not a rendering defect.
    if (!short.some((t) => hasReadableContent(t))) continue;
    const pool = new Map<string, number>();
    for (const t of long) pool.set(t, (pool.get(t) ?? 0) + 1);
    let hit = 0;
    for (const t of short) {
      const c = pool.get(t) ?? 0;
      if (c > 0) {
        hit++;
        pool.set(t, c - 1);
      }
    }
    worst = Math.max(worst, hit / short.length);
  }
  return worst;
}

// ── generator ──────────────────────────────────────────────────────────────
// Deterministic LCG: the same corpus on every run, on every machine, and no
// Math.random anywhere in the suite.
// mulberry32. The previous LCG multiplied a 31-bit state by 1103515245, whose
// product overflows the 2^53 exact-integer range of a double, so the low bits
// were noise and `% n` collapsed the corpus. `Math.imul` keeps the arithmetic
// exactly 32-bit, which is the whole point.
function makeRnd(seed: number): (n: number) => number {
  let s = seed >>> 0;
  return (n: number): number => {
    s = (s + 0x6d2b79f5) >>> 0;
    let t = Math.imul(s ^ (s >>> 15), 1 | s);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) % n;
  };
}

const SUBJ = ['The Vendor', 'The Purchaser', 'Each Party', 'The Guarantor', 'The Company'];
const MODAL = ['shall', 'must', 'shall not', 'may', 'is obliged to'];
const VERB = ['indemnify', 'defend', 'hold harmless', 'notify', 'reimburse', 'pay'];
const OBJ = ['the Purchaser', 'the Vendor', 'each Indemnified Person', 'the Seller'];
const PREP = [
  'against all Losses',
  'in respect of any Claim',
  'for any Liability',
  'against any and all Losses whatsoever',
];
// Tails carry every atomic kind, so the sweep also exercises P5 (an amount,
// date, duration, percentage or cross-reference changes whole or not at all).
const TAIL = [
  'arising out of any breach of the Warranties',
  'howsoever arising in connection with this Agreement',
  'arising directly or indirectly out of or in connection with the Transaction',
  'in accordance with Clause 4.2',
  'within 30 days of the Completion Date',
  'up to a maximum of S$1,250,000',
  'on or before 1 January 2026',
  'subject to a threshold of 7.5%',
];
const QUALIFIER = ['promptly', 'in writing', 'without delay', 'forthwith'];
const ALL = [...SUBJ, ...MODAL, ...VERB, ...OBJ, ...PREP, ...TAIL];

describe('redline readability — generalisation sweep (§5.3)', () => {
  it('never leaves a replace block for the reader to diff, over 4,000 unseen clause pairs', () => {
    const rnd = makeRnd(424242);
    const pick = <T,>(xs: readonly T[]): T => xs[rnd(xs.length)]!;
    const clause = (): string[] => [
      pick(SUBJ),
      pick(MODAL),
      pick(VERB),
      pick(OBJ),
      pick(PREP),
      pick(TAIL),
    ];

    /** One editing act: substitute a slot, drop it, qualify it, or swap two. */
    const edit = (parts: string[]): string[] => {
      const out = [...parts];
      const kind = rnd(6);
      const at = rnd(out.length);
      if (kind === 0) out[at] = pick(ALL);
      else if (kind === 1) out.splice(at, 1);
      else if (kind === 2) out.splice(at, 0, pick(QUALIFIER));
      else if (kind === 3 && at + 1 < out.length) [out[at], out[at + 1]] = [out[at + 1]!, out[at]!];
      else if (kind === 4) out[at] = `${out[at]!} (as the case may be)`;
      else out[at] = out[at]!.replace(/\ball\b/, 'any and all').replace(/\bshall\b/, 'must');
      return out;
    };

    let sumRedundancy = 0;
    let sumMarks = 0;
    let swept = 0;
    const offenders: string[] = [];
    const sweptOriginals: string[] = [];

    for (let n = 0; n < CASES; n++) {
      const a = clause();
      let b = [...a];
      for (let e = 0; e < 1 + rnd(3); e++) b = edit(b);
      const orig = `${a.join(' ')}.`;
      const rev = `${b.join(' ')}.`;
      if (orig === rev) continue;

      const out = wordDiff(orig, rev, CFG);
      swept++;
      sweptOriginals.push(orig);
      sumRedundancy += redundancy(out, CFG);
      sumMarks += out.filter((d) => d.op !== 'eq').length;

      // ADR-0008 on every single pair: the chain may only re-partition the
      // same two texts.
      const side = (s: 'ins' | 'del'): string =>
        out.filter((d) => d.op === 'eq' || d.op === s).map((d) => d.text).join('');
      if (side('del') !== orig || side('ins') !== rev) {
        offenders.push(`reconstruction: ${orig} -> ${rev}`);
        continue;
      }

      // The hard floor. 0.60 means most of the shorter side of a block is
      // repeated verbatim on the longer side — the "diff salad" the reader
      // has to untangle. Not one pair may reach it.
      if (redundancy(out, CFG) >= 0.6) {
        offenders.push(
          out
            .map((d) => (d.op === 'eq' ? d.text : d.op === 'del' ? `[-${d.text}-]` : `{+${d.text}+}`))
            .join(''),
        );
      }
    }

    expect(swept).toBeGreaterThan(CASES * 0.9);
    // The corpus must actually be a corpus. The previous generator produced 92
    // distinct originals here, and the aggregates below were first tuned against that.
    expect(new Set(sweptOriginals).size).toBeGreaterThan(CASES * 0.8);

    // A RATCHET on a disclosed gap, not a clean pass. The number may only go
    // DOWN; raising it means a rendering got worse.
    expect(offenders.length).toBeLessThanOrEqual(KNOWN_OFFENDERS);

    // Aggregates, a little above the measured 0.0388 / 2.3287 so ordinary
    // improvement does not trip the gate, but below the 0.0635 the chain scores
    // without `cleanupPreservingRelocations` and far below the 0.1422 it scores
    // without `refineReplaceBlocks`.
    expect(sumRedundancy / swept).toBeLessThan(0.045);
    expect(sumMarks / swept).toBeLessThan(2.5);
  });
});
