/**
 * Feature 08 — Compare §5.3 · relocation-preserving semantic cleanup (ADR-0034).
 *
 * `diff_main` renders a slot relocation correctly — one span struck where it
 * sat, re-inserted a few words later, the text between untouched. In SENTINEL
 * space, where one character is one TOKEN, `diff_cleanupSemantic` then measures
 * the equality between the two ends as shorter than the edits either side and
 * eliminates it, fusing the move into a block that strikes and re-inserts the
 * middle for nothing (P1, P2).
 *
 * These tests pin: the shape that survives, each guard that keeps the pass
 * narrow, the no-op property (nothing changes where no relocation is present),
 * and the invariant nothing may break — both sides reconstruct exactly
 * (ADR-0008).
 */
import { describe, it, expect } from 'vitest';

import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import { wordDiff, cleanupPreservingRelocations } from '@features/compare/diff/index.js';
import type { DmpDiff } from '@features/compare/diff/index.js';
import type { Token } from '@features/compare/diff/atomic-tokens.js';
import type { WordDiff } from '@features/compare/model.js';

const SG = getCompareConfig(SG_PROFILE);

function render(diffs: readonly WordDiff[]): string {
  return diffs
    .map((d) => (d.op === 'eq' ? d.text : d.op === 'del' ? `[-${d.text}-]` : `{+${d.text}+}`))
    .join('');
}

function side(diffs: readonly WordDiff[], which: 'ins' | 'del'): string {
  return diffs
    .filter((d) => d.op === 'eq' || d.op === which)
    .map((d) => d.text)
    .join('');
}

/**
 * A tiny sentinel alphabet for the unit-level tests: one letter per token, so a
 * diff can be written by hand. `_` is whitespace, `.` is punctuation.
 */
const ALPHABET: Record<string, Token> = {
  a: { text: 'shall' },
  b: { text: 'notify' },
  c: { text: 'promptly' },
  d: { text: 'Vendor' },
  // Same TEXT as `d`, different atomic kind. Upstream keys sentinels on kind +
  // text (ADR-0029, `word.ts`), so these are two different characters and the
  // identity guard must refuse to pair them.
  D: { text: 'Vendor', atomic: 'clauseRef' },
  _: { text: ' ' },
  '.': { text: '.' },
};
const tokenAt = (ch: string): Token => ALPHABET[ch] ?? { text: ch };

/** Sentinel-space `[op, tokens]` runs, as `diff_main` would hand them over. */
function diff(...runs: Array<[number, string]>): DmpDiff[] {
  return runs.map(([op, s]): DmpDiff => [op, s]);
}

const DEL = -1;
const INS = 1;
const EQ = 0;

describe('cleanupPreservingRelocations (§5.3, ADR-0034)', () => {
  // ── the shape it exists for ───────────────────────────────────────────────

  it('keeps a relocated word out of the untouched text beside it', () => {
    const out = wordDiff(
      'The Purchaser shall notify each Indemnified Person.',
      'The Purchaser notify shall each Indemnified Person.',
      SG,
    );
    expect(render(out)).toBe(
      'The Purchaser [-shall -]notify {+shall +}each Indemnified Person.',
    );
  });

  it('relocates a multi-word phrase without borrowing the words between', () => {
    const out = wordDiff(
      'The Guarantor may hold harmless the Vendor for any Liability.',
      'The Guarantor may the Vendor hold harmless for any Liability.',
      SG,
    );
    expect(render(out)).toBe(
      'The Guarantor may [-hold harmless -]the Vendor {+hold harmless +}for any Liability.',
    );
  });

  it('never adds to or drops from either side (ADR-0008)', () => {
    const original = 'The Company shall not pay the Vendor for any Liability.';
    const revised = 'The Company pay shall not the Vendor for any Liability.';
    const out = wordDiff(original, revised, SG);
    expect(side(out, 'del')).toBe(original);
    expect(side(out, 'ins')).toBe(revised);
  });

  // ── guards ────────────────────────────────────────────────────────────────

  it('GUARD verbatim identity — a near-match is not a relocation', () => {
    // `a` (shall) is struck and `c` (promptly) inserted: different spans, so the
    // equality is fair game for the ordinary cleanup and is eliminated.
    const out = cleanupPreservingRelocations(
      diff([DEL, 'a_'], [EQ, 'b'], [INS, '_c']),
      tokenAt,
    );
    expect(out.map(([op]) => op)).toEqual([DEL, INS]);
  });

  it('GUARD verbatim identity — whitespace placement may differ, tokens may not', () => {
    // `del("shall ")` vs `ins(" shall")`: the same span, differently spaced.
    const out = cleanupPreservingRelocations(
      diff([DEL, 'a_'], [EQ, 'b'], [INS, '_a']),
      tokenAt,
    );
    expect(out.map(([op]) => op)).toEqual([DEL, EQ, INS]);
  });

  it('GUARD readable content — a punctuation-only span is not a relocation', () => {
    // `del(". ")` / `ins(" .")` around an equality: verbatim-identical spans,
    // but nothing a reader could watch move. The ordinary cleanup owns it.
    const out = cleanupPreservingRelocations(
      diff([DEL, '._'], [EQ, 'b'], [INS, '_.']),
      tokenAt,
    );
    expect(out.filter(([, s]) => s !== '').map(([op]) => op)).toEqual([DEL, INS]);
  });

  // ── no-op where there is nothing to protect ───────────────────────────────

  it('is byte-identical to diff_cleanupSemantic where no relocation is present', () => {
    // A plain one-word substitution: the pass must not perturb the cleanup.
    const out = wordDiff(
      'The Vendor shall indemnify the Purchaser against all Losses.',
      'The Vendor shall indemnify the Purchaser against all Claims.',
      SG,
    );
    expect(render(out)).toBe(
      'The Vendor shall indemnify the Purchaser against all [-Losses.-]{+Claims.+}',
    );
  });

  it('leaves the coarse doublet rendering ADR-0030 refines to untouched', () => {
    const out = wordDiff(
      'The Vendor shall indemnify and hold harmless the Purchaser against any and all Losses whatsoever.',
      'The Vendor shall indemnify the Purchaser against all Losses.',
      SG,
    );
    expect(render(out)).toBe(
      'The Vendor shall indemnify [-and hold harmless -]the Purchaser against [-any and -]all Losses[- whatsoever-].',
    );
  });

  // ── the never-suppress fallback ───────────────────────────────────────────

  it('renders a relocation beside cross-reference values without disturbing them', () => {
    // END-TO-END, and it pins only that: this input is never fused, so it
    // passes with the pass ablated. The kind-keying and mirror-order claims it
    // used to carry are pinned by the two sentinel-level tests below, which is
    // where they can actually be isolated.
    const original = 'Nothing in Clause 4.2 limits the Vendor liability under Clause 8.2.';
    const revised = 'Nothing in Clause 4.2 limits liability the Vendor under Clause 8.2.';
    const out = wordDiff(original, revised, SG);
    expect(side(out, 'del')).toBe(original);
    expect(side(out, 'ins')).toBe(revised);
    expect(render(out)).toBe(
      'Nothing in Clause 4.2 limits {+liability +}the Vendor [-liability -]under Clause 8.2.',
    );
  });

  // ── the guards the second review found unpinned ────────────────────────────

  it('GUARD the window is a del/ins PAIR — two deletions are not a relocation', () => {
    // `del(a) eq(b) del(a)`: verbatim-identical spans, but the same span struck
    // TWICE is not something that moved, and protecting it would leave the
    // ordinary cleanup unapplied. Relaxing `isPair` to "any two like ops"
    // protects 81 extra windows on the 3,675-pair sweep, so this is not
    // hypothetical — it is the one guard no other test reaches.
    const out = cleanupPreservingRelocations(diff([DEL, 'a_'], [EQ, 'b'], [DEL, '_a']), tokenAt);
    expect(out.map(([op]) => op)).not.toEqual([DEL, EQ, DEL]);
  });

  it('GUARD identity is compared on SENTINELS, so atomic kind is part of it (ADR-0029)', () => {
    // `d` and `D` are both the text "Vendor"; only `D` is a clauseRef VALUE.
    // Comparing token TEXT would pair them and protect the window; comparing
    // sentinels — which is what the code does — must refuse.
    const out = cleanupPreservingRelocations(diff([DEL, 'd_'], [EQ, 'b'], [INS, '_D']), tokenAt);
    expect(out.map(([op]) => op)).toEqual([DEL, INS]);
    // Control: the SAME shape with the same kind on both sides IS protected, so
    // the assertion above cannot pass merely because the window never fires.
    const same = cleanupPreservingRelocations(diff([DEL, 'd_'], [EQ, 'b'], [INS, '_d']), tokenAt);
    expect(same.map(([op]) => op)).toEqual([DEL, EQ, INS]);
  });

  it('protects the mirror order (ins … eq … del) as well as (del … eq … ins)', () => {
    // 144 of the 442 windows the sweep detects are this order. Previously only
    // the sweep ratchet pinned it, and a ratchet is not a test.
    const out = cleanupPreservingRelocations(diff([INS, 'a_'], [EQ, 'b'], [DEL, '_a']), tokenAt);
    expect(out.map(([op]) => op)).toEqual([INS, EQ, DEL]);
  });
});
