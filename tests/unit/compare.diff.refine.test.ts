/**
 * Feature 08 — Compare §5.3 · replace-block refinement (ADR-0030).
 *
 * `diff_cleanupSemantic` measures an equality with `String.length`, and the word
 * diff runs it in sentinel space where one character is one TOKEN — so a shared
 * space out-weighs real shared prose and the fusion avalanches. This pass undoes
 * that locally. These tests pin what it must recover, what it must decline to
 * touch, and the invariant it may never break: nothing leaves the redline.
 */
import { describe, it, expect } from 'vitest';

import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import { wordDiff, refineReplaceBlocks, tokenize } from '@features/compare/diff/index.js';
import type { WordDiff } from '@features/compare/model.js';

const SG = getCompareConfig(SG_PROFILE);

/** Rebuild one side's text from the diff — what that side must still say. */
function side(diffs: readonly WordDiff[], which: 'ins' | 'del'): string {
  return diffs
    .filter((d) => d.op === 'eq' || d.op === which)
    .map((d) => d.text)
    .join('');
}

/** Reader's view: brackets round strikes, braces round insertions. */
function render(diffs: readonly WordDiff[]): string {
  return diffs
    .map((d) => (d.op === 'eq' ? d.text : d.op === 'del' ? `[-${d.text}-]` : `{+${d.text}+}`))
    .join('');
}

describe('refineReplaceBlocks (§5.3, ADR-0030)', () => {
  it('frees the prose a collapsed doublet was hiding', () => {
    const orig = 'The Vendor shall indemnify and hold harmless the Purchaser against any and all Losses whatsoever.';
    const rev = 'The Vendor shall indemnify the Purchaser against all Losses.';
    expect(render(wordDiff(orig, rev, SG))).toBe(
      'The Vendor shall indemnify [-and hold harmless -]the Purchaser against [-any and -]all Losses[- whatsoever-].',
    );
  });

  it('leaves the word a longer formulation keeps outside both marks', () => {
    const orig = 'all Losses howsoever arising in connection with the Properties.';
    const rev = 'all Losses arising directly or indirectly out of or in connection with the Properties.';
    expect(render(wordDiff(orig, rev, SG))).toBe(
      'all Losses [-howsoever -]arising {+directly or indirectly out of or +}in connection with the Properties.',
    );
  });

  it('splits a punctuation change from the capitalisation beside it', () => {
    expect(render(wordDiff('on a continuing basis; it survives.', 'on a continuing basis. It survives.', SG))).toBe(
      'on a continuing basis[-;-]{+.+} [-it-]{+It+} survives.',
    );
  });

  it('declines to split a rewritten phrase that shares only a space', () => {
    // `out of any` → `from each` shares nothing but the space between the words.
    // Splitting there buys the reader nothing and turns one editing act into two
    // marks (P3), so the coarse block stands.
    expect(render(wordDiff('arising out of any breach', 'arising from each breach', SG))).toBe(
      'arising [-out of any-]{+from each+} breach',
    );
  });

  it('declines a candidate that strikes a word here and re-inserts it there', () => {
    // A transposition. Every fine-grained rendering of it makes the reader prove
    // two words identical across marks (P2), so the coarse block stands and
    // phrase-scale move detection — not fragmentation — is the right answer.
    const toks = (text: string, op: WordDiff['op']): WordDiff[] =>
      tokenize(text, SG).map((t) => ({ op, text: t }));
    const block = [...toks('indemnify, defend', 'del'), ...toks('defend, indemnify', 'ins')];
    expect(refineReplaceBlocks(block, SG)).toEqual(block);
  });

  it('never adds to or drops from either side (ADR-0008)', () => {
    const pairs: Array<[string, string]> = [
      ['The Vendor shall indemnify and hold harmless the Purchaser against any and all Losses whatsoever.', 'The Vendor shall indemnify the Purchaser against all Losses.'],
      ['pay S$1,250,000 within 30 days of the Completion Date', 'pay S$1,500,000 within 45 days of the Longstop Date'],
      ['under paragraph 2 of Schedule 4 hereof', 'under paragraph 3 of Schedule 4 hereof'],
      ['The notice shall specify the claim.', 'The notice shall specify the nature of the claim in reasonable detail.'],
    ];
    for (const [a, b] of pairs) {
      const out = wordDiff(a, b, SG);
      expect(side(out, 'del')).toBe(a);
      expect(side(out, 'ins')).toBe(b);
    }
  });

  it('returns its input untouched when there is no replace block to refine', () => {
    // Names what this actually pins: the early exit and the no-block path, NOT
    // the reconstruction guard (whose branch the fuzz below proves unreachable).
    const single: WordDiff[] = [{ op: 'eq', text: 'unchanged' }];
    expect(refineReplaceBlocks(single, SG)).toBe(single);
    const noBlock: WordDiff[] = [
      { op: 'eq', text: 'the Vendor ' },
      { op: 'del', text: 'promptly ' },
      { op: 'eq', text: 'shall pay' },
    ];
    expect(refineReplaceBlocks(noBlock, SG)).toEqual(noBlock);
  });

  it('never adds to or drops from either side, over a generated corpus (ADR-0008)', () => {
    // `refineReplaceBlocks` and `detectTransposition` both end in a
    // reconstruct-both-sides guard that returns the input unchanged on any
    // mismatch. That branch is a safety net: in correct code it is unreachable,
    // so it cannot be provoked without injecting a fault. What IS testable is
    // the property it protects — every pass in the chain only ever
    // RE-PARTITIONS the same two texts — so pin that over generated pairs
    // rather than pretending to exercise the bail-out.
    //
    // Measured coverage of this corpus, so the claim above is checkable rather
    // than asserted: 400 of 400 pairs distinct, `refineReplaceBlocks` sees a
    // replace block in ~23, and 23 renderings come out in the relocation shape
    // `detectTransposition` produces. Both earlier numbers were worse — 128
    // distinct pairs and ZERO transposition rewrites — which is why the
    // generator draws from the LCG's high bits and builds every fifth pair as a
    // transposition by construction.
    // Deterministic LCG: same corpus on every run, no Math.random.
    // Drawn from the HIGH bits — an LCG's low bits have short periods, and
    // `seed % n` on them collapsed 400 draws into 128 distinct pairs.
    let seed = 20260904;
    const rnd = (n: number): number => {
      seed = (seed * 1103515245 + 12345) & 0x7fffffff;
      return (seed >>> 15) % n;
    };
    const vocab = [
      'the', 'Vendor', 'Purchaser', 'shall', 'must', 'promptly', 'indemnify', 'defend',
      'hold', 'harmless', 'against', 'any', 'and', 'all', 'Losses', 'whatsoever',
      'arising', 'out', 'of', 'in', 'connection', 'with', 'Clause', '4.2', 'S$1,250,000',
      '30 days', '7.5%', '1 January 2026', 'Schedule', 'paragraph', 'reasonable', 'detail',
    ];
    const sentence = (len: number): string[] =>
      Array.from({ length: len }, () => vocab[rnd(vocab.length)]!);

    // Every fifth pair is a TRANSPOSITION BY CONSTRUCTION — `A sep B` → `B sep A`
    // over a noise-only connective, the shape `detectTransposition` exists for.
    // Random slot edits alone entered that pass 31 times in 400 pairs and made
    // it rewrite ZERO times, so the comment above used to claim coverage the
    // corpus did not provide.
    const SEP = [', ', ' / ', '; ', ' - '];
    const transposition = (): [string, string] => {
      const head = sentence(1 + rnd(4)).join(' ');
      const tail = sentence(1 + rnd(4)).join(' ');
      const sep = SEP[rnd(SEP.length)]!;
      let A = sentence(1 + rnd(2)).join(' ');
      let B = sentence(1 + rnd(2)).join(' ');
      if (A === B) B = `${B} further`;
      if (rnd(2) === 0) [A, B] = [B, A];
      return [`${head} ${A}${sep}${B} ${tail}`, `${head} ${B}${sep}${A} ${tail}`];
    };

    for (let n = 0; n < 400; n++) {
      let a: string;
      let b: string;
      if (n % 5 === 0) {
        [a, b] = transposition();
      } else {
        const base = sentence(4 + rnd(14));
        const rev = [...base];
        // 1–3 random edits: delete, insert, replace, or swap two adjacent spans.
        for (let e = 0; e < 1 + rnd(3); e++) {
          if (rev.length === 0) break;
          const at = rnd(rev.length);
          const kind = rnd(4);
          if (kind === 0) rev.splice(at, 1);
          else if (kind === 1) rev.splice(at, 0, vocab[rnd(vocab.length)]!);
          else if (kind === 2) rev[at] = vocab[rnd(vocab.length)]!;
          else if (at + 1 < rev.length) [rev[at], rev[at + 1]] = [rev[at + 1]!, rev[at]!];
        }
        a = base.join(' ');
        b = rev.join(' ');
      }
      const out = wordDiff(a, b, SG);
      expect(side(out, 'del')).toBe(a);
      expect(side(out, 'ins')).toBe(b);
    }
  });

  it('keeps an atomic value whole and keeps its kind', () => {
    const out = wordDiff('pay S$1,250,000 today', 'pay S$1,500,000 today', SG);
    expect(out.find((d) => d.op === 'del')?.text).toBe('S$1,250,000');
    expect(out.find((d) => d.op === 'del')?.atomic).toBe('amount');
    expect(out.find((d) => d.op === 'ins')?.atomic).toBe('amount');
  });
});

/**
 * The third branch of the acceptance rule — a shared character AFFIX (ADR-0040).
 *
 * `recoversWord`, `selfDiffing` and `overlapSplit` all compare token identity
 * through `keyOf`, so none of them can see that `faith` and
 * `faith-][+unilateral,` share five characters: to all three they are simply
 * different tokens. Accepting left the duplication visible and declining left
 * the block coarse, which showed it anyway — both branches produced the same
 * defect, and there was no third option.
 *
 * Each guard below has an ablation: remove that guard alone and the named test
 * fails. All of them were run.
 */
describe('refineReplaceBlocks — recovering a shared character affix', () => {
  const render = (diffs: readonly WordDiff[]): string =>
    diffs
      .map((d) => (d.op === 'eq' ? d.text : d.op === 'del' ? `[-${d.text}-]` : `{+${d.text}+}`))
      .join('');

  it('renders a word with new material appended as one insertion', () => {
    // A known problem shape: a word struck and re-inserted VERBATIM
    // because a clause was appended after it. The reader reads `faith` twice to
    // learn that nothing happened to it.
    const out = wordDiff(
      'It shall act in good faith at all times.',
      'It shall act in good faith and with reasonable care at all times.',
      SG,
    );
    // The insertion is anchored after the space by `slideEditBoundaries`
    // (ADR-0028), which is the better placement — the mark opens on a word.
    expect(render(out)).toBe('It shall act in good faith {+and with reasonable care +}at all times.');
    expect(out.some((d) => d.op === 'del')).toBe(false);
  });

  it('renders new material PREPENDED as one insertion too', () => {
    // The affix may be at either end; a shared suffix is the same defect
    // mirrored, and `overlapSplit` cannot see it either.
    const out = wordDiff('It shall pay the fee.', 'It shall pay the reasonable fee.', SG);
    expect(render(out)).toBe('It shall pay the {+reasonable +}fee.');
  });

  it('LEAVES a word that merely gained a suffix as a whole-word swap', () => {
    // Refused by BOTH affix guards (the cut is mid-word AND the remainder is
    // no new word), so it is not an ablation for either on its own — see the
    // two tests below. The reviewer graded this shape GOOD and wrote
    // the reason down in corpus case C02 — *"Whole word Loss/Losses swapped,
    // not `Loss` + inserted `es`"*. Rendering it as `Loss{+es+}` is the
    // character-level interleave P5 exists to forbid, and it would be bought
    // with a metric against a human verdict. The guard is what separates "a
    // word gained an ending" from "a clause was appended after a word".
    const out = wordDiff(
      'The Vendor is liable for any Loss arising.',
      'The Vendor is liable for any Losses arising.',
      SG,
    );
    expect(render(out)).toBe('The Vendor is liable for any [-Loss-]{+Losses+} arising.');
  });

  it('LEAVES a word that merely gained a POSSESSIVE as a whole-word swap', () => {
    // ABLATION: `addsNewWords`, and the only shape that isolates it. An
    // apostrophe is not a letter or a digit, so `splitsBetweenWords` is happy to
    // cut here — without `addsNewWords` this renders `Purchaser{+'s+}`. That is
    // the same "a word gained an ending" shape C02's reviewer graded good, one
    // punctuation mark along, and the branch exists for appended CLAUSES, not
    // for inflections. `Purchaser` → `Purchaser's agent` still splits, because
    // there the remainder really is new words.
    const out = wordDiff(
      'Notice to the Purchaser was served.',
      "Notice to the Purchaser's was served.",
      SG,
    );
    expect(render(out)).toBe("Notice to the [-Purchaser-]{+Purchaser's+} was served.");
  });

  it('LEAVES a whole-word swap when the appended material is a suffix AND new words', () => {
    // ABLATION: `splitsBetweenWords`. The mixed case, and the one `addsNewWords`
    // alone cannot reach: it scans the WHOLE remainder, so the single space in
    // `es and Damages` licenses a boundary that still lands mid-way through
    // `Losses`. Without this guard the pair below renders `Loss{+es and
    // Damages+}` — the same character-level interleave the sibling guard exists
    // to forbid, arrived at from the other side. A reader cannot see where
    // `Loss` ends and the new material begins.
    const out = wordDiff(
      'The Vendor is liable for any Loss arising.',
      'The Vendor is liable for any Losses and Damages arising.',
      SG,
    );
    expect(render(out)).toBe(
      'The Vendor is liable for any [-Loss-]{+Losses and Damages+} arising.',
    );
  });

  it('LEAVES a whole-word swap when a word grows an infix before the new words', () => {
    // ABLATION: `splitsBetweenWords`, second shape. `Vendor` is wholly a prefix
    // of `Vendorial agent`, and the remainder `ial agent` contains a space, so
    // `addsNewWords` passes it. The boundary is still inside `Vendorial`.
    const out = wordDiff(
      'The Vendor shall pay the fee.',
      'The Vendorial agent shall pay the fee.',
      SG,
    );
    expect(render(out)).toBe('The [-Vendor-]{+Vendorial agent+} shall pay the fee.');
  });

  it('LEAVES a whole-word swap when the cut would land mid-word at the OTHER end', () => {
    // ABLATION: `splitsBetweenWords`, mirrored. `holder` is wholly a SUFFIX of
    // `the Shareholder`, and the remainder `the Share` has two spaces, so
    // `addsNewWords` passes it. Without the boundary test this renders
    // `{+the Share+}holder` — a span that begins mid-construction, which is the
    // machine residue the whole-side-containment guard is justified by:
    // containment held, the word boundary did not.
    const out = wordDiff(
      'A holder may vote at the meeting.',
      'A the Shareholder may vote at the meeting.',
      SG,
    );
    expect(render(out)).toBe('A [-holder-]{+the Shareholder+} may vote at the meeting.');
  });

  it('still splits at a possessive or a hyphen, which a reader reads as a boundary', () => {
    // The other side of `splitsBetweenWords`: `'` and `-` are inside the
    // tokenizer's word-run class (`WORD_CHAR`), so a boundary test written on
    // that class would refuse these too and withdraw the branch's whole point.
    // A reader seeing `Purchaser{+'s agent+}` reads two things, not a word cut
    // in half — which is why the guard tests for a LETTER OR DIGIT across the
    // boundary, not for a word-run character.
    const out = wordDiff(
      'Notice is given to the Purchaser under this clause.',
      "Notice is given to the Purchaser's agent under this clause.",
      SG,
    );
    expect(render(out)).toBe(
      "Notice is given to the Purchaser{+'s agent+} under this clause.",
    );
  });

  it('does not recover a meaningless shared character at a word edge', () => {
    // ABLATION: whole-side containment. `Vendor` and `Purchaser` share a final
    // `r`, and recovering it would render `[-Vendo-]{+Purchase+}r` — a span
    // that begins mid-construction, which is machine residue and costs the
    // reader a second pass (P4). The branch fires only where one side's text is
    // WHOLLY a prefix or suffix of the other's.
    const out = wordDiff(
      'The Vendor shall indemnify the Purchaser.',
      'The Purchaser shall indemnify the Purchaser.',
      SG,
    );
    expect(render(out)).toBe('The [-Vendor-]{+Purchaser+} shall indemnify the Purchaser.');
  });

  it('never cuts an atomic value in half (P5)', () => {
    // ABLATION: `splitTokensAt`'s atomic refusal. `30` is wholly a prefix of
    // `30 days`, and the affix boundary therefore falls INSIDE the duration
    // token — the one place P5 forbids a boundary. Without the refusal this
    // renders `30{+ days+}`, splitting a period, and `atomicIntact` goes false
    // across the corpus. A period, amount, date or reference value changes
    // whole or not at all, and no readability gain is worth cutting one.
    const out = wordDiff(
      'The Vendor shall pay 30 today.',
      'The Vendor shall pay 30 days today.',
      SG,
    );
    expect(render(out)).toBe('The Vendor shall pay [-30-]{+30 days+} today.');
    expect(out.find((d) => d.op === 'ins')?.atomic).toBe('duration');
  });

  it('runs to a fixed point, because an accepted re-diff can leave a new block', () => {
    // The reason one pass is not enough, and the reason a single pass still
    // showed `faithfaith` after the affix branch first landed. The re-diff here
    // IS accepted — it recovers `good ` — and its own output contains a FRESH
    // replace block, `faith` beside `faith-][+unilateral,`, which only the
    // affix branch can answer and which a single pass never looks at.
    //
    // The revised text carries redline notation as literal body text. It is source
    // content, not a rendering artefact.
    const out = wordDiff(
      'It shall act in good faith today.',
      'It shall act in [-good faith-][+unilateral, unchallengeable+] today.',
      SG,
    );
    // `good faith` survives as unmarked text on BOTH sides — nothing is struck
    // and re-inserted to say that a bracket was added around it.
    expect(render(out)).toBe(
      'It shall act in {+[-+}good faith{+-][+unilateral, unchallengeable+]+} today.',
    );
    expect(out.some((d) => d.op === 'del')).toBe(false);
  });

  it('is length-preserving on both sides, as ADR-0008 requires', () => {
    const a = 'The Vendor shall act in good faith and pay the Loss within 30 days.';
    const b = 'The Vendor shall act in good faith and fair dealing and pay the Losses within 30 days.';
    const out = wordDiff(a, b, SG);
    const side = (s: 'ins' | 'del'): string =>
      out.filter((d) => d.op === 'eq' || d.op === s).map((d) => d.text).join('');
    expect(side('del')).toBe(a);
    expect(side('ins')).toBe(b);
  });
});
