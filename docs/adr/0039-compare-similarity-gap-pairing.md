# ADR-0039: Compare — inside a gap, pair paragraphs by similarity, not by position

- **Status:** Accepted
- **Date:** 2026-10-01
- **Relates to:** ADR-0026 (change-complexity demotion — the gate this shares); ADR-0041 (its contiguity escape); ADR-0032 (merge/split — the pass this must not starve); ADR-0008 (categorise, never suppress); ADR-0017 (the PDF path's own fuzzy re-pairing, left untouched); feature 08 (`spec/features/08-compare.md` §5.1)
- **Scope boundary:** G1 (maintainer decision, 2026-09-11) — better paragraph *pairing* is inside the scope register's Compare row; fuzzy *moved-and-edited* detection stays v2 (the scope register, "Deferred"). This ADR stays on the pairing side of that line, and §3 below says exactly where the line is.

## Context

Level 1 (`diff/structural.ts`) anchors on paragraph hashes and runs an LCS between anchors. What is
left between two matches (a *gap*) went to `emitGap`, which paired it **by position**: the k-th
unmatched original with the k-th unmatched revised, surplus as delete/insert.

One deleted or inserted paragraph inside a gap therefore shifts every pairing after it. ADR-0026
then does what it is specified to do with a mismatched pair (Dice below 0.5): it demotes the pair
to a whole-paragraph delete + insert. The reader sees a clause struck and re-typed beside a clause
whose real edit was two words. This was reported on 2026-10-01 on an ordinary definitions
list: a definition deleted directly above an edited definition.

The readability work had found the same cause twice, and
measured that the engine already computed the right answer and never asked: on one real-world pair (not published) the
positional pairs scored Dice 0.31, while the counterparts beside them scored 0.95 and 0.96.

## Decision

Inside **one gap**, `emitGap` now applies three rules in order.

1. **Lock.** Every positional pair that renders inline today keeps its partner. "Renders inline
   today" is not re-derived: `lockPositional` runs the gap's positional ops through the two passes
   that decide it before Level 2, ADR-0032's merge/split re-pairing and then ADR-0026's demotion,
   and locks every positional pair that comes out of them as a `replace`.
2. **Align.** In each stretch between locked pairs, choose the **order-preserving** set of pairs
   whose Dice is ≥ `replaceSimilarityThreshold` and whose total Dice is maximal (a
   Needleman–Wunsch-shaped dynamic programme; pairs never cross).
3. **Fall back.** What the alignment leaves is paired by position, except that a record **held for
   a move** (below) is never paired. A stretch the alignment does not touch is emitted exactly as
   positional pairing emitted it.

Candidacy in rule 2 uses the Dice half of the demotion gate with the same arithmetic (`bagDice` /
`diceFromOverlap` in `diff/replace-threshold.ts`, the one implementation), so a pair chosen in rule
2 can never be demoted afterwards.

Three kinds of record are kept out of rule 2:
- **A table unit** (`ParaRecord.table`, ADR-0027). It is one alignment unit with its own redline
  author, so a stretch that holds one stays positional.
- **A paragraph with no content tokens** (`* * *`). Two of these score Dice 1 against each other and
  would outbid a real counterpart they cross, so it is never a candidate.
- **A record whose text appears verbatim on the other side.** It is **held for a move**: never a
  candidate, left unpaired by the fallback beside re-aligned pairs, and refused as an orphan by
  ADR-0032's merge/split pass unless positional pairing's own merge pass absorbed it
  (`GapPairing.mayAbsorb`, which `detectMergesAndSplits` takes as an optional predicate). With all
  three, Level 3's hash-based move detection (§5.4) finds it exactly as before. The first two
  alone were not enough (§1, part (b)).

The work is bounded per stretch (`MAX_ALIGN_WORK`, `MAX_ALIGN_CELLS`) and across the whole
comparison (`MAX_ALIGN_WORK_TOTAL`), all below. Token bags are interned and cached once
per record, and a Dice upper bound (`2·min(a,b)/(a+b)`) skips pairs that cannot reach the threshold.

Applied by `runComparePipeline` only, so Word-vs-Word, Text-vs-text and both batch modes get it.
`structuralDiff` takes the pairing as an optional argument and stays pure and config-free.

### 1. What is guaranteed, exactly

The requirement was that no other part of the redline get worse. A rule that only
*measured* well could not promise that, so the guarantee is structural. It has four parts:

- **(a) Locked pairs.** A positional pair that renders inline today (after merge/split and
  demotion) keeps its partner.
- **(b) Held records.** A record whose text appears verbatim on the other side is never chosen by
  the alignment and never paired by the fallback beside a re-aligned pair. Where positional
  pairing had already paired it (a locked pair, or a stretch the alignment leaves alone) it stays
  exactly as it was. Nor is it absorbed by ADR-0032's merge/split pass as an orphan unless
  positional pairing's own merge pass absorbed it too (`GapPairing.mayAbsorb`, from the lock
  probe's record of what it merged). Either way its route to Level 3's move detection is
  unchanged.
- **(c) A stretch the alignment does not touch** is emitted with the same pairs, op types and
  ranges as positional pairing emitted it. Not byte-for-byte: op ids, and the authored `w:id`s
  numbered from them, shift once an earlier gap is re-aligned. They are bookkeeping, not
  rendering.
- **(d) Document order.** Every op list walks both documents in order, as positional pairing's
  always did, because §5.5 authors the redline by walking it. The fallback beside re-aligned pairs
  emits each held record in its place, not after the pairs. A generated-document test (600 pairs
  mixing moves, repeated boilerplate, edits, insertions and deletions) checks this on every op.

So every record that changes partner is one that positional pairing sent to whole-paragraph
delete + insert, and is not held for a move. It is either paired with a partner that passes the
same Dice threshold, or absorbed by a merge or split that ADR-0032's own guards accept (next list). A locked pair keeps its
partner, but it can **gain** such an absorbed record.

**What it does not promise, stated rather than discovered:**

- **Op order inside a re-aligned stretch changes.** A re-paired delete can now precede its
  neighbour's replace. ADR-0032's merge/split pass matches only an orphan *after* its replace
  (its correction 1 removed the mirrored orderings as unreachable), so a merge presented
  orphan-first is not detected. Because the lock already keeps every merge positional pairing
  produced, that shape arises only among records that were churn before. It is an accepted limit
  and does not re-open ADR-0032.
- **New adjacency can create a merge or split that did not exist before.** The absorbed
  orphan is always a record that was churn before and is not held for a move. The host can be a re-aligned pair, or a
  **locked** pair whose positional neighbour was freed by the alignment (re-validation found one:
  a locked definition absorbing the following sentence it had been joined with). The host keeps
  its partner and gains the orphan, and the result renders better than the churn it replaces.
  ADR-0032's own guards and its reconstruction check apply unchanged, and the widened op cannot
  be demoted. The joined Dice is at least the mediant of a remainder Dice ≥ 0.5 and a containment
  ≥ 0.85.

**The first version of this section was false, and the reason matters.** It said the lock was "a
pair that survives ADR-0026's gate", asked of the positional pair on its own. Two spec-conformance review
passes built the counterexample independently: a merge or split whose **head** was
also edited. The head's pair fails the gate alone and renders inline only because ADR-0032 scores
the **joined** text. The alignment took the merge apart: the head was struck whole and its edited
wording re-inserted inside the tail's paragraph, which is the repeated-text failure ADR-0032
removed. The same review found that a verbatim move could be stolen (part (b)), and that the
section claimed merge/split "handles all four orderings", which ADR-0032's correction 1 had already
made untrue. **A lock that re-derives a downstream decision is a second copy of that decision.
This lock runs the decision instead.**

**And part (b) was false a second time, through the same downstream pass.** It held a
record back from the alignment and the fallback, and said nothing about Stage 2.25. Re-pairing can
put a held record directly after a replace it never followed (a re-aligned host, or a locked host
whose next positional partner was re-aligned away), the merge/split pass absorbed it as an orphan,
its verbatim twin lost its partner, and the move became a plain insertion, in a gap that had no
churn at all (so part (c) broke with it: Level 3 is global). A differential fuzz found it once in
3,000 generated pairs. The merge pass now asks `mayAbsorb`, which refuses a held record unless the
lock probe's own merge pass absorbed it. With that, 9,000 generated pairs over three seeds lose no
move and change nothing outside a churned gap. **Holding a record back from one pass says nothing
about the passes after it; ask which downstream pass can still take it.**

**And the fix for that carried its own defect** (re-validation). The first fallback
beside re-aligned pairs emitted its pairs first and its held records after. A held record that
came before a fallback pair on its own side therefore landed behind it, and reject-all rebuilt
that side out of order. No ADR-0008 check could see it, because every record was still present.
Part (d) and the generated-document test exist because of it. The same pass found the work budget
admitting millions of cells for two-token paragraphs (7.2 s, ~0.9 GB), now capped by
`MAX_ALIGN_CELLS`.

### 2. Why order-preserving and gap-local

Pairing across an anchor (an `equal` paragraph) would pair a paragraph with one that sits on the
other side of unchanged text. That is a **move** with an edit in it: The scope register's deferred fuzzy
moved-and-edited detection. A gap is bounded by matched paragraphs on both sides, and an
order-preserving alignment inside it cannot cross one. On a private real-world contract pair, all 7
whole-paragraph deletions left over after this change have their best counterpart on the far side
of a matched paragraph. They stay delete + insert, as the scope register requires.

### 3. The line against v2

**This is pairing:** choosing, among paragraphs already known to be unmatched and in the same
order-bounded gap, which original corresponds to which revised. **This is not move detection:** no
op is typed `move`, nothing is paired across a matched paragraph, and nothing is badged as
relocated.

## Rejected alternatives (each measured on the prototype, 2026-10-01)

Measured over eight whole-document pairs (private real-world contract pairs and six
synthetic smoke-test contracts) and the 30 synthetic readability cases, rendered before and after.

- **The alignment without the lock.** 39 of 41 unchanged; the contract improved, and **C21 (two
  paragraphs merged) regressed**: marked characters 13 → 136, redundancy 0 → 0.85. Both source
  paragraphs are contiguity-confident against the merged paragraph. The alignment preferred the
  longer one (higher Dice), so ADR-0032 never saw the host it needs and the merge rendered as a
  struck paragraph plus a long insertion. The lock keeps C21's positional pair, and its rendering
  is byte-identical to before.
- **Locking by asking the demotion gate about each positional pair alone.** This was shipped first
  and was wrong; see §1. Two independent review passes reproduced the edited-head merge and split.
- **Leaving a whole stretch positional when it holds a would-be move.** This kept every move, but
  lost a re-pairing on a private contract pair (`wholeParaChurn` 20 → 21). In a table, a severity
  cell that moved shared a stretch with an edited cell beside it. Holding only the record itself
  keeps the move *and* the re-pairing; both are pinned by tests.
- **Contiguity as a route into rule 2.** Identical output to Dice-only on every case measured, and
  it would let a short heading that appears verbatim inside a long clause be chosen as that
  clause's partner. No benefit, some risk. Contiguity is used only in the lock (rule 1), where it
  is already the shipped gate.
- **Greedy cross-gap re-pairing, as `fuzzyAnchorPdf` does for PDFs (ADR-0017).** That is move
  detection (§2) and is v2 for Word-vs-Word.
- **Loosening ADR-0026's 0.5 threshold.** The demotion is the amplifier here, not the cause. It is
  correct on a mismatched pair, and loosening it would put "diff salad" back on the genuine
  rewrites it exists to clean up.
- **Kuhn–Munkres over the gap (an earlier sketch, reusing `batch/pairing/`).** An
  unconstrained assignment can cross, which brings back the move problem; restricted to
  non-crossing pairs, it is the DP above.

## Accepted limits

- **A stretch above `MAX_ALIGN_WORK` or `MAX_ALIGN_CELLS` keeps positional pairing.**
  Both limits apply per stretch between locked pairs.
  - The cell cap (10⁶) bounds memory: the DP holds two `Float64Array`s of m·n cells.
  - The work budget is 10⁸ cost units, where cost is cells × (`ALIGN_CELL_OVERHEAD` (32) + mean
    distinct tokens). It is fitted to measured time per cell: ~90 ns at 1 token, ~256 ns at 40,
    ~566 ns at 100, ~1,860 ns at 600.
  - **Two budgets came before it, and both were wrong.**
    - A bare cell count could not see paragraph length: 1,000² cells of 600-token paragraphs took
      **24 s**.
    - Cells × tokens alone could not see the fixed per-cell cost: 8,600² two-token paragraphs
      took 7.2 s and ~0.9 GB, and a reviewer measured 4–19 s on other shapes.

  Measured end to end through `runComparePipeline`, **each shape in a fresh process**. Every
  shape is one unanchored gap with every paragraph edited and one inserted at the top. Measuring
  shapes back to back in one process inflated later rows several-fold through memory pressure;
  one such run read 7.6 s for a row that measures 3.0 s fresh.

  | Gap | Tokens / paragraph | Before ADR-0039 | Now |
  |---|---|---|---|
  | 12,000 × 12,001 | 1 | 2.26 s, 54 inline | 2.02 s, **11,946** inline |
  | 6,000 × 6,001 | 3 | 0.61 s | 0.67 s (over the cell cap: positional) |
  | 1,900 × 1,901 | 40 | 0.32 s | 0.45 s (over the cell cap: positional) |
  | 999 × 1,000 | 100 | 0.36 s | 0.47 s (over budget: positional) |
  | 400 × 401 | 40 | 0.08 s, 0 inline | 0.20 s, **400** inline |
  | 600 × 601 | 140 | 0.31 s, 0 inline | 0.81 s, **600** inline |
  | 700 × 701 | 300 | 1.05 s | 1.35 s (over budget: positional) |
  | 999 × 1,000 | 600 | 2.25 s | 2.97 s |

  Where a shape falls back to positional pairing, the remaining overhead (up to ~0.7 s, in the
  last row) is the lock's probe. It runs ADR-0032's and ADR-0026's passes over the gap, and the
  pipeline runs them again at Stages 2.25 and 2.5. Sharing those results between the two would
  remove the overhead. That was judged not worth the coupling for shapes no real document has
  produced: the eight document pairs' gaps are tens of paragraphs, not hundreds. Stage 2 has no
  yield point inside it.

  Both budgets above bound one **stretch**, and a document of many large gaps paid them once per
  gap: ten gaps of 999 × 1,000 forty-token paragraphs, separated by unchanged paragraphs, took
  6.3 s of Stage 2 CPU against 0.15 s positional, linear in the number of gaps. A
  third budget, `MAX_ALIGN_WORK_TOTAL` (3 × 10⁸ in the same units), bounds the whole comparison:
  once the stretches aligned so far have spent it, later stretches keep positional pairing.
  Deterministic, because gaps are visited in document order. The same ten-gap shape then spent
  ~3.0 s aligning on a machine at load average ~58 (5 gaps ≈ 2.4 s, so the cap binds near four
  such gaps). The lock probe is not under this budget; it is linear in the gap and ~0.09 s per
  1,000-paragraph gap.
- **Two algorithms compute one multiset intersection.** `bagDice` uses string maps; the
  alignment's `score` uses interned, sorted ids, measured 3–4.5× faster on all-pairs scoring
  (999² cells at 100 tokens: 3,154 ms against 971 ms). The formula is shared (`diceFromOverlap`)
  and a property test pins that the two agree on every pair scored. One formula, two counting
  methods, measured reason recorded: a deliberate exception to the project principles, not an oversight.
- **A wrong positional pair that already passes the gate is kept** (rule 1). Sibling clauses that
  share boilerplate (`… passes to the Buyer on …`) can score Dice 0.67 against each other. The lock
  keeps such a pair rather than trading it for a better-scoring one, because breaking a pair that
  renders inline today is the very thing this design rules out. Pinned by a test that states what it
  becomes if the lock is ever relaxed. **Measured: it occurs 0 times** across the eight document
  pairs (Level 1, locked against unlocked).
- **A record held for a move whose other-side copy is not actually free** (its twin is matched
  `equal` elsewhere) is still held. It is conservative, and costs a re-pairing only where a held
  record's neighbours could have been aligned *through* it.
- **The 4,000-pair generalisation sweep cannot see this change.** It runs `wordDiff` on single
  paragraphs and never reaches Level 1. The evidence is the eight document pairs and the corpus.
- **Headers/footers/footnotes (`parts-diff.ts`) and the PDF path are unchanged.**

## Consequences

- On a private contract pair: `wholeParaChurn` 27 → 20, `alignmentRate` 0.73 → 0.76, marked
  characters 13,529 → 12,483, reconstruction true on both sides. Every other measured rendering is
  byte-identical.
- Spec §5.1 now describes gap pairing as similarity-first with a positional fallback.
