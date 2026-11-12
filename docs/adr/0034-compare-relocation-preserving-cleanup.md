# ADR-0034: Compare — a slot relocation must survive `diff_cleanupSemantic`

- **Status:** Accepted
- **Date:** 2026-09-04
- **Relates to:** ADR-0008 (Compare categorises, never suppresses); ADR-0028 (edit-boundary anchoring); ADR-0029 (cross-reference tokenisation); ADR-0030 (replace-block refinement); ADR-0031 (adjacent transposition); feature 08 (`spec/features/08-compare.md` §5.3)

## Context

The corrected generalisation sweep (3,675 generated clause pairs) breached the
0.60 redundancy floor 158 times, and **149 of those were slot relocations** —
one span struck where it used to sit and re-inserted a few words later, with the
text in between untouched. 52 were pure permutations: the same word multiset
struck as inserted, nothing added or removed. The same sweep showed
`detectTransposition` (ADR-0031) changing **nothing at all** — mean redundancy
and offender count byte-identical with it removed — so the one pass that exists
to render a phrase-scale move does not generalise past its curated case.

### The relocation is already correct, and then it is destroyed

`diff_main` renders the shape properly. Measured, before any cleanup:

```
The Purchaser [-shall -]notify{+ shall+} each Indemnified Person …
```

`diff_cleanupSemantic` then eliminates the equality. Its rule — remove an
equality no longer than the edits on both sides of it — is measured with
`String.length`, and `diff/word.ts` runs it in **sentinel** space where one
character is one **token**. `notify` counts 1 against `shall` + a space
counting 2, so the equality goes and both edits grow by what they swallowed:

```
The Purchaser [-shall notify-]{+notify shall+} each Indemnified Person …
```

`notify` is now struck and re-inserted for nothing. That is P1 ("never widen a
mark to swallow text that is word-for-word identical on both sides") and P2
("never make the reader diff the redline") in a single move. It is the same
root cause ADR-0030 documented — `String.length` in token space — striking a
different shape.

DMP does have an overlap pass (`diff_commonOverlap_`) that can restore an
equality after a fusion, but **it is not what kept this defect hidden** — a
second review measured that, and an earlier draft of this ADR had the mechanism
wrong. The rescue needs the middle equality to be at least as long as the span,
while the fusion needs it no longer than the span, so the two can only coincide
when they are exactly equal. It never runs on the curated corpus at all.

The defect stayed hidden for a much plainer reason: **neither of the corpus's two
relocation cases is ever fused.** C08's connective `", "` is two tokens against
one-token edits, so the elimination rule does not apply; C25's middle equality is
a whole sentence. For both, the cleanup's output is byte-identical to the raw
`diff_main` output. The one-token headline example above **is** fused — which is
exactly what `tests/unit/compare.diff.relocate.test.ts` now catches.

### Nothing downstream can undo it, by design

`refineReplaceBlocks` re-diffs each fused block on real tokens, and its second
acceptance rule refuses any candidate that strikes a word in one mark and
re-inserts it in another. Every finer rendering of a relocation does exactly
that, so the refinement declines it — correctly, and ADR-0030 says so outright:
"a transposition stays coarse and is left to move detection".

`detectTransposition` is that move detection, and it matches `del(A) ins(B)
eq(mid) del(B) ins(A)` in the **coalesced** diff — a five-part shape the fusion
has already erased. It never fires on this input.

So the relocation is visible in exactly one place: **before the cleanup**.

## Decision

`diff/relocate.ts → cleanupPreservingRelocations` replaces the bare
`dmp.diff_cleanupSemantic(diffs)` call in `diff/word.ts`. It applies DMP's own
cleanup **everywhere except across a relocation**: the diff is cut at each
protected window and the cleanup is run on each remaining stretch on its own, so
every other over-fusion is still cleaned up exactly as before.

A window is a `del` run, one `eq` run and an `ins` run (either order). Guarded
on **structure alone** — no vocabulary, no semantics, no equivalence scoring
(ADR-0008):

- **the two spans are verbatim identical.** The same tokens in the same order,
  ignoring only pure-whitespace tokens, compared **as sentinels** — so the
  comparison includes each token's atomic kind and a reference value can never
  match an identical-looking ordinary word (ADR-0029). This is exact sequence
  equality, never a similarity score, and it is the guard that makes "this
  moved" a fact about the text rather than a judgement about it;
- **the relocated span carries readable content** — a letter or a digit. A
  stray comma "relocating" is not something a reader can watch move, and a
  contentless span is the artefact the readability metric itself had to be
  corrected for.

The relocation keeps rendering as an ordinary `del` here plus `ins` there. It is
**not** re-labelled `w:moveFrom`/`w:moveTo`: that alternative was measured and
is rejected below.

Two things the pass deliberately does **not** have:

- **No bound on how far the span moved.** DMP's rule is self-limiting — it only
  eliminates an equality no longer than the edits flanking it — so where the
  untouched text between the two ends is long the fusion never happens and there
  is nothing to protect. A distance bound would be a knob with no failure mode
  behind it.
- **No first-match-wins de-duplication.** Protection only decides where the
  cleanup is applied, so a run taking part in two windows costs nothing. An
  earlier draft justified this with a synthetic rendering that de-duplication
  made worse; a second review could not reproduce that case, and measured
  de-duplication as **inert** — 441 windows against 442, identical redundancy
  and marks, 0 pairs changed. The claim is withdrawn. The rule stands on the
  weaker and true ground that it is one fewer moving part, and it is recorded
  here as **not pinned by any test**.

## Consequences

Measured over the 3,675-pair generalisation sweep:

| chain | mean redundancy | blocks ≥ 0.60 | mean marks |
|---|---|---|---|
| **as shipped (this ADR)** | **0.0388** | **66** | **2.3287** |
| without `cleanupPreservingRelocations` | 0.0635 | 158 | 2.3233 |
| without `detectTransposition` | 0.0388 | 66 | 2.3287 |
| without `refineReplaceBlocks` | 0.1422 | 397 | 2.3290 |
| without both refinements | 0.1693 | 495 | 2.3224 |

The disclosed ratchet `KNOWN_OFFENDERS` drops **158 → 66**. Scored pair by
pair against the same chain with the pass ablated: **93 pairs improved, 0 got
worse**, 3,582 unchanged. Marks: 16 pairs gained one or more, 1 lost one, **net
+20 over 3,675 pairs** (8,538 → 8,558; mean 2.3233 → 2.3287, +0.2%) — the
intended direction, and the same trade ADR-0030 made: honest small marks beat
one confusing block. The gain comes entirely from **not fusing** — the untouched
middle stays an equality — and not from any scoring exemption: the metric only
looks at a struck span sitting immediately beside an inserted one, which a
protected relocation never is.

**The curated 30-case corpus does not move at all** — redundancy 0.11, 83
marks, 8 of 9 recorded ideals, C21 still the one unmet. Instrumented, the
window fires on exactly **2 of the 30** (C08, the transposition; C25, two
sentences transposed in one paragraph) and **both render byte-identically with
the pass ablated**: DMP's own overlap heuristic already rescued those two, C08's
because its spans are one token each. So the defect this ADR fixes is one the
curated corpus does not contain — it was found by measurement on unseen input,
which is what the readability workstream exists for, and it is why the corpus could not have found
it.

`detectTransposition` still changes nothing the sweep measures. It is not
removed here — it is pinned by C08 and its own unit tests. An earlier draft
claimed this pass makes its shape rarer; that contradicts this ADR's own
Context, and is false in the other direction too, since protection can only
*preserve* equalities. Measured firings on the sweep: **0 with the pass, 0
without** — no effect in either direction.

### What is left — and why it needs no code

66 offenders remain. A later measurement of what they actually are corrects the
first characterisation of them, which was wrong in a way that mattered:

| the 66 survivors | count |
|---|---|
| a phrase genuinely relocated, nothing else changed | **4** |
| a relocation **plus another edit inside the same span** | **62** |

Of the 62: 24 gained a parenthetical, 19 had a slot substituted or dropped, 12
gained a qualifier, 3 are `shall`→`must`, 4 other. Marks: min 2, **median 2**,
max 6 — already the smallest rendering a replaced phrase can have. Where the two
sides differ in content and not only in order, one replaced phrase is a
defensible rendering, not a defect.

More importantly, **the generator's revised texts for these are not English.**
Its slot-swap edit permutes adjacent fragments with no regard for grammar:

```
shall not The Company hold harmless for any Liability the Vendor ...
The Guarantor shall not the Vendor hold harmless howsoever arising ...
is obliged to The Purchaser indemnify for any Liability each Indemnified Person ...
```

So `KNOWN_OFFENDERS` is a **regression tripwire under stress**, not an estimate
of what a reader will meet. It must not be quoted as a user-facing defect count,
and the residual 66 is **not a backlog**. The curated 30 — lawyer-plausible and
reviewer-approved — remains the realistic measure of readability.

The sibling approach that lowers the tripwire to 55 (a bounded exception to
ADR-0030 rule 2) was prototyped, measured and **declined**: its 11 extra cases
are scrambles, it amends a settled ADR, it broke C12 until confined, and 6 of
its 9 guards are unpinned by its own account. It is preserved unmerged on branch
`worktree-agent-ad56f38ca80a53681` (`c4ebbb1`) should a real user complaint ever
justify the risk.

## Alternatives rejected

- **Re-label the protected pair as `w:moveFrom`/`w:moveTo`.** **Decided
  against.** `spec/features/08-compare.md` §5.3 already states the rule at this
  scale — a phrase-scale relocation "stays an ordinary insert/delete pair … P7
  does not need `w:moveFrom`/`w:moveTo` inside one sentence" — and §5.4 reserves
  the native markup for paragraph-scale verbatim moves. ADR-0031 settled the
  same point. `WordDiff.op` is `'eq' | 'ins' | 'del'` (`model.ts:76`), so there
  is no phrase-scale move op to set: re-labelling would be a model change plus
  new OOXML authoring, not a label change. Closing the question here rather than
  leaving it open, because it was already closed elsewhere.

  One caveat, stated because an earlier draft of this ADR had it backwards.
  `tests/fixtures/readability-metrics.ts` *does* exempt `moveFrom` from the
  redundancy measure — but that would **not** flatter this pass's score today: a
  protected relocation is never adjacent to its counterpart, so it scores 0
  either way, which is what the "0 pairs whose score changed" measurement above
  says. The exemption is a hazard for any **future, looser** detector, not an
  artefact of this one.
- **Fix it after the fusion, in `refineReplaceBlocks`.** Structurally
  impossible without reversing that pass's second acceptance rule (P2), which
  ADR-0030 measured and which C07 and C12 pin. The information is gone by then:
  which words were one span is no longer recoverable from the fused block.
- **Drop `diff_cleanupSemantic` entirely.** Measured twice already (ADR-0028,
  ADR-0030): same redundancy, more fragmentation, and it re-introduces the
  `[-a-] {+b+}` shape. This ADR keeps the cleanup and removes it from exactly
  one place.
- **Accept any `del(X) eq(mid) ins(Y)` window, not only a verbatim one.**
  Measured: it drives the sweep to redundancy **0.0212** and **17** offenders —
  the best numbers of any variant — while marks jump **2.33 → 3.48**, a 50%
  increase, because every ordinary rewrite that happens to straddle a shared
  word is now fragmented. C07 and C12 both fail. This is the metric-optimal and
  reader-hostile option, the same trap ADR-0030 recorded, and it is the reason
  the identity guard is exact equality: the guard **costs** the objective
  function, which is how it is known not to be tuned to it.
- **Bound the distance the span may move**, or **de-duplicate overlapping
  windows** — both measured to change nothing; see above.
