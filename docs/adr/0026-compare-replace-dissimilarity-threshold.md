# ADR-0026: Compare — change-complexity threshold (demote a too-dissimilar `replace` to clean delete+insert)

- **Status:** Accepted
- **Date:** 2026-06-17
- **Relates to:** ADR-0008 (Compare categorises, never suppresses); feature 08 (`spec/features/08-compare.md` §5.1, §5.3)

## Context

Mutatis's Compare pipeline aligns paragraphs at Level 1 (anchors + segmented LCS), then for every index-aligned, hash-unequal pair emits a `replace` op and runs a Level-2 word diff (`diff/word.ts`) to produce an inline run-level `<w:ins>`/`<w:del>` redline.

> **Note (2026-10-01, ADR-0039).** "Index-aligned" was true when this was written and is no longer the whole story. Inside a gap, Level 1 now locks the positional pairs that render inline today, decided by running this ADR's pass (and ADR-0032's, which precedes it) over the gap's positional ops, and pairs the rest by similarity using this ADR's Dice threshold and arithmetic. This gate was the *amplifier* of the positional pairing's failure, not its cause: a deleted paragraph in a run of changes shifted every pair after it, and this gate then correctly demoted each mismatched pair to whole-paragraph churn. The gate and its thresholds are unchanged.

That inline word diff is the *right* rendering when the two paragraphs are genuinely similar — a changed amount, a `shall → may` swap, a reworded phrase. It is the *wrong* rendering when the aligned pair is a near-total rewrite: diff-match-patch then produces a confusing token-by-token interleave ("diff salad") where a handful of incidental shared words ("the", "of", "shall", a shared defined term) are shown as `equal` islands inside a sea of insertions and deletions. The reader cannot tell what actually happened.

Conventional legal comparison tools treat this as a solved, named problem: when the ratio of matched text to changes falls below a threshold, they show the entire paragraph as deleted and the new one as inserted rather than as many small word changes, in service of redline **readability**. (Same idea is applied to whole tables; that table variant is **out of scope for this ADR** — see Rejected alternatives.)

Mutatis had no such gate: it word-diffs every `replace` unconditionally. This is the one substantive "smartness" gap identified against conventional redline style; alignment quality, move rendering, numbering handling and the §6 classifier were already sound.

## Decision

Add a **change-complexity threshold** to the Word-vs-Word pipeline. Between Level-1 structural diff and the Level-2 word diff, run a pass (`diff/replace-threshold.ts → demoteDissimilarReplaces`) that:

1. For each `replace` op, computes a **paragraph similarity** between the two paragraphs' `normText` — the **Sørensen–Dice coefficient over content tokens** (the same atomic-aware `tokenize` the word diff uses, whitespace tokens excluded). Range `[0,1]`: `1` = identical token multiset, `0` = disjoint.
2. If similarity **`< compare.replaceSimilarityThreshold`** (new config key, **default `0.5`**), demotes the `replace` into a **`delete` of the whole original paragraph + an `insert` of the whole revised paragraph**. Otherwise the `replace` is left untouched and word-diffs inline as before.

The pass runs **only in the Word-vs-Word pipeline** (`runPipeline`). The PDF pipeline (`runPipelinePdf`) is unchanged: it deliberately moves in the *opposite* direction, re-pairing reconstruction-noise `delete`+`insert` into `replace` via its fuzzy-anchor pass (ADR-0017), and must not be fought by a dissimilarity gate.

`compare.replaceSimilarityThreshold = 0` disables the pass entirely (similarity is always ≥ 0, so nothing is ever demoted).

### Why `0.5`, and why conservative

The default demotes only when **a majority of the combined content differs**. This is deliberately conservative so that informative inline diffs are preserved:

- A changed amount / date / percentage, or a `shall → may` modal swap, sits in an otherwise-shared sentence → Dice well above `0.5` → stays `replace`, so the atomic-token and obligation-language classification still fire on the inline word diff.
- A clause rewritten wholesale, or a short heading fully replaced (`Definitions` → `Interpretation`) → Dice below `0.5` → clean delete+insert.

Firms can tune the key (e.g. raise it to demote more aggressively for noisier drafts).

## This does not violate the never-suppress invariant (ADR-0008 / the project principles)

Demotion changes only **how a change is rendered**, never **whether** it is shown. A demoted pair emits the **entire original paragraph as a `<w:del>` and the entire revised paragraph as a `<w:ins>`** — strictly *more* revision markup than the inline diff it replaces, never less. No text is dropped, hidden, or filtered; both versions remain fully present in the redline and individually accept/rejectable in Word. It is the opposite of suppression: a near-total rewrite is shown as a clean, complete swap instead of a misleading partial-match.

The one second-order effect — a demoted pair is classified `structural` (delete+insert) rather than carrying the inline word diff's atomic/obligation tags — is acceptable and bounded: `structural` changes are **always surfaced and never filterable** (§6 priority 1), and the conservative `0.5` default keeps small in-place obligation/value edits as `replace` precisely so their classification is retained. This matches the spec's existing stance that classification is heuristic and "can never hide" a change (§12.5).

## Rejected alternatives

- **Do nothing (status quo).** Rejected: leaves the diff-salad readability failure on every wholesale-rewritten clause — the exact gap against conventional redline style, and the thing this change exists to fix.
- **Demote inside `structuralDiff` at op-creation time.** Rejected: `structuralDiff` is pure and config-free and is shared by the Word and PDF pipelines and by many unit tests. Threading config in and branching by pipeline there is more invasive than a small, separately-testable post-pass, and would wrongly apply the gate to the PDF path.
- **Use diff-match-patch's own output to decide (count ins/del vs eq characters after diffing).** Rejected as the default path: it requires running the expensive word diff first just to discard it on demoted pairs, defeating the lazy/large-document performance design (§9.10). A cheap token-Dice pre-check is O(n) and decides before any word diff runs.
- **Also fold in the table-level complexity threshold** (show a whole table as insert/delete when its cell-level change complexity is high). Rejected **for this ADR/slice** — it is a separate, broader change touching `diff/tables*.ts` with its own existing tests to reconcile. Deferred as a follow-up; recorded here so the boundary is explicit, not forgotten.
- **A higher default threshold (e.g. 0.7) to demote more eagerly.** Rejected as the default: it risks demoting genuine in-place edits (and losing their atomic/obligation classification) for marginal readability gain. `0.5` is conservative and configurable; firms can raise it.
