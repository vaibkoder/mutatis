# ADR-0038: Compare — a readability instrument that can see a real document

- **Status:** Accepted
- **Date:** 2026-09-11
- **Relates to:** ADR-0008 (categorise, never suppress — the invariant this instrument asserts on real input); ADR-0028/0029/0030/0031/0032/0034 (the passes measured by the yardstick this ADR corrects); ADR-0037 (precedent for a dated in-place correction); feature 08 (`spec/features/08-compare.md` §5.3–5.5)

## Context

The redline readability workstream is governed by one number. `redundancy` — the
share of a struck span that is word-for-word identical to the inserted span
beside it — reproduced a reviewing lawyer's grading of a 22-case corpus exactly,
which is why it became the objective function. Six ADRs were measured against it
and it was driven from 5.04 to 0.11, the corpus floor, with eight of the nine
renderings the reviewer wrote down produced exactly.

The engine was then run over three real-world share-purchase-agreement
pairs (private test documents, not published) and the output read `faithfaith`, `StatementsStatements`,
`AdjustmentAdjustment` — word for word the defect `redundancy` exists to
measure, from an engine recorded at the metric's floor.

Both things were true. The corpus was at 0.11 and the redline was duplicating
words. What reconciles them is that **the corpus contained 30 synthetic cases and
zero `.docx` files**, and the instrument turned out to be blind in two
independent ways, each of which this ADR now records.

**Blindness 1 — the measure exempts the failure it most needs to see.** The
within-paragraph half exempts `isCleanSwap` (a paragraph with no `eq` segment)
and the cross-paragraph half (ADR-0032) scores a wholly-struck paragraph only
against insertions drawn from paragraphs that are *not* wholly inserted. Put
together: when alignment fails completely and a counterpart pair renders as
delete-everything-here plus insert-everything-there, `redundancy` scores it
**0**. The objective function improves as the alignment degrades. Measured on
the private test pairs, this is not a theoretical concern — `pair2` pairs **one of
sixteen** original paragraphs and deletes the other fifteen beside their
counterparts' insertions, and scores **0** on every redundancy measure in the
suite.

**Blindness 2 — the yardstick is derived from the thing it measures.**
`readerTokens` builds its idea of a word out of the production tokenizer. So
when that tokenizer emits `faith` and `faith-][+unilateral,` as two unrelated
tokens, the metric also sees two unrelated tokens, finds nothing shared, and
scores **0** — on the exact rendering that exposed the defect. `Statement` →
`Statements` scores 0 for the same reason. A yardstick built out of the thing it
measures cannot find a defect in that thing.

The plan's own decision log had already written down the rule this situation
calls for, from ADR-0032:

> a metric that cannot see a defect is evidence of nothing, and the first move on
> a "this already scores fine" case is to check the instrument.

## Decision

**The instrument is extended before any engine code changes, and real documents
are measured on their own scale with their own companions.**

### 1. Real documents enter the corpus as a case kind, not as more cases

`ReadabilityCase` becomes a discriminated union — `{kind:'text', original,
revised}` | `{kind:'docx', originalPath, revisedPath}` — and `renderCase`
branches into `extractParts` → `preprocessSide`, the two calls the Word-vs-Word
driver already makes. `runComparePipeline` has always taken `ParaRecord[]` rather
than strings, so `extractTextRecords` was only ever one adapter; there was no
missing capability here, only missing wiring.

Real cases are **not** members of `READABILITY_CASES` and are **never summed into
the synthetic total**. The `5.04 → 0.11` history stays comparable, and the
synthetic figures are unchanged by the widening: 0.11 total redundancy, 83 marks,
8 of 9 ideals, byte-identical.

### 2. `redundancy`'s max-aggregation is kept for synthetic cases and not used as
the headline for real ones

A synthetic case isolates one editing act, so its max *is* its score. Over a
26-paragraph contract a max stops discriminating — one bad block pins it near 1.0
— so real cases report `redundancyMean` and `redundancyP95` over their replace
blocks, and `offendingBlocks` (the 0.60 "the reader must diff this block himself"
floor, the same one the 4,000-pair sweep counts) in absolute terms and per 1,000
words.

### 3. Two companions measure what `redundancy` exempts

- **`wholeParaChurn`** — original paragraphs inside a whole-paragraph `delete`
  op that sits immediately beside a whole-paragraph `insert` op. Ratchets down.

  **Counted in paragraphs, not ops** (corrected on review, before first push).
  The op-level count is not monotone under a ratchet: a change that merely
  merged three adjacent unpaired deletes into one op would drop the figure from
  3 to 1 without pairing a single paragraph, and the ratchet would read that as
  progress. Multi-paragraph `delete` ops are reachable — `demoteDissimilarReplaces`
  copies a widened `origRange` onto the op it demotes — so this is not
  hypothetical. None of the three real pairs produces one today, which is why
  the published figures below are **unchanged** by the correction: 15 / 6 / 5
  before and after, still matching the planning probe.
- **`alignmentRate`** — the share of original paragraphs matched to a counterpart
  (`equal`/`replace`/`move`) rather than deleted outright. Ratchets up.

Both are read from the **ops**, not from the painted paragraphs. A struck
paragraph does not always render as an all-`del` paragraph, because ADR-0035
revises the paragraph mark too, so counting the rendering measures the renderer
while counting the ops measures the pairing. The distinction is not academic: the
op-level definition reproduces the planning probe's counts exactly (6 on pair 3,
5 on pair 4) and counting the painted paragraphs is one short on both. Reading
the ops and *weighting by paragraph* is the combination that is both faithful to
the probe and monotone.

### 4. A character-level yardstick sits beside the token one, and both are published

`charRedundancy` is the same measure with the longest verbatim **character** run
in place of the token LCS, knowing nothing about tokenisation. It inherits
`redundancy`'s exemptions (clean swaps, contentless spans) deliberately, so that
the **only** difference between the two numbers is the yardstick — and the gap
between them is the signal. `offendingBlocks` is therefore reported on both: on
the real pairs the token count is **0** on all three while the character count is
13 and 6, so publishing only the token count would publish a clean sheet over the
worst thing these documents show.

### 5. `reconstruction` is a hard assert, and it compares modulo the pipeline's own normalisation

`del`+`eq`+`moveFrom` must rebuild the original and `ins`+`eq`+`moveTo` the
revised, across the whole document. This is ADR-0008 on real input and it
outranks every readability finding in the workstream.

It compares **body text**, through the production `normaliseText` and
`stripNumberingPrefix` rather than a second copy of either, because two things
legitimately differ between the authored redline and the records the pipeline
matched on:

- the redline is authored from the source `rawXml`, so a curly apostrophe
  survives where §5.2 folded it to a straight one;
- ADR-0028 restores the **revised** clause number onto an edited paragraph, so a
  paragraph renumbered 1.1 → 1.2 shows only `1.2` and the original's number is
  legitimately absent.

Neither is a lost word. Whether the number survived is `numbered`'s question
(P9), not this one's; this one asks whether a word was lost.

**One consequence of that split needs stating in one place, because each half is
disclosed and the composition is not.** `reconstruction` strips a numbering
prefix from each rendered paragraph, and the source it compares against already
had one stripped in preprocess — so the two cancel exactly, and a redline that
**drops or invents** a leading `1.1 ` / `(a) ` still reconstructs clean. §6 then
sets `numbered` to `null` for docx cases. Net: on a real document, **nothing in
the instrument checks that the clause number survived.** That is a gap in the
instrument, not a defect in the pipeline, and it is deliberate rather than
overlooked — the real fixtures carry no `<w:numPr>` at all, so a string-parsing
`numbered` would be vacuous on them. It is recorded here so that a future
contributor reads it as a known hole rather than as coverage.

### 6. `idealMatch` and `numbered` are `null` for a real case

There is no hand-authored wanted rendering for a 26-paragraph contract, and
**a conventional redline's output is a reference for what a reader expects, not an
assertion target**. Matching another tool's span choices is not evidence that the
redline reads well, and pinning them would make every future improvement look
like a regression. `numbered` opts out because it parses clause prefixes out of
the revised *string*; on these fixtures every number is style-inherited
(`document.xml` carries zero `<w:numPr>`), so the check would be vacuously 1
rather than honest. P9 on real documents is a separate piece of work.

`marks` is **recorded, not pinned**: an exact mark count on a 26-paragraph
contract is brittle, and raising marks is frequently the correct fix.

### 7. The real documents are NOT committed

They stay outside the repository, found through `MUTATIS_REAL_FIXTURES`
(no default; a local folder outside the repository), and every test that needs them
**skips when they are absent** so the suite stays green on a fresh clone and for
anyone who does not hold the fixtures. The absent run reports skips, not passes.

The reasoning is confidentiality, not tidiness. Real documents carry identifying
metadata and client drafting, and the project's own principle is that document
content never leaves the user's machine; committing such a document to a remote is
the same principle from the other end. Visibility is one click and history is
permanent, so a fixture is committed only if it would be safe published. Reference
output from other tools is likewise not committed.

**The measured numbers are not confidential; the documents are.** The numbers are
recorded in this ADR. Permanent, committed regressions become scrubbed
structural derivatives — same style families, same style-inherited numbering
topology, same complex-field usage, same `titlePg`/header topology, new prose, no
`docProps`.

### 8. Every companion is pinned by a test that fails when that measure alone is broken

The standard is `tests/unit/compare.diff.refine.test.ts`, and it is applied by
actually running the mutation, not by asserting that one would fail. One of the
nine mutations written for this ADR **did not** fail on its first run: the
contentless-run fixture struck ` . `, which the block filter removes before the
guard under test is ever reached, so the test passed with the guard deleted. The
fixture was rewritten to strike readable text sharing only a space. That is the
failure mode this clause exists to catch.

## Consequences

Measured 2026-09-11, engine unchanged, **with the contamination caveat below
attached to every figure**:

| | pair2 | pair3 | pair4 |
|---|---|---|---|
| reconstruction (original / revised) | **pass / pass** | **pass / pass** | **pass / pass** |
| `wholeParaChurn` | **15** | **6** | **5** |
| `alignmentRate` | **0.06** | **0.63** | **0.69** |
| `redundancy` (max, token) | **0** | 0.08 | 0.33 |
| `redundancyMean` / `P95` (token) | 0 / 0 | 0.01 / 0.08 | 0.08 / 0.33 |
| `charRedundancy` (max) | **0** | **1.00** | **1.00** |
| `charRedundancyMean` / `P95` | 0 / 0 | **0.97** / 1.00 | **0.87** / 1.00 |
| `offendingBlocks` (token) | 0 | **0** | **0** |
| `charOffendingBlocks` | 0 | **13** | **6** |
| replace blocks · words · marks | 0 · 2 099 · 40 | 13 · 1 721 · 69 | 6 · 2 101 · 50 |

Synthetic corpus unchanged: **0.11 · 83 marks · 8 of 9 ideals**.

Read the table the way this ADR argues it must be read. `pair2` is the
worst-aligned document of the three by a distance and **every number the corpus
is judged by reports it as perfect**; only `wholeParaChurn` and `alignmentRate`
can see it. On `pair3` the token yardstick reads 0.01 mean over thirteen replace
blocks and the character yardstick reads 0.97, and inspection of all thirteen
shows every one is a word struck and re-inserted verbatim. Not one is a
false positive.

**ADR-0008 holds on real input.** Both sides of all three pairs reconstruct
exactly. Nothing is suppressed; what fails is readability, which is the right way
round and is why this is a workstream rather than a recall.

### The caveat that travels with every number above

The supplied "modified" documents are **not ordinarily edited contracts**. They
carry LLM redline notation (`[-old-][+new+]`) and citation blocks
(`[<24-hex>#Pn"…"]`) as literal body text — 21 % of `modified3`'s revised text
and 38 % of `modified4`'s, none of it in the base. So at several sites a word
genuinely occurs twice in the revised document and the engine is partly telling
the truth about a poisoned input. They are well-formed, they parse, they
reconstruct, and they are valid tests — but a baseline measured on them is a
**floor on how bad things are, not a representative reading**, and it must be
quoted with this caveat every time. Ordinarily-edited variants of the same base were obtained afterwards (see the note below).

The caveat must not be read the other way either. **Both word-diff defects
reproduce on clean prose with no annotation anywhere near them**, and are pinned
by synthetic cases in the suite.

### Known limitation, disclosed rather than engineered away

The character yardstick has one false positive by construction: a whole-for-whole
swap of two *similar atomic values* scores high although the rendering is exactly
what P5 mandates. Measured: `1 January 2026 → 1 February 2026` scores **0.64**,
because the two dates share `uary 2026`.

Every way of removing it consults the tokenizer to ask "is this span atomic?" —
and a yardstick that consults the tokenizer is blind to tokenizer defects, which
is the entire reason this measure exists. Masking atomic spans would have hidden
`atomic-tokens.ts`'s sentence-length bogus possessive span (corpus case C33, 158
characters), the worst instance in the brief. So the false positive is kept and disclosed, and the token yardstick —
which has the complementary blindness — is published beside it. Neither number is
trustworthy alone; that is the design. On the three real pairs the false positive
does not fire once.

## Rejected alternatives

- **Fold the real cases into the existing 30-case total.** The brief's preamble
  expected the total to "jump well above 0.11". It does not — `pair2` scores 0 —
  and summing them would both destroy the comparability of the `5.04 → 0.11`
  history and let a catastrophically misaligned document *improve* the headline
  figure. Real cases get their own scale.
- **Keep max-aggregation as the headline for real documents.** One bad block pins
  the max near 1.0 on a long document and it stops discriminating between a
  document with one bad block and one with forty.
- **Commit the real documents so the regression is permanent and portable.**
  Rejected on confidentiality (§7). Scrubbed structural derivatives carry what
  the *tests* need; the real documents are for the *baseline measurement*.
- **Assert against another tool's output.** Its span choices are a
  reference for what a reader expects, not ground truth, and pinning them would
  make every independent improvement read as a regression. Cleaner-looking
  output or matching change counts are not, on their own, evidence of success.
- **Fix `redundancy`'s clean-swap exemption instead of adding companions.** The
  exemption is principled — P2 admits the whole-for-whole swap as the other
  honest rendering, and ADR-0026 demotes to it on purpose — so narrowing it would
  score a deliberate design decision as a defect. The shape is real, so it gets
  its own measure rather than a distorted existing one.
- **Give the character yardstick a minimum run length, or an atomic-span mask,
  to suppress its false positive.** Both were considered and rejected: the mask
  is circular (see the limitation above), and a length threshold is an arbitrary
  knob that would have to be tuned against the very cases it is meant to judge.
  `[-the-]{+they+}` scoring 1.00 is the measure working.
- **Measure `wholeParaChurn` from the rendered paragraphs.** Simpler to read, and
  wrong: it measures the renderer rather than the pairing, and undercounts by one
  on two of the three real pairs.

---

## Note — 2026-09-12: the fixture set was replaced, and two recorded limits change with it

The three `modified2/3/4` pairs this ADR was measured on are no longer in use.
A clean private pair replaced them — `v1` against `v1.1` of the same
contract, ordinary drafting edits (a real-world contract pair, not published) —
together with a conventionally styled redline of the same two files, which served
as an external reading reference. Recorded here because two things
this ADR states are now different.

### 1. The contamination caveat is LIFTED

§ "The caveat that travels with every number above" applied to documents that
carried LLM redline notation and citation blocks as literal body text — 21 % and
38 % of the revised text. Nothing measured on the new pair carries that caveat:
its figures are a reading, not a floor. **The old baselines (`wholeParaChurn`
15 / 6 / 5, `alignmentRate` 0.06 / 0.63 / 0.69) cannot be reproduced** — the
documents are gone — and survive only in this ADR.

The `pair2` specimen goes with them. **The finding it demonstrated stands**: a
document that pairs 1 of its 16 paragraphs and still scores 0 on every
redundancy measure in the suite is why this instrument exists, and it is why
§ 3's two companions are published first. What is lost is the *specimen*, not
the conclusion; the shape itself is still pinned on hand-built renderings in
`compare.readability.metrics.test.ts`, which is where a measure's own tests
belong.

### 2. The character yardstick's false-positive class is broader than § "Known limitation" says

That section discloses one shape: a whole-for-whole swap of two **similar
atomic values** (`1 January 2026 → 1 February 2026`, 0.64). The general rule is
wider, and the new pair shows it plainly. The measure is longest-common-substring
over the shorter side, so it scores **1.00 for any whole-word swap where one
side is a substring of the other** — which includes:

- a plural or other inflection (`Statement` → `Statements`, corpus C32);
- a **single-character typo correction** (`each` → `eac`, `by` → `b`) — five of
  these on the new pair, and its entire `charOffendingBlocks` count;
- two similar atomic values, the case already disclosed.

In every one of them the whole-word swap is the **correct** rendering — it is
what the reviewing lawyer graded good in C02, and it is the conventional rendering
of a typo correction. So `charOffendingBlocks` on a real pair is a
**ceiling to ratchet against, never a defect count to drive to zero**, and the
test that pins it says so.

The reasoning in § "Known limitation" is unchanged and is why the measure is not
"fixed": every way of removing the false positive consults the tokenizer to ask
"is this span atomic?", and a yardstick that consults the tokenizer is blind to
tokenizer defects, which is the entire reason this one exists.

### 3. What the benchmark showed that the instrument could not

The benchmark found one defect the instrument *did* catch and one it structurally
cannot:

- **Caught, and fixed — ADR-0041.** Two paragraphs that each lost a large
  contiguous block were demoted to whole-paragraph delete + insert.
  `wholeParaChurn` read 2 and `redundancy` 0.8, so the instrument was pointing
  straight at it.
- **Not caught, and it cannot be.** A reference redline shows two numbering changes
  the engine emitted nothing for, because the documents carry no `w:numPr` at all
  and the clause numbers come from the style chain. Every metric in this ADR
  reads the **authored redline**; a change that produces no op produces no
  markup, and no measure over markup can see an op that was never emitted. This
  is a detection gap, not a rendering one, and it was handled as a separate piece of work. Recording it
  here because it marks the instrument's outer boundary: it measures how well
  the engine renders what it found, and says nothing about what it failed to
  find.

