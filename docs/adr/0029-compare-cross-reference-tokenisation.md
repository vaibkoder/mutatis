# ADR-0029: Compare — cross-reference tokenisation (recognise the lower-case noun, mark only the value)

- **Status:** Accepted
- **Date:** 2026-09-04
- **Relates to:** ADR-0008 (Compare categorises, never suppresses); ADR-0026 (change-complexity threshold); ADR-0028 (edit-boundary anchoring); feature 08 (`spec/features/08-compare.md` §5.3)

## Context

ADR-0028 turned redline readability from a matter of taste into a measured one. The
30-case corpus in `tests/fixtures/readability-corpus.ts` now runs the real pipeline and
scores each rendering, and the score that reproduced the human grading is **redundancy** —
the share of a struck span that is word-for-word identical to the inserted span beside it,
i.e. text the reader reads twice to learn that nothing happened to it (P1/P2 in
`docs/compare-readability-principles.md`).

Running that measure over cross-reference cases exposed three defects in the `clauseRef`
atomic matcher, which interact:

**1. The matcher was case-sensitive.** `config.referenceTokens` carries the profile's
display spelling — `Clause`, `Schedule`, `paragraph`. In-text references in E&W and SG
drafting are routinely lower-case (`the limitations in clause 9.13`), and the sentence-
initial and in-text forms differ within a single document. A lower-case reference was
therefore not recognised as a reference at all: its value was not atomic (P5 unenforced
exactly where numbers matter most), the op classified as `substantiveText` rather than
`crossRefRenumber` (so P10 triage put a renumbering in the substantive bucket), and it
could not join a renumbering cascade.

**2. The atomic span was too coarse.** The whole match — `Clause 4.2`,
`paragraph 2 of Schedule 4` — was one indivisible token. A renumbering therefore struck
and re-inserted the noun along with the number: `[-Clause 9.13-]{+Clause 9.14+}`. The
reader reads "Clause" twice to learn it did not change. Measured: C23 redundancy 0.55,
C24 0.80 — the two worst scores in the corpus after the genuinely-hard C11/C21.

**3. The two defects were cancelling.** The unrecognised lower-case form rendered
*better*: falling through to ordinary word tokenisation happened to mark only the number.
Adding the `i` flag on its own would have made the corpus worse, not better. Case-
insensitivity and split emission had to land together.

## Decision

**A cross-reference is matched as a whole span but emitted as several tokens.** The
reference noun, the whitespace and the `of` connective are ordinary word tokens; only each
reference **value** carries `atomic: 'clauseRef'`. P5 makes the number indivisible; it does
not make the noun in front of it borrowable.

To support that:

- **Tokens carry their kind.** `tokenizeTagged(text, config): Token[]` — `{ text, atomic? }`
  — is the primary tokenizer; `tokenize()` is a thin text-only wrapper for callers that do
  not need kinds. A bare value (`9.13`) is not recognisable out of context, so consumers
  read `Token.atomic` instead of re-deriving the kind with `atomicKindOf`.
- **Sentinels are keyed on kind + text** in `diff/word.ts` (the atomic kind, a NUL
  separator, then the token text), so a reference value never aligns with an
  identical-looking ordinary word.
- **Word runs break only at real word boundaries.** A case-insensitive noun match must not
  turn `subparagraph 3` into `sub` + `paragraph 3`, or strip the `Sub` from `Subclause 4.2`.
- **`anchor.ts` trusts the tag.** `splitTail` refuses to split any token the tokenizer
  tagged atomic, rather than re-deriving the kind from the token text.
- **`cascade.ts` recovers the noun from context.** The noun now sits *outside* the mark, so
  `deriveMember` reads it back off the unchanged text immediately before the mark
  (`refKindBefore`) and normalises it to the profile's own spelling — `Clause` and `clause`
  group into one cascade rather than two.

## Consequences

Measured over the 30-case corpus, before → after:

| case | before | after |
|---|---|---|
| C16 (multiple refs, one clause) | 0.25 | **0** |
| C23 (single renumber) | 0.55 | **0** |
| C24 (`paragraph N of Schedule M`) | 0.80 | **0** |
| C29 (renumber + substantive insert) | 0.25 | **0** |
| corpus total redundancy | 5.04 | **3.19** |

Mark count is unchanged at 82 across the corpus, `atomicIntact` holds for all 30 cases, and
no case regressed. C16 and C23 now classify `crossRefRenumber` instead of `substantiveText`,
fixing the P10 triage half of the defect as well as the P1 rendering half.

**ADR-0008 is untouched.** This is tokenisation and marking granularity. Nothing is hidden,
merged away or judged equivalent; every change still appears, and cascade collapse remains
the presentation-only hint it always was.

**P5 still holds.** The *value* is indivisible. What changed is that the noun is no longer
part of the value.

## Alternatives rejected

- **Add the `i` flag alone.** Measured worse (context, defect 3).
- **Lower-case `referenceTokens` in the jurisdiction profiles.** Loses the display spelling
  the cascade labels and navigation rows use, and pushes a diff-engine problem into the
  ADR-0002 locale config.
- **Trim shared prefixes/suffixes inside a mark at output time.** A general fix for a
  general problem, and still on the table for the non-reference redundancy cases (C03, C11,
  C12, C21) — but for cross-references it re-derives at the output stage something the
  tokenizer already knows, so it is the wrong layer here.

## Open

- `Subclause` is not in `referenceTokens`, so `Subclause 4.2`'s value is not tagged. That is
  a jurisdiction-config question, recorded rather than fixed.
- ~~C24 still classifies as `definedTerm`: the Title-Case-multiword rule sees `Schedule 4`.~~
  **Closed by ADR-0033.** The cause was more general and much worse than `Schedule 4`:
  the rule tested the whole PARAGRAPH, and `TITLE_CASE_MULTIWORD` matches a
  sentence-initial capital plus a capitalised word, so `The Vendor shall …` matched. Since
  row 4 outranks row 5 and `computeCascades` only collapses ops classified
  `crossRefRenumber`, the renumbering cascade silently did not fire on any clause opening
  that way. The predicate now tests the change, per §6 row 4's own wording.

---

## Correction — 2026-09-12: the possessive apostrophe is not a quote delimiter

Found by running the engine over three real share-purchase agreements (ADR-0038)
after this ADR had been Accepted for a week. It corrects a rule this ADR
recorded, in the file this ADR owns, so it is written here rather than opened as
a new decision.

### What was wrong

The `definedTerm` matcher this ADR left in place read:

```js
{ kind: 'definedTerm', re: new RegExp('"[^"]+"|“[^”]+”|\'[^\']+\'|‘[^’]+’', 'y') }
```

The `'[^']+'` alternative matches **everything between any two apostrophes**.
In English legal drafting that means **everything between any two possessives**
— and §5.2's `normaliseText` folds `’ → '` before the tokenizer ever runs, so
**every Word document hits it**. One ordinary clause tokenises to a single
atomic `definedTerm` the length of a sentence. Corpus case C33 is the shape,
written for this suite (the documents it was found on stay outside the
repository, and their words are not reproduced anywhere in it):

```
"…the Landlord and the Landlord's agent shall be given full notice of the
  claims and reports made to the Insurer, the brokers of, and any surveys
  commissioned by, the Tenant and the Tenant's insurer."

  → atomic definedTerm, 158 characters, running from the first possessive to the second
```

The span is atomic, so **P5 forbids splitting it**, so changing `full` to
`reasonable` — one word — strikes the whole 158-character span and re-inserts
it six characters longer, nearly all of it verbatim. The worst instance on the
real fixtures restated **355 identical characters** to say that a clause had
been appended.

This is the single largest-payoff defect found in the readability benchmarking, and it had
been invisible for the entire life of the workstream: the 30-case corpus
contained **seven possessives and never two in one paragraph**, and the
4,000-pair generalisation sweep's generator emits none at all.

### The correction

Both single-quote alternatives are guarded on **both** sides. An opening quote
never follows a word character, and a closing quote is never followed by one:

```
(?<![\p{L}\p{N}])'[^']+'(?![\p{L}\p{N}])
```

Each guard is load-bearing for a case the other cannot reach, which is why both
are present and why each is ablated separately in
`tests/unit/compare.diff.word.test.ts`:

- The **closing** guard rules out the singular possessive — the span
  `'s agent … the Tenant'` ends on an apostrophe followed by `s`.
- The **opening** guard rules out the **plural** possessive, which is the harder
  half: `Vendors'` ends *on* its apostrophe, so `' rights and the Purchasers'`
  closes on an apostrophe followed by a **space** and the closing guard passes
  it. Only the opening guard sees it.

The curly pair carries the same guards, because `’` is the typographic
possessive and the defect is identical on any text §5.2 has not folded.

The guards need `u`, so the matcher is now `'yu'`, and the character class
matches `hasReadableContent` in the same module (the project principles) rather than
the ASCII-only `[A-Za-z0-9]` the word-run branch uses — a party name may carry
an accent.

### One supporting change, without which the guard misfires

`startsAtomic` took a **slice** (`text.slice(j)`) and matched at offset 0.
Against a slice a lookbehind sees nothing before position 0 and **passes
vacuously**, so the word-run loop would break at a possessive apostrophe, and
the main loop — matching at the real offset, where the lookbehind correctly
fails — would then emit the apostrophe as a lone punctuation token. `Vendors'`
would tokenise as `Vendors` + `'`, splitting a word a reader sees as one word
(the very property `readerTokens` is built on). It now takes the whole text and
an offset. Sticky matching at an offset is identical for every other pattern in
the set; reverting it alone fails two of the six new tests.

### What it changed, measured

**The 30-case corpus did not move at all** — total redundancy **0.11**, marks
**89**, `atomicIntact` and `numbered` green throughout, and the same **8 of 9**
reviewer ideals (C03, C08, C11, C12, C16, C23, C24, C29; C21 unchanged at
`false`). **This ADR's own prediction that C09, C22 and C26 would re-classify
did not happen**, and no classification test moved. Recorded because the original plan
expected a corpus move and treat it as the fix working: the
honest result is that the corpus never contained the defect at all.

**The 4,000-pair generalisation sweep is byte-identical**: offenders **66**,
mean redundancy **0.0388**, mean marks **2.3287**, 3,675 swept. The generator
has no possessives.

**Where it does show is the two places that were built to see it.** The three
corpus cases added for this defect (C31–C33):

| case | `redundancy` | `charRedundancy` before | after |
|---|---|---|---|
| C31 two possessives, `shall`→`may` between them | 0 | **0.70** | **0.33** |
| C33 the 158-character span, `full`→`reasonable` | 0 | **0.79** | **0.25** |
| C32 `Statement`→`Statements` (defect A, untouched here) | 0 | 1.00 | 1.00 |

And the real documents, **with ADR-0038's contamination caveat attached**:
`charOffendingBlocks` **13 → 11** on pair 3 and **6 → 5** on pair 4 — the four
possessive-spanned blocks scoring 0.81–0.86 are gone. `charRedundancyMean` rose
(0.97 → 0.99, 0.87 → 0.93) **because the mean is over a shrinking set**: what
remains is the pure `faithfaith` shape scoring 1.00, which is defect A and is
not addressed here. A rising mean over a falling count is not a regression, and
the count is the number to read.

`wholeParaChurn` and `alignmentRate` are unchanged on all three pairs, as they
must be — `ParaRecord.hash` is computed over `normText`, which tokenisation does
not touch.

### Rejected here

- **Delete the single-quote alternatives outright.** Simpler, and it would
  withdraw a real rule: a single-quoted defined term is ordinary E&W drafting.
  The guards narrow the rule to the cases where it is right.
- **A lookbehind alone.** Measured: it leaves `'the Vendor'`s matched as a
  defined term the document does not have.
- **A lookahead alone.** Measured: it passes the plural possessive, which is the
  shape C01 of the corpus is built on.
- **Lower-casing or otherwise pre-processing apostrophes before tokenising.**
  Pushes a diff-engine problem into §5.2 normalisation, where it would also
  change hashing and therefore paragraph matching.
