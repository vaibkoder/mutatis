# ADR-0033: Compare — the defined-term class tests the change, not the neighbourhood

- **Status:** Accepted
- **Date:** 2026-09-04
- **Relates to:** ADR-0008 (Compare categorises, never suppresses); ADR-0013 (defined-term recognition, feature 01); ADR-0029 (cross-reference tokenisation — recorded this as **Open**); feature 08 (`spec/features/08-compare.md` §6 rows 4–5)

## Context

Spec §6 row 4 defines the defined-term class as *"Change **touches** a quoted
term or a Title-Case multi-word term, **or alters the definition paragraph** of
one."* The implementation tested the **paragraph**:

```ts
return paras.some((p) => TITLE_CASE_MULTIWORD.test(p.normText));
```

`TITLE_CASE_MULTIWORD` is `/\b[A-Z][a-z]+(?:\s+[A-Z][a-z]+)+\b/`, which matches a
**sentence-initial** capital followed by a capitalised word. So `The Vendor
shall …` matches — and that is how most clauses in an SG or E&W contract open.
The class therefore fired on almost every edit in almost every document,
whatever the edit actually was.

Mis-labelling would have been the smaller half. Row 4 **outranks** row 5
(cross-reference renumbering), and `classify/cascade.ts` collapses only ops
whose **primary** classification is `crossRefRenumber`:

```ts
if (op.classification !== 'crossRefRenumber') continue;
```

So the renumbering cascade — the flagship §6 "Collapsed" behaviour, spec'd in T3
as *"47 +1 Clause shifts → one CascadeAnnotation"* — **silently did not fire** on
any clause whose paragraph contained a Title-Case bigram. Reproduced through the
production pipeline on twelve consistent `+1` shifts, where the only difference
between runs is prose the edit never touched:

| paragraph | classification | cascades |
|---|---|---|
| `subject to Clause 4.N above` | `crossRefRenumber` | **1** (affectedCount 12) |
| `The Vendor shall comply with Clause 4.N above` | `definedTerm` | **0** |
| `The Purchaser shall observe Clause 4.N in full` | `definedTerm` | **0** |
| `each Permitted Encumbrance under Clause 4.N applies` | `definedTerm` | **0** |

The existing unit suite passed throughout, for two reasons worth recording. The
47-shift cascade test hand-builds synthetic paragraphs (`Para 0`, `Para 1`) that
carry no Title-Case bigram, so it exercised the one shape where the bug is
invisible. And one test — *"a Title-Case multi-word phrase in para text flags
definedTerm"* — asserted the buggy behaviour by name, which is what made this
look settled rather than broken. ADR-0029 had already recorded the symptom
(C24 classifying `definedTerm`) as **Open**, attributing it to `Schedule 4`; the
sentence-initial `The Vendor` cause is more general and much worse.

## Decision

Both limbs of row 4, as written, and nothing else:

1. a changed `WordDiff` is a quoted defined-term atomic; **or**
2. the changed **text** carries a Title-Case multi-word term; **or**
3. the touched paragraph is the one that **defines** a term — a quoted term
   followed by a definitional verb (`DEFINITION_PARA`), where an edit anywhere
   in the paragraph is an edit to the definition.

Limb 3 keeps a paragraph-level test exactly where the spec puts one, and it is
deliberately narrow: ADR-0013's broad unquoted-candidate recognition belongs to
feature 01, and borrowing it here would re-introduce the over-matching this
replaces.

**No priority order changes**, and no class is removed from the taxonomy. This is
the row-4 *predicate* being made to match its own definition.

## Consequences

Triage across the readability corpus, before → after:

| case | act | before | after |
|---|---|---|---|
| C24 | compound cross-reference renumbered | `definedTerm` | **`crossRefRenumber`** |
| C11 | redundant doublet collapsed | `definedTerm` | **`substantiveText`** |
| C12 | phrase replaced with a longer formulation | `definedTerm` | **`substantiveText`** |
| C21 | two paragraphs merged | `definedTerm` | **`substantiveText`** |
| C06 | standard-of-fault word deleted | `definedTerm` | **`substantiveText`** |
| C02 | singular → plural | `definedTerm` | **`substantiveText`** |
| C22 | defined term renamed throughout | `definedTerm` | `definedTerm` (unchanged — the change touches the term) |
| C26 | quoted defined term renamed | `definedTerm` | `definedTerm` (unchanged) |
| C09 | `Third Party Claim` → `Third-Party Claim` | `definedTerm` | `definedTerm` (unchanged — the term is inside the change) |
| C01 | apostrophe moved (singular → plural possessive) | `definedTerm` | **`substantiveText`** |
| C04 | capitalisation only — a word becoming a defined term | `definedTerm` | **`substantiveText`** |

C01 and C04 were omitted from the first version of this table. Both move for the
same reason as C02 and C06: the surrounding paragraph carries a Title-Case
multi-word phrase, so the paragraph-level test matched, while the change itself
touches no defined term. Verified in both directions — restoring the
paragraph-level test returns all four to `definedTerm`.

C04 is worth a second look by the reviewer rather than a claim here: the edit
capitalises a word so that it *becomes* a defined term, which is arguably row 4's
subject matter even though the changed text is a single word and so cannot match
a *multi-word* pattern. Recorded as an open question in the plan, not decided.

**C24 closes ADR-0029's open item.** The renumbering cascade now fires on all
four prose shapes above.

**Renderings are untouched.** This is a classification predicate; it does not
reach the diff. Corpus redundancy **0.11**, marks **83**, ideals **8/9** — all
unchanged, and the 4,000-pair generalisation gate is unaffected.

**ADR-0008 holds.** Nothing is dropped: `classifyClasses` still always returns at
least one class, every op is still surfaced and still accept/rejectable, and the
change only alters which label leads. If anything it *reduces* the risk P10 warns
about — a substantive edit labelled a defined-term change is a substantive edit a
reviewer might fold away on a first pass.

Full suite green at 1710 tests + 1 todo.

## Alternatives rejected

- **Keep the paragraph-wide signal in `classes` but bar it from being primary.**
  Preserves the old test verbatim and fixes the cascade. Rejected: §6 says
  `classes` are the categories *the change* matches, so a class the change does
  not match does not belong in the set either — and a "present but never
  primary" class is a trap for the next reader.
- **Reorder the priority table so row 5 outranks row 4.** Fixes the cascade and
  nothing else: every substantive edit would still be labelled a defined-term
  change. It also re-opens a settled ordering (substantive beats formatting)
  to work around a predicate bug.
- **Tighten `TITLE_CASE_MULTIWORD` to ignore a sentence-initial capital.** Helps,
  but only narrows the false-positive rate of a test that is asking the wrong
  question — `each Permitted Encumbrance under Clause 4.N applies` still
  mis-classified in the reproduction above.
- **Borrow feature 01's defined-term index (ADR-0013) to decide this.** The
  right long-term answer if Compare ever needs to know a document's actual
  defined terms, but it makes the classifier depend on another feature's
  analysis for a row the spec defines syntactically. Not needed for either limb.

---

## Note — 2026-09-04: one limb of `DEFINITION_PARA` is unreachable

`DEFINITION_PARA` tests for a definition paragraph by two limbs, one of which
looks for a term in **curly** quotes (`\u201C…\u201D`). That limb can never
match: `normaliseText` straightens curly quotes to `"` before any classifier sees
the text (`src/features/compare/preprocess/normalise.ts:23`), so every definition
paragraph reaches limb one. It is harmless — the straight-quote limb does the
work — but it reads as coverage that does not exist. Left in place rather than
removed, since a change to normalisation would make it live again; recorded here
so the next reader does not take it for a tested path.
