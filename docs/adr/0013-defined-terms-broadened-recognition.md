# ADR-0013: Defined Terms — broaden recognition beyond quoted-term + default verbs

- **Status:** Accepted
- **Date:** 2026-06-06

## Context

`spec/features/01-defined-terms.md` §5.1 recognises definitions through two shapes that are **both anchored on a quoted candidate term** (Shape A: quote-then-verb; Shape B: inline parenthetical), using a fixed default verb set in `CONVENTIONS.md#defined-term-conventions`. Quote-anchoring was chosen as a false-positive control for sub-check 1 (used-but-not-defined, Critical).

In review of the first implementation plan, the maintainer flagged that SG/UK drafting uses several legitimate definition formats this model silently drops:

1. **Unquoted** definition headwords — `**Affiliate** means …`, `AFFILIATE means …` (bold/caps term, no quotes). `ParaNode` does not expose run-level bold, so styling cannot be the anchor.
2. **Colon-as-verb** — `"Business Day": any day on which banks in Singapore are open …` (quoted term, no verb token).
3. **Verb phrasings outside the default set** — `shall have the meaning given to it in`, `has the meaning ascribed to it in`, `is defined in`, `shall mean and include`.

The maintainer accepted a higher false-positive rate ("OK as long as not too frequent") in exchange for catching these forms, and proposed the precision control: **ground unquoted definitions on a required definitional verb** — the verb does the anchoring the quote used to.

## Decision

Broaden the Defined Terms recogniser, in v1, beyond spec §5.1's quoted-anchored model:

- **Expanded verb set** (config, `definedTerms.verbs`): add `shall mean and include`, `shall have the meaning given to it in`, `shall have the meaning given in`, `shall have the meaning set out in`, `shall have the meaning specified in`, `has the meaning ascribed to it in`, `shall have the meaning ascribed to it in`, `is defined in`. The set remains config-extensible per matter.
- **Shape A′ — colon-as-verb:** a quoted term immediately followed by `:` is a definition (`"Term": <definition>`), `verb = null` (the colon is the marker, as for tables).
- **Shape C — unquoted, verb-anchored:** an unquoted capitalised candidate (a 1–6-word Title-Case run, or an ALL-CAPS acronym ≥2 chars) is a definition **only when immediately followed by a definitional verb** from the expanded set. The mandatory verb is the false-positive control. Additional guards (mandatory, to keep FPs bounded):
  - the candidate must be at a **definition position** — sentence-initial or paragraph-initial (definitions read `Affiliate means …`, not mid-sentence);
  - bare `means`/`includes`/`refers to` after an unquoted candidate fire **only** sentence-initial; the `…meaning…`-family and `shall mean` verbs may fire in either position;
  - the candidate must not be a self-reference / allowlisted structural word (`this Clause means …`, `the Company means business` must NOT flag — the determiner + non-initial position excludes them).
- **Config gating:** `definedTerms.recogniseUnquoted` (default `true`) and `definedTerms.recogniseColon` (default `true`) so the broadened shapes are tunable and can be disabled per matter if a document proves noisy.

Recognition order: quoted Shapes A/A′/B first; an unquoted Shape C candidate whose character span is already covered by a quoted match is skipped, so the same definition is never double-recorded.

Cross-reference vs cross-document classification, `contextScope`, and the sub-check pipeline are unchanged — the new shapes feed the **same** `DefinitionsMap`.

## Rationale

- The maintainer explicitly accepts bounded false positives to gain real-format coverage; the verb-anchor + definition-position guards bound them without needing run-level styling (which the Document Model does not surface).
- Verb-set expansion is already within the spec's design (the verb list is declared config-extensible); only Shapes A′ and C are genuine extensions of §5.1, and the verb requirement keeps Shape C far tighter than "any bold phrase."
- Config gating preserves the original conservative behaviour as a one-flag fallback, so the precision/recall trade-off stays in the user's hands.

## Rejected alternatives

- **Stay strictly quote-anchored (spec §5.1 as written).** Rejected: silently misses unquoted and colon-style definitions that are common in SG/UK practice, producing false "used-but-not-defined" Criticals on documents that *do* define their terms — the worst failure mode for this checker. The maintainer directed broadening.
- **Recognise any unquoted bold/caps phrase as a candidate (no verb requirement).** Rejected: false-positive explosion — every emphasised phrase becomes a "definition," and bold is not even available from `ParaNode`. The mandatory verb is what makes Shape C tractable.
- **Use run-level bold/italic as the unquoted anchor.** Rejected for v1: the Document Model does not expose per-run styling, and adding it touches the performance-critical single-pass loader for a marker the spec already calls non-load-bearing. The verb anchor is a better signal anyway (bold is inconsistently applied in practice).
- **Defer the broadened shapes to v2.** Rejected: the gap directly undermines sub-check 1's correctness on real SG/UK documents in v1; the maintainer wants it now.

## Consequences

- **Easier:** the checker recognises the unquoted, colon, and extended-verb definition forms that appear in real SG/UK contracts; fewer false "used-but-not-defined" Criticals on documents using those forms.
- **Harder / follow-on:**
  - Sub-check 1 false-positive surface grows; the definition-position and verb guards plus the `recogniseUnquoted`/`recogniseColon` config flags are the mitigations. Hardening must include the FP-trap edge cases (`The Company means business.`, `This Clause means that …` → **no** definition).
  - `spec/features/01-defined-terms.md` §2.1 and §5.1 carry a note pointing to this ADR; `DefinedTermConfig` gains `recogniseUnquoted` and `recogniseColon` and the expanded `verbs`.
  - This ADR modifies the recognition algorithm only; it does **not** touch the four invariants (still client-side, SG/UK-config-driven, in-scope feature 01, no Compare suppression). the scope register row 1 is unchanged.
