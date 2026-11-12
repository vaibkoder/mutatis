# 09 — Party-Role Consistency (Proof tab)

**Status:** ✅ v1 — scope reopened by **ADR-0016** (supersedes the party-consistency clause of ADR-0005). scope row 11.

> Verifies that each party to the contract is referred to by **one consistent label** throughout. Surfaces two defects: (PR1) a party-role label used repeatedly in the body but **never introduced** as a party anywhere, and (PR2) **two different labels that the document's own text shows denote the same party** (the "Buyer"/"Purchaser" precedent-copy slip). Detect-and-navigate only; no auto-fix. Deliberately conservative on false positives — genuine dual-role deals (a real "Buyer" *and* a separate "Purchaser") must never be flagged.

## 1. Purpose & user value

When a precedent is copied and adapted, a party defined as **"the Buyer"** in the preamble is often later called **"the Purchaser"** in operative clauses — two names for one party, a silent ambiguity. The reverse slip is a role like **"the Agent"** used throughout the body that was never actually defined in the parties clause or definitions. Both are common, legally consequential, and invisible to the existing checkers: Defined Terms (feature 01) deliberately **allowlists** party-role words (ADR-0015) so it never flags them, and Cross-Reference (feature 02) is about clause numbers, not parties.

**User value in one line:** "Tell me when one party has two names, or when a party-role is used but never properly introduced — without nagging me about a deal that genuinely has both a Buyer and a Purchaser."

This checker is the narrowed revival approved in **ADR-0016**. The **execution / attestation-block** half of the original cut item (ADR-0005) **remains cut** and is **not** part of this feature.

## 2. Scope

### 2.1 In scope (v1) ✅

| # | Sub-check | Default severity | Description |
|---|-----------|------------------|-------------|
| PR1 | **Role used-but-never-introduced** | warning | A party-role label (from the config party-role vocabulary) used **≥ `minUsesForUndefined`** times in the body but with **no introduction anywhere** — no definition record, no inline parenthetical `(the "Agent")`, and not present in the parties preamble / Definitions clause region. |
| PR2 | **Inconsistent label (same party, two names)** | warning | Two distinct party-role labels for which the document supplies **explicit textual evidence** that they denote the **same** party. Evidence-gated (Clause 5.3); never fired on mere co-presence. |

Cross-cutting:
- **Reuse, no re-scan.** Consumes the in-memory Document Model snapshot, the Defined-Terms harvester (`DefinitionsMap`, `normaliseTerm`, usage alternation), and the shared **Section Index**. Issues **no** document-wide walk of its own.
- **Findings → Proof tab.** Every defect is a shared `Finding` with `featureId: 'party-role'`; this feature owns no results UI beyond the existing Findings renderer (it inherits the free-text filter and the occurrence stepper).
- **Footnotes** are scanned (the Document Model already exposes footnote-body paragraphs).
- **Config-driven vocabulary.** Party-role labels and entity suffixes come from the Config layer; nothing US-specific, nothing hardcoded (project invariants).

### 2.2 Out of scope / cut ❌🔜

- ❌ **Execution / attestation-block validation** (Companies Act 2006 s.44 two-signatory/witness variants; SG common-seal / constitution attestation). The other half of the ADR-0005 cut item — **stays cut** (ADR-0016).
- ❌ **Auto-fix / one-click relabelling.** Detect-and-navigate only (project invariants).
- ❌ **ML / embedding coreference / semantic entity resolution.** Evidence is explicit textual signal only (ADR-0016; mirrors the ADR-0008 no-suppression philosophy).
- ❌ **Cross-document party resolution** (Office JS cannot read other files — ADR-0005).
- 🔜 **Co-reference via shared definition-body similarity** (e.g. two roles whose definitions paraphrase each other). v1 uses only the two explicit evidence forms in Clause 5.3.
- 🔜 **Endnotes** (the Document Model scans footnotes only — same boundary as feature 01).

## 3. Dependencies (consume; do not re-specify)

- [Document Model / Paragraph Index](../ARCHITECTURE.md#document-model--paragraph-index) — paragraph text, `container` (`body`/`footnote`/`tableCell`), paragraph id, lazy range. Read-only; **no** independent scan.
- **Defined Terms harvester** (`src/features/defined-terms/`) — `buildDefinitionsMap` + `mergeTableDefinitions` build the `DefinitionsMap` (Shape B already captures inline `(the "Buyer")` introductions); `normaliseTerm` / `resolveLookup` for keying; `buildAlternation` for the usage scan. **Consumed, not duplicated** (project invariants).
- [Section Index](../ARCHITECTURE.md#section-index-builder) (`src/shared/section-index/`) — locates the **Definitions/Interpretation** clause and recital region, to decide whether a role was "introduced in the parties/definitions section" (Clause 5.2) and to anchor findings.
- [Findings model](../ARCHITECTURE.md#findings-model) — emit `Finding`s; `featureId: 'party-role'`; optional `occurrences[]` (added in ADR-0015) for the multi-use stepper.
- [Config layer](../ARCHITECTURE.md#localisation--config-layer) — party-role vocabulary (`definedTerms.exclusions.partyRoles`), entity suffixes (`entityFields`), and a new `partyRoleConsistency` config block (Clause 6).
- Conventions: [Defined-term conventions](../CONVENTIONS.md#defined-term-conventions) (party-role allowlist is the role vocabulary) and [Entity & Template Fields](../CONVENTIONS.md#entity--template-fields) (entity suffixes: UK `Limited/Ltd/plc/LLP`, SG `Pte. Ltd./Ltd./LLP` — **never** `Inc./LLC`).

## 4. Data model (built per run, then discarded — no persistence)

```
RoleRecord {
  key:          string        // normaliseTerm(label), e.g. "purchaser"
  label:        string        // canonical surface form, e.g. "Purchaser"
  introduced:   boolean       // defined / inline-parenthetical / in parties-preamble region
  introSites:   Anchor[]      // where introduced (for navigation)
  uses:         Anchor[]      // body+footnote occurrences, excluding the introduction occurrence
}

PartyAnchor {                 // one "<Entity> (the \"Role\")" hit (PR2 evidence form 1)
  entityKey:    string        // normalised preceding entity name, e.g. "acme pte ltd"
  entityText:   string        // surface entity name
  roleKey:      string        // normaliseTerm(role)
  roleLabel:    string
  anchor:       Anchor        // the parenthetical
}
```

`Anchor` is the existing `TextAnchor` (`{ paragraphId, startOffset, endOffset }`).

## 5. Algorithm

```
1. Load config (Clause 6) + active JurisdictionProfile.
2. map  = buildDefinitionsMap(bodyParas, footnoteParas, profile)
         + mergeTableDefinitions(map, tableParas, profile)        // REUSE — no re-scan
3. roleVocab = normalised set of profile.definedTerms.exclusions.partyRoles
              (minus the generic-collective words "Party"/"Parties" — Clause 8)
4. Build RoleRecords:
     - introduced  ⇐ key present in `map` (any record, incl. inline parenthetical / cross-ref)
                     OR label appears in the parties-preamble region (Clause 5.2)
     - uses        ⇐ usage alternation over body+footnote paras (REUSE buildAlternation),
                     excluding spans inside an introduction
5. PR1: for each role with introduced=false AND uses.length ≥ minUsesForUndefined → finding
6. Build PartyAnchors (entity-before-parenthetical scan, Clause 5.4)
7. PR2: detect same-party-two-labels via the two evidence forms (Clause 5.3) → findings
8. Emit Findings (Clause 7), clearing the 'party-role' prefix first.
```

### 5.1 Why reuse the harvester rather than re-scan

The `DefinitionsMap` already records every introduction shape that matters here: Shape A/A′ (`"Purchaser" means …`), Shape B (`(the "Purchaser")`), Shape C (`Purchaser means …`), table rows, footnotes, and cross-reference/cross-document classification. A role is "introduced" iff its normalised key is in the map (or it appears in the parties-preamble region). This is a pure read of an already-built structure — single-pass loading is preserved (one Document Model load feeds Defined Terms *and* this checker in the same Proof run).

### 5.2 The parties-preamble region

The parties preamble ("**BETWEEN: (1) ACME … and (2) BIDCO …**") often predates Clause 1 and may not carry a definitional verb, so a role can be introduced there without landing in the `DefinitionsMap`. Locate it via the Section Index: paragraphs **before the first clause/recital node** plus the recital region, and any clause node whose `displayLabel`/heading matches the configured **Definitions/Interpretation** noun. A role label appearing inside that region counts as `introduced=true` (suppresses PR1). This region is also where most `(the "Role")` introductions sit, so it mostly reinforces what the map already knows; it is the safety net that keeps PR1 conservative.

### 5.3 PR2 — same-party-two-labels, evidence forms (the only triggers)

PR2 fires **only** on explicit textual co-reference. Two forms, both auditable:

- **Form 1 — shared introducing entity.** Two distinct role labels are each introduced by an inline parenthetical against the **same** preceding entity name: `ACME Pte Ltd (the "Buyer")` … later … `ACME Pte Ltd (the "Purchaser")`. Built from `PartyAnchor`s (Clause 5.4) grouped by `entityKey`; a group with ≥2 distinct `roleKey`s is a hit.
- **Form 2 — definition-by-aliasing.** One role's definition body **is** the other role: a `DefinitionRecord` whose definition text (the span after the verb/colon) matches `^(the\s+)?<otherRole>$` after normalisation — e.g. `"Purchaser" means the Buyer`. Detected by reading the map entry's definition text against the role vocabulary.

**Never** fire because both "Buyer" and "Purchaser" merely appear (the dual-role deal). Absent Form 1 or Form 2 evidence, emit nothing (false-negatives-over-false-positives — ADR-0016).

### 5.4 Entity-before-parenthetical scan (feature-local, minimal)

The Defined-Terms Shape B regex discards the entity name preceding `(the "Role")`, so this checker adds one small local scan over the **same cached paragraph text** (still no document re-walk): for each `(\s*the\s+)?"(<role from vocab>)"\s*)` parenthetical, capture the immediately-preceding Title-Case / quoted entity run, optionally ending in a configured entity suffix (`Pte Ltd`, `Limited`, `plc`, `LLP`, …). Produce a `PartyAnchor`. Entity name normalisation: lowercase, collapse whitespace, strip trailing punctuation and a trailing entity suffix variant so `ACME Pte Ltd` and `ACME Pte. Ltd.` group together. If no entity precedes the parenthetical, no `PartyAnchor` is produced (Form 1 simply can't fire — correct).

## 6. Config

A new optional block on `JurisdictionProfile`, with safe defaults applied in code (so an older profile object still works). Role vocabulary and entity suffixes are **read from existing config**, not duplicated.

```jsonc
{
  "partyRoleConsistency": {
    "minUsesForUndefined": 3,        // PR1 recurrence gate — conservative (cf. ADR-0015 all-caps rule)
    "pr1Severity": "warning",
    "pr2Severity": "warning",
    "excludeCollectiveRoles": ["Party", "Parties"]   // too generic to flag
    // roleVocabulary  ← profile.definedTerms.exclusions.partyRoles  (consumed, not restated)
    // entitySuffixes  ← profile.entityFields                        (consumed, not restated)
  }
}
```

> The party-role vocabulary in `defined-term-lexicons.ts` currently lacks **Agent** (and a few financing roles). Adding `Agent`, `Trustee`, `Chargor`, `Mortgagor`, `Sponsor`, `Obligor`, `Indemnifier`, `Beneficiary` to the **shared** `PARTY_ROLES` lexicon is in-scope for this feature (it also improves feature 01's allowlist) and is an additive change — see the exec plan. No US roles are added.

## 7. UI / task-pane behaviour (Proof tab)

- One card in the Proof tab, run on demand as part of the existing **Run Proof** flow (`clearThenEmit('party-role', …)` alongside the other checkers in `FindingsList.handleRunProof`).
- Findings render through the existing Findings list (inherits the free-text filter and the ADR-0015 occurrence stepper / frequency sort). PR1 findings set `occurrences[]` for the `N×` stepper.
- Click-to-navigate via the standard lazy `getRange`/`search` fallback. PR2 findings expose **both** label sites as sub-links.

### 7.1 Findings mapping

| Sub-check | category | featureId | severity | suggested action (text only) |
|-----------|----------|-----------|----------|------------------------------|
| PR1 | `party-role/undefined-role` | `party-role` | warning | "'\<role\>' is used N times but never introduced as a party. Define it or confirm it is an ordinary word." |
| PR2 | `party-role/inconsistent-label` | `party-role` | warning | "'\<roleA\>' and '\<roleB\>' appear to be the same party (\<evidence\>). Use one consistent label." |

## 8. Edge cases

| Edge case | Handling |
|-----------|----------|
| **Genuine dual-role deal** ("Buyer" *and* "Purchaser" are different entities) | PR2 needs Form 1/Form 2 evidence; co-presence alone never fires. If each is introduced against a *different* entity, no finding. **The headline false-positive guard.** |
| **Collective roles** ("Party", "the Parties") | Excluded via `excludeCollectiveRoles` — generic, structurally always "present". |
| **Role introduced only in the preamble (no verb)** | Clause 5.2 parties-preamble region marks it `introduced=true` → no PR1. |
| **Role defined by cross-reference / in a table / footnote** | Counts as introduced (it is in the `DefinitionsMap`). |
| **Plurals / possessives** ("Purchasers", "the Buyer's") | `normaliseTerm` + `resolveLookup` (REUSE feature 01 rules). |
| **Same entity, same role twice** (`ACME (the "Buyer")` … `ACME (the "Buyer")`) | One `roleKey` in the entity group → not a hit (PR2 needs ≥2 distinct roles). |
| **Sentence-initial / heading capitalisation** | Usage scan reuses feature-01 anchoring; the role vocabulary is a closed set, so sentence-initial "Buyer" is still a real role use — counted normally. |
| **Bare `$` / entity suffix variants** (`Pte Ltd` vs `Pte. Ltd.`) | Normalised in the entity-name fold (Clause 5.4) so they group. |
| **Tracked changes** | v1 reads the displayed/final text (same boundary as feature 01). |
| **No party roles at all** | Empty result; no findings; not an error. |

## 9. Test cases / acceptance

`**…**` = bold; fragments are minimal.

- **T1 — clean.** `ACME Pte Ltd (the "Buyer") … the Buyer shall pay.` → **no** findings.
- **T2 — PR2 Form 1 (shared entity).** `ACME Pte Ltd (the "Buyer") … ACME Pte Ltd (the "Purchaser") shall complete.` → **1 warning** inconsistent-label (Buyer ≡ Purchaser; evidence: shared entity "ACME Pte Ltd"); both sites linked.
- **T3 — PR2 Form 2 (aliasing).** `"Purchaser" means the Buyer.` (Buyer introduced elsewhere) → **1 warning** inconsistent-label.
- **T4 — dual-role, NO flag.** `ACME Pte Ltd (the "Buyer") … BIDCO Ltd (the "Purchaser") …` (different entities) → **no** PR2 finding.
- **T5 — PR1 undefined role.** `the Agent shall notify … the Agent may … the Agent's fees …` (Agent in vocab, used 3×, never introduced) → **1 warning** undefined-role, `occurrences = 3`.
- **T6 — PR1 below threshold.** `the Agent shall notify.` (used once) → **no** finding (`minUsesForUndefined` gate).
- **T7 — PR1 introduced in preamble.** `BETWEEN … (2) BANKCO (the Agent) …` then `the Agent shall …` ×3 → **no** PR1 (introduced in parties region).
- **T8 — collective role.** `the Parties agree …` ×5, never "defined" → **no** finding (excluded collective).
- **T9 — plural/possessive.** Defined `(the "Shareholder")`; body "the Shareholders" / "the Shareholder's rights" → resolves to one role; introduced ⇒ no PR1; no spurious PR2.
- **T10 — co-presence without evidence.** Both "Buyer" and "Seller" used, each introduced against its own entity → **no** PR2 (different parties, expected).

**Acceptance (rollup):** T1–T10 produce exactly the expected sets; PR2 fires only on Form 1/Form 2 evidence; dual-role deals never flag; PR1 honours the recurrence gate and the preamble/definition introduction suppressors; all analysis reuses the harvester + Section Index with no second document walk; findings carry `featureId: 'party-role'` and render in the Proof tab.

## 10. Known limitations

- **Evidence-gated PR2 misses unmarked coreference.** If a precedent slip introduces "Buyer" and "Purchaser" with no shared-entity parenthetical and no aliasing definition, v1 will not catch it. Deliberate (ADR-0016): a false negative is preferred to flagging legitimate dual-role deals.
- **PR1 recurrence floor.** A genuinely-undefined role used once or twice is suppressed (conservative, cf. ADR-0015). It surfaces once it recurs.
- **Vocabulary-bounded.** Only labels in the config party-role vocabulary are considered; a bespoke role word not in the lexicon is invisible. The lexicon is config-extensible per matter.
- **Entity-name heuristic.** The entity-before-parenthetical scan is a surface heuristic; an introduction without a recognisable entity run yields no Form-1 anchor (Form 2 and PR1 still apply).
- **Tracked changes / endnotes** — same boundaries as feature 01.

## 11. Build milestones

1. **M1 — Types + role/usage model.** `RoleRecord`, `PartyAnchor`; build RoleRecords from a passed-in `DefinitionsMap` + usage alternation. Unit-test on plain text (no Office JS). (T5, T6, T9.)
2. **M2 — Parties-preamble region via Section Index.** Introduction suppressor. (T7.)
3. **M3 — PR1.** Recurrence gate + introduced/collective suppressors → findings. (T5–T8.)
4. **M4 — Entity-before-parenthetical scan + PR2 Form 1 & Form 2.** (T2, T3, T4, T10.)
5. **M5 — Config block + lexicon additions** (`partyRoleConsistency`, add `Agent` et al. to shared `PARTY_ROLES`). en-GB; SG/UK profiles. (T1.)
6. **M6 — Findings mapping + Proof-tab wiring** (`clearThenEmit('party-role')`, category labels, sub-link sites). Full T1–T10 pass; reuse verified (no second document walk).
