# Cross-Reference Checker & Static Numbering Detector (Feature 02)

Two **distinct** detect-and-flag features under the Proof tab, kept separate by
the `featureId` discriminator (spec §5.3, §8.4). They share the word "static"
by coincidence of English, not by domain — **never merge their findings**.

- **Feature A — Cross-Reference Checker** (`featureId: 'xref'`). Validates prose
  references ("see Clause 4.2", "paragraph 3 of Schedule 2") against the Section
  Index → OK / Not Found / Ambiguous / externalMaybe. Badges manually-typed
  (non-REF-field) pointers. Offers an Assign-source repair (static text only).
- **Feature B — Static Numbering Detector** (`featureId: 'staticNumbering'`).
  Flags clause *numbers* that are hand-typed rather than list-engine-generated.
  **Detect-only — never converts** (ADR-0006).

## Module map

| File | Role |
| --- | --- |
| `driver.ts` | Orchestrates both tracks; `clearThenEmit` clears both prefixes. |
| `scanner.ts` | Feature A: compiles the reference regex library from config; scans body prose → `ReferenceMatch[]`. Pure. |
| `resolve.ts` | Feature A: resolves keys against the index (incl. prefix-implied-parent ambiguity, compound both-halves validation, externalMaybe). Pure. |
| `fields.ts` | Feature A: feature-local REF/PAGEREF field + bookmark read (ADR-0014). Async; graceful fallback. |
| `static-numbering.ts` | Feature B: layers (a) leading-token regex + (b) list-membership. Pure. |
| `static-numbering-ooxml.ts` | Feature B: layer (c) capped `getOoxml()` probe for residuals. Async. |
| `assign-source.ts` | Feature A repair: `range.insertText(label, 'Replace')` (static text only; v2 `refField` seam). |
| `locate.ts` | Finding builders with hard `featureId` split + category constants. |
| `types.ts` | `ReferenceMatch`, `StaticNumberingResult`, `UnitType`, etc. |

The Section Index itself is shared infra (`src/shared/section-index/**`), authored
on behalf of this feature and consumed later by Interactive Lookup and Numbering.

## Architecture notes

- **ADR-0014** — REF fields, bookmarks, and `getOoxml()` are read **feature-locally**
  (bounded, batched, capped), not by enriching the shared single-pass loader.
  All reads are client-side; no bytes leave the machine.
- **ADR-0014 fallback (mandatory)** — if `paragraph.fields` is unavailable on the
  host, every reference degrades to un-backed (static badge applies, broken-field
  detection disabled) behind a non-blocking capability banner. Never throws.
- **ADR-0006** — Feature B performs **zero** mutation. No `numbering.xml` edits,
  no `insertFileFromBase64`, no static→dynamic conversion. The UI offers no fix
  action on Feature B rows.
- **Symmetric normalisation** — the cited-number normaliser (scanner) and the
  index key-builder use the **same** routine (`shared/section-index/normalise.ts`)
  so "Clause 4.2.1" / "clause 4.2.1" / a REF result all collapse to `clause:4.2.1`.
- **Jurisdiction** — all unit nouns come from `profile.xref.*`. Default **Clause /
  Schedule**; "Section" / "Exhibit" are recognised via aliases but **never emitted**
  as default replacement text (spec §7 hard rule; AC §12.3).

## Known limitations (spec §10)

1. **No true staleness detection** — the static-reference badge is a proxy
   ("Word won't auto-maintain this"), not proof that a reference moved.
2. **Static text replacement only** in Assign-source (no live REF field in v1;
   `replaceStrategy: 'refField'` seam present, not implemented).
3. **No cross-document resolution** — sibling-file schedules/exhibits are flagged
   "may be in a separate file" and never resolved (Office JS cannot read other files).
4. **Blind to text boxes, headers, footers, and footnotes/endnotes** — both
   features scan **`body.paragraphs` only**. References and numbering in those
   containers are **not checked**. (`container !== 'body'` paragraphs are skipped
   in both `scanner.ts` and `static-numbering.ts`.) This is the documented v1
   blind spot; header/footer/text-box coverage is a post-v1 seam.
5. **Feature B is detect-only** — no conversion; the user fixes numbering manually.
6. **Roman/letter ambiguity** (`(i)`/`(v)`/`(x)`) is resolved heuristically by
   sibling sequence; pathological mixed lists may mis-classify, so the matcher
   accepts either reading (`ambiguousRomanLetter`) to avoid false Not Found.
7. **OOXML probing is capped** (`staticNumbering.maxOoxmlProbes`, default 200);
   beyond the cap, residual ambiguous paragraphs are reported `indeterminate`.
