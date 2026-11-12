# ADR-0018: Text-vs-text paste comparison with instant standalone redline

- **Status:** Accepted  <!-- user-approved 2026-06-08, in the same change as the scope register update below -->
- **Date:** 2026-06-08
- **Builds on / extends:** feature 08 (Compare) and **ADR-0008** (categorise, never suppress). Adds a new ✅ item to the scope register. Sibling of **ADR-0017** (PDF comparison sources) but a distinct decision — see Context.

## Context

A third comparison workflow sits outside the shipped Compare feature (08, Word-vs-Word) and the PDF modes (ADR-0017):

- **Paste two plain-text blobs** — an *Original* and a *Modified* — into the task pane and get an **instant tracked-changes redline**, **auto-generated as a standalone `.docx`** that **opens only when the user clicks** (never auto-launched).

This is a routine need: comparing two snippets (a clause two ways, a counterparty's pasted wording vs the house position, two e-mailed paragraphs, a term-sheet line) **without first saving each as a `.docx`**. Today that means manual reconciliation or pasting into a cloud diff tool — the same confidentiality problem the product exists to remove.

The shipped Compare feature (08) already authors a true Word redline (`<w:ins>`/`<w:del>` with `w:author`/`w:date`) and materialises it as a **standalone new document** via the OOXML Service (`Word.createDocument(base64).open()` — `output/writeout.ts → materialise({mode:'newDocument'})`). That diff → classify → author → materialise chain operates on `ParaRecord[]` plus a `.docx` bundle and carries **no dependency on where the text came from**. Therefore, by **synthesising a minimal valid `.docx` from each pasted blob**, the text path can ride the *entire existing pipeline* — the **highest reuse of any new comparison source**, and the only one that yields an editable Word redline.

This is a **different decision from ADR-0017** and is recorded separately because:
- It uses **no PDF.js**, so it has **none** of ADR-0017's offline-asset / worker / cmap / reconstruction-heuristic risk surface.
- It **does author a redline** (`output/**` reused) — ADR-0017 explicitly authors *no* redline.
- It introduces one genuinely new capability: **synthesising a standalone, valid `.docx` package from plain text** (a controlled template carrying the styles tracked-change rendering needs).

## Decision

Add a third Compare **input mode — Text-vs-text** — built on the existing pipeline:

1. **Input:** two text areas, *Original (base)* and *Modified (compared)*; plain-text paste.
2. **Adapt:** split each blob into paragraphs (**one paragraph per line**, deterministic), then **synthesise a minimal valid `.docx` per side** — a controlled package (`[Content_Types].xml`, `_rels/.rels`, `word/document.xml` with one `<w:p>` per line, `word/styles.xml` carrying the styles a redline needs, `word/settings.xml`, rels) with `styleId=null` semantics and text escaped into a single run per paragraph.
3. **Reuse (unchanged):** feed both synthesised bundles into the **existing** `loadBothSides → runPipeline → output/author → materialise` chain. Structural/word diff, the full SG/UK classifier, cascade collapse, and move detection all apply verbatim.
4. **Output:** a **real tracked-changes `.docx`** (accept/reject in native Word) — *strictly more* than the PDF modes offer.
5. **Eager generate, click to open:** the redline is **built eagerly** on Compare (instant), held in memory as bytes, and **opened only on the user's explicit click** (a button → `materialise({mode:'newDocument'})` / a `.docx` download). It is **never auto-opened** — no surprise Word windows.

**v1 boundary:** **plain text only** (deterministic line-split). No rich-text/HTML/Markdown paste parsing, no style inference, no formatting round-trip.

## Rationale

- **Maximum reuse / minimum new surface.** The diff, the entire SG/UK classifier, the redline author, and the standalone-document materialiser are reused as-is. The only new code is the text→docx synthesiser, the two-text-area UI, and the eager-generate / click-to-open wiring.
- **Zero friction.** Paste-and-compare beats save-each-as-`.docx`-then-upload for the snippet use case.
- **Genuine editable redline.** The user gets native accept/reject — the thing lawyers actually want — not just a read-only view.
- **Confidentiality preserved.** The synthesised package and the redline live entirely in memory (Blob / base64); `Word.createDocument().open()` is a **local Office API**; nothing is transmitted.

## Rejected alternatives

- **Auto-open the redline on Compare.** Rejected — the user asked for **click-to-open**; eager-opening spawns Word windows unexpectedly and steals focus. Generate eagerly, open lazily.
- **A bespoke in-pane text differ with no `.docx` output.** Rejected — it would rebuild the classifier and forgo the editable-Word-redline value, which is the whole point.
- **Require the user to save each blob as a `.docx` and upload it.** Rejected — defeats the instant-paste value; that path already exists (feature 08 docx-vs-docx upload).
- **Rich-text / HTML / Markdown paste parsing with style inference.** Rejected for v1 — non-deterministic, a styling rabbit hole, and scope creep. Plain text only; the line-split is deterministic and testable.
- **A pixel/visual output surface.** N/A — text in, Word redline out.
- **ML semantic-equivalence / any suppression.** Remains ❌ (ADR-0008). The redline carries every change; the classifier only categorises/filters.
- **Folding this into ADR-0017.** Rejected — opposite output policy (authors a redline vs authors none) and a disjoint risk surface (no PDF.js); bundling would muddy both decisions.

## Consequences

- A new, **small** feature lands on the existing Compare engine. New code is confined to: `compare/io/text/synthesise-docx.ts` (the package synthesiser), a `text` source kind in `compare/io/load.ts`, the two-text-area UI in `ComparePanel.tsx`, and driver wiring for eager-generate / click-to-open. `output/author.ts` and `output/writeout.ts` are reused **unchanged**.
- **New capability to get right:** synthesising a **valid** minimal `.docx` that Word opens cleanly *and* renders tracked changes (correct `[Content_Types].xml`, rels, and a `styles.xml`/`settings.xml` supporting revisions). This is the one real engineering risk; it is bounded and testable (OOXML validation + open-in-Word live-verify). Covered by the internal implementation plan.
- Confidentiality, jurisdiction, scope, and categorise-never-suppress are all preserved; this ADR + the user's approval + the scope register update **is** the scope gate.
- Lands with the combined feature spec `spec/features/10-compare-additional-sources.md` and a the scope register update, in the same change.
