# ADR-0003: Client-side OOXML read-modify-write loop

- **Status:** Accepted
- **Date:** 2026-06-06

## Context

Several v1 operations cannot be expressed by the typed Office JS API: stripping `w:rsid*` attributes, deleting/merging section breaks (`SectionCollection` is read-only for structure), surgical `numbering.xml` edits, removing orphaned bookmarks, and authoring redline `<w:ins>`/`<w:del>` output. Word's native Compare is not exposed in Office JS either. ADR-0001 forbids a backend, so these cannot be done server-side with the Open XML SDK. A `.docx` is, structurally, just a ZIP of XML parts — which is editable in the browser.

## Decision

For operations the typed API cannot do, use a **client-side OOXML read-modify-write loop**, entirely in the task pane:

```
getFileAsync(Compressed) → reassemble slices → unzip (PizZip) → edit XML parts
  → re-zip → base64 → insertFileFromBase64(replace)
  (fallback: createDocument(base64).open() on detected fidelity artifacts)
```

This is encapsulated in the `shared/ooxml` service so feature code never hand-rolls slice plumbing or base64. **PizZip is the chosen ZIP library** for the synchronous loop (JSZip reserved for async/streaming of very large packages).

## Rationale

It keeps everything client-side (ADR-0001) while still delivering rsid stripping, section-break removal, and redline output. The `.docx`-is-a-ZIP fact makes pure-JS editing tractable. PizZip is synchronous and small, matching the loop's shape.

## Rejected alternatives

- **Server-side Open XML SDK.** Rejected: violates ADR-0001 — document bytes would leave the machine.
- **WASM / Pyodide in the task pane.** Rejected: ~6.4 MB+ load for what targeted JS string/XML edits already do; disproportionate weight in a 320–400 px sidebar. Reserved only for a future provably CPU-bound feature.
- **Giving up the features.** Rejected: rsid stripping, section-break cleanup, and redline authoring are core to Repair and Compare; dropping them guts those features.

## Consequences

- Must handle `insertFileFromBase64(replace)` fidelity edge cases (stray leading paragraph/blank page; package-level style/numbering/section loss). The service detects artifacts and falls back to `createDocument().open()`; freshly-authored packages (Compare) prefer a standalone new document.
- Some package-level parts (`fontTable.xml`, `theme1.xml`, certain `styles.xml` internals) are *not* safely editable client-side in v1 and must not be rewritten.
- The two-documents-in-memory `getFileAsync` limit forces serialised reads and prompt `closeAsync()`.
