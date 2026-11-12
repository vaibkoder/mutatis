# ADR-0004: Technology stack and WordApi requirement-set floor

- **Status:** Accepted
- **Date:** 2026-06-06

## Context

Mutatis is a Word task-pane SPA running in an evergreen-browser WebView, doing heavy nested `ctx.load`/`ctx.sync` orchestration and OOXML string surgery, in a narrow (~320–400 px) sidebar over large (200–400 page) legal documents. It needs a small bundle with fast first paint, strong static typing, and access to specific Office JS capabilities (footnotes, native TOC fields, granular change events, tracked-changes accept/reject).

## Decision

Record the stack from `spec/ARCHITECTURE.md`:

- **Language:** TypeScript, `strict` mode (mandatory).
- **UI framework (recommended):** **Preact (~3 kB) + `@preact/signals`**; acceptable alternative plain TS + lit-html; avoid Angular and full Fluent UI React.
- **Build:** **Vite** (scaffolded once from `yo office`, then migrated).
- **Key libraries:** **PizZip** (sync zip) + **fast-xml-parser**; **diff-match-patch** (+ fast-diff) for Compare; **chrono-node** (day-first); **words-to-numbers**; **Fuse.js** (library search); **PDF.js** (v2 only).
- **WordApi requirement-set floor: 1.5** (footnote bodies and native TOC field insertion are core; 1.3/1.4 subsumed). Features needing **1.6** (`getTrackedChanges().acceptAll()/rejectAll()`) are feature-detected and degrade gracefully.

**Not yet ratified:** the **front-end framework and the overall UX/visual design are RECOMMENDED, not ratified.** Preact+signals is the working recommendation; a follow-up ADR (ADR-0012) confirms the front-end/UX decision.

## Rationale

A tiny reactive framework gives fast first paint and fine-grained Findings-list reactivity in a narrow pane; TypeScript's typing is load-bearing for the `ctx.sync` and OOXML work. A 1.5 floor unlocks the core feature set across all platforms while keeping 1.6-only capabilities optional.

## Rejected alternatives

- **Heavier SPA frameworks (Angular, full Fluent UI React).** Rejected: bundle weight and slow first paint in a sidebar; disproportionate to the UI's complexity.
- **Lower requirement-set floor (1.3/1.4).** Rejected: would lose footnote-body scanning, native TOC field insertion, and (at 1.6) tracked-changes accept/reject APIs that core features depend on.

## Consequences

- Features gate on requirement sets via `Office.context.requirements.isSetSupported`; the Capabilities module (`shared/office`) is the single source of truth.
- Word Online and any host lacking a desktop-only/1.6 API degrades gracefully (hide/disable with a tooltip), never throws.
- The front-end/UX decision is explicitly open until ratified by ADR-0012; do not treat Preact as final.
