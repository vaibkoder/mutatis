# ADR-0012: Front-end framework, UX design, and product name

- **Status:** Accepted — **partially superseded by ADR-0044** for the standalone web target (see the two marked clauses below). Every task-pane decision here stands in full.
- **Date:** 2026-06-06
- **Supersedes (in part):** ADR-0004 — ratifies the "not yet ratified" front-end/UX section

## Context

ADR-0004 recorded the technology stack but explicitly left the front-end framework and overall UX/visual design **unratified**, deferring to a user interview in the first coding session. ADR-0004's consequences note: *"The front-end/UX decision is explicitly open until ratified by a later ADR; do not treat Preact as final."*

The first coding session (2026-06-06) conducted that interview using a set of Phase 1 questions, covering: framework, visual language, tab/IA structure, Findings UX, theming, accessibility, and product name.

## Decision

### Product name
**Mutatis** — an earlier working name is retired and replaced by this one. The new name is used throughout (`package.json` `name`, manifest `DisplayName`, UI strings, documentation).

### Front-end framework
**Preact (~3 kB) + `@preact/signals`** — confirmed as the framework. Signals provide fine-grained Findings-list reactivity without full re-renders; the bundle stays small for a narrow (~320–400 px) task pane. This matches ADR-0004's primary recommendation.

### Visual language
**Fluent UI Web Components (`@fluentui/web-components`) + `Office.context.officeTheme` tokens.** The pane blends into Word's chrome using Microsoft's own component tokens and the live Office theme object. This uses Web Components (framework-agnostic), not `@fluentui/react-components` (React-only, heavy).

### Task-pane tab structure
**Keep as specified in `spec/ARCHITECTURE.md`:**
Proof · Review · Repair · Numbering · Content · Compare · Settings
(badge counts from the Findings store; each tab maps to a feature group as spec'd).

### Findings / results UX
**Severity-grouped list, click-to-navigate:** findings grouped Critical → Warning → Info; each row shows title + location label; clicking the row calls `range.select()` on the target location; one-click `suggestedAction.apply` where present; badge counts on tab headers. Matches the spec's Findings model exactly.

### Theming
**Follow `Office.context.officeTheme` automatically** — light, dark, and high-contrast inherited from Word's own setting; no user toggle in the task pane. **(Scoped to the task pane by ADR-0044: the standalone web target — which did not exist when this was written, see ADR-0022 — has no host theme to follow and carries an explicit light/dark/system control.)** CSS custom properties bridge the Office theme tokens into the component layer via `styles/office-theme.css`.

### Accessibility
**WCAG 2.1 AA** — industry standard; aligns with Microsoft's Accessibility Standard for Office Add-ins. Covers: colour contrast ≥ 4.5:1 (text), keyboard navigation, focus management (focus trap on modals, visible focus ring), screen-reader semantics (ARIA roles/labels on dynamic regions), and logical reading order.

## Rationale

Every question was confirmed as the spec's recommendation, so this ADR ratifies the spec rather than deviating from it. The key architectural reasons:

- **Preact over React**: avoids the ~40 kB React overhead for a sidebar whose UI is 90% a list + a tab bar. Signals avoid full-tree re-renders on every Findings-store update.
- **Web Components over Fluent UI React**: framework-agnostic; integrates without wrapping Preact children in React portals; Office's own newer UI layer (Copilot, Loop) already uses Web Components.
- **Office theme bridge**: required for Word Online and high-contrast accessibility modes; one bridge file removes all per-component colour hardcoding.
- **WCAG 2.1 AA over 2.2**: 2.2 adds useful focus-appearance criteria but the delta is small; adopting 2.1 AA now and upgrading to 2.2 incrementally is lower-risk for v1.

## Rejected alternatives

- **React + `@fluentui/react-components` v9**: ~40–50 kB gzipped for React alone; full Fluent UI React adds more. Disproportionate for the sidebar's UI complexity. ADR-0004 already advised against this.
- **Plain TypeScript + lit-html**: acceptable per spec, but Preact's component model is more ergonomic for the shell (tabs, modals, settings panel); signals give reactivity without a custom event bus.
- **Custom design system / own brand colours**: adds a design-maintenance burden with no functional benefit; the Word-native look builds user trust for a tool used inside Word.
- **Task-first "what do you want to do?" flow**: hides feature discoverability; experienced lawyers want direct tab access to specific checkers, not a wizard that adds a routing layer.
- **User-selectable independent theme**: adds state management (theme persistence) and requires designing a second colour palette, for marginal UX gain when Word already surfaces the user's preference. **(Still rejected for the task pane. Revisited and ACCEPTED for the standalone web target by ADR-0044 — the rationale above turns on the HOST surfacing the preference, and a browser tab does not.)**
- **WCAG 2.1 A (minimal)**: omits several colour-contrast and keyboard-nav requirements; unacceptable for a tool targeting professional, often accessibility-sensitive law firm environments.

## Consequences

- `@fluentui/web-components` is added as a dependency; `@fluentui/react-components` is **not** used.
- `styles/office-theme.css` bridges `Office.context.officeTheme` → CSS custom properties consumed by all components.
- All UI strings are en-GB; no US variants in the component layer.
- The seven tab identifiers (`proof`, `review`, `repair`, `numbering`, `content`, `compare`, `settings`) are stable string keys — features register themselves by tab key, and the shell renders them.
- WCAG 2.1 AA is an acceptance criterion for the Task Pane Shell and every component it renders; a spec-conformance review checks for obvious violations (missing ARIA labels, low-contrast hardcoded colours).
- ADR-0004's "not yet ratified" note is now superseded; Preact + signals and the rest of this ADR are the ratified decisions.
