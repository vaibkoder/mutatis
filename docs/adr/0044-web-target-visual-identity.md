# ADR-0044: The standalone web target has its own visual identity and an explicit theme preference

- **Status:** Accepted
- **Date:** 2026-09-19
- **Relates to:** ADR-0012 (front-end/UX — **partially superseded**: its "no user toggle" decision and its rejected "user-selectable independent theme" alternative are now scoped to the task pane); ADR-0022 (standalone implementation — the target this applies to); ADR-0037 (batch compare — web-only surfaces re-skinned here); SCOPE row 14
- **Supersedes:** ADR-0012 §Decision "Follow `Office.context.officeTheme` automatically … no user toggle in the task pane" **for the web target only**, and its Rejected alternative "User-selectable independent theme".

## Context

`src/styles/web.css` began as a *bridge*: its own header called it "a
system-theme dark bridge (mirrors office-theme.css dark token set)", and it did
exactly that — restated the add-in's Fluent palette against
`prefers-color-scheme` because there is no `Office.context.officeTheme` in a
browser tab. The web app therefore looked like a Word task pane that had been
stretched to 1280px: 14px body copy, 2px radii, `1.25rem` headings, a Fluent
blue accent, and no way for a user to choose a theme.

Two forces made that untenable.

**One: Mutatis Desktop exists and looks nothing like it.** It is a dense,
quiet, border-separated UI — oklch tokens named for roles rather than hues,
three surface tiers and three text tiers, 12px body / 11px labels / 10px
metadata, 4px and 6px radii, sunken rails against a plain work area, and
essentially no shadows or gradients. A user moving between the desktop app and
the web app was moving between two products.

**Two: ADR-0012's no-toggle decision does not transfer.** Its reasoning is
stated in the ADR itself — *"for marginal UX gain when Word already surfaces
the user's preference"*. That premise is about the **host**. A browser tab
surfaces no such preference beyond the OS setting, and a lawyer who wants the
tool dark on a light laptop has, today, no way to say so. The decision was also
taken in June 2026, before ADR-0022 created a standalone target at all, so the
web app is not a case it weighed and rejected — it is a case it never saw.

Against that sits a real hazard. `shell.css` is ~3,500 lines of consumers, it is
**shared** with the task pane, and ADR-0012's token discipline ("no component
hardcodes a colour") is the only reason a palette change is tractable at all.
Anything that re-values tokens must not reach the add-in.

## Decision

**1. The web target carries the desktop app's design language, expressed as new
VALUES for the existing `--dp-*` tokens inside `src/styles/web.css`.**

No token is renamed or removed. `web.css` is loaded only by `src/web/index.html`
and `src/web/main.tsx`; the task pane loads `tokens.css` + `office-theme.css` +
`shell.css` and never `web.css`. That seam is the whole safety argument and it
is asserted, not assumed (`tests/unit/web/web-theme-tokens.test.ts`), including
that no `oklch(` appears in any shared stylesheet.

Tokens added alongside the existing ones, because the desktop has tiers
`tokens.css` does not: `--dp-bg-raised`, `--dp-border-strong`, `--dp-fg-faint`,
`--dp-accent-soft`, `--dp-focus`, `--dp-radius-md`.

**2. The web target gains a light / dark / follow-system control.** It writes
`data-theme` on `<html>`, sets `style.colorScheme`, and persists to
`localStorage` inside `try/catch` on both read and write. **The task pane does
not and must not** — it follows the live Office theme, which is the user's own
Word setting, and ADR-0012 settled that.

**"Follow system" REMOVES the attribute rather than writing the resolved
value.** `web.css`'s pre-paint block is scoped to `:root:not([data-theme])`,
and that block is what makes the *first* paint correct for everyone who has not
overridden anything — the module script is deferred, so it cannot run before
that paint. Writing `light`/`dark` for a system user would type-check, look
identical in every screenshot, and cost them a frame of the wrong theme on
every load.

**3. WCAG 2.1 AA (ADR-0012 §Decision) is a constraint on the palette, not a
casualty of it — and honouring it forced five values away from the desktop's.**

Measured by converting every token to sRGB and computing WCAG 2.1 contrast for
every text/surface pair this stylesheet can create. A direct port of the
desktop's values fails **12 pairs** in light mode:

| pair (light) | desktop value | ratio |
|---|---|---|
| `fg-faint` on bg / raised / subtle / hover | `L 0.66` | 3.02 / 3.11 / 2.81 / 2.53 |
| `warning` on bg / raised / subtle / hover | `L 0.58` | 4.25 / 4.37 / 3.95 / 3.56 |
| `check` on bg / raised / subtle / hover | `L 0.60` | 3.87 / 3.99 / 3.60 / 3.24 |
| `fg-subtle` on `bg-hover` | `L 0.52` | 4.48 |
| `error` on `bg-hover` | `L 0.55` | 4.35 |

The structural finding: **against a near-white surface, AA caps text at about
L = 0.56**, so all three text tiers have to fit inside `[0.25, 0.56]`. The
desktop's faint tier cannot pass at any chroma. Shipped values, all measured:
`fg-subtle 0.48`, `fg-faint 0.53 / 0.645`, `error 0.51`, `warning 0.50`,
`check 0.50`. Worst surviving pair **4.75:1**.

**`--dp-fg-faint` is AA-safe only on `bg`, `bg-raised` and `bg-subtle`, and not
on `bg-hover` (4.09:1).** Compressing it further to survive a hover fill would
put it *below* `fg-subtle` and collapse the third tier, so the tier is kept and
the limit is stated on the token itself. Anything inside a rule with a `:hover`
fill uses `--dp-fg-subtle`. This is an accepted limit, not an oversight.

**4. Redline marks are untouched.** They live in `src/shared/render-core/types.ts`
as document ink (`insertColor`, `deleteColor`, `moveColor`, change bar), not as
theme tokens, and the desktop's `--redline-*` values were derived *from* this
add-in's convention so that a reviewer reads the same colour as the same thing
in both products. There is nothing to reconcile.

**5. Nine tokens that were silently falling back are now defined web-side.**
`shell.css` and `web.css` reference `--dp-accent-rgb`, `--dp-accent-bg`,
`--dp-accent-hover`, `--dp-warn`, `--dp-ok`, `--dp-fg-error`,
`--dp-bg-elevated`, `--dp-space-md` and `--dp-font-size` — none of which
`tokens.css` declares. Six carried a hardcoded Fluent-era literal in the
`var()`; **two (`--dp-bg-elevated`, `--dp-space-md`) carried no fallback at
all**, which makes the whole declaration invalid at computed-value time and
paints a transparent background and a zero padding. Defining them in `web.css`
fixes this for the web without touching the shared file. Measured against the
built bundle: four of the nine (`--dp-accent-rgb`, `--dp-warn`, `--dp-ok`,
`--dp-accent-hover`) have web-reachable consumers today; the rest have only
Numbering / AI-export consumers the web app never mounts, so defining them has
**no visible effect today** and is insurance against the first web surface that
uses them.

## Rationale

- **Re-valuing beats re-theming.** ADR-0012's "no component hardcodes a colour"
  is what makes a whole-palette change a single-file edit. Honouring that rule
  is cheaper than any alternative and keeps the add-in provably untouched.
- **The no-toggle rejection was host-specific and is preserved where it
  applies.** The task pane still has no toggle. Only the target with no host
  theme to follow gains one.
- **AA is already the project's own acceptance criterion** (ADR-0012 §Decision,
  applied to "the Task Pane Shell and every component it renders"). A palette
  that fails it on twelve pairs is not a palette this project can ship, and the
  five deviations are small enough that the two products still read as one.
- **Amber is the hardest hue to bring to 4.5:1 on near-white.** `tokens.css`'s
  originals (`#8a6d00`, `#8a6300`) were already dark for that reason, so a
  naive port of the desktop's lightness would have been a **regression** on a
  value someone had previously got right.

## Rejected alternatives

- **Port the desktop values exactly and accept the twelve sub-AA pairs.**
  Pixel-identity with the desktop app, but it breaks ADR-0012's stated
  accessibility bar in a tool aimed at professional, often
  accessibility-sensitive, law-firm environments. Raised with the maintainer with
  the measurements; rejected by them.
- **Keep `fg-subtle` at the desktop's `L 0.52` (nudged to `0.51` to pass).**
  Passes AA and is closer to the reference, but leaves only `0.03` of lightness
  between the bottom two text tiers — which is not three tiers, it is two tiers
  and a rounding error. Rejected in favour of `0.48`.
- **Compress `fg-faint` until it is AA-safe on `bg-hover` too.** Requires
  `L ≤ 0.519`, which is *darker than `fg-subtle`*. The tier inverts. Rejected;
  the usage limit is stated on the token instead.
- **Re-tune the redline marks into the oklch family.** They are document ink,
  they must print identically in light and dark, and their colours are the
  shared convention between the two products. Rejected outright.
- **Edit `shell.css` to fix the nine falling-back tokens at source.** It is
  shared with the add-in, and changing a fallback there changes the task pane.
  Rejected — defining the tokens web-side achieves the same result with no
  blast radius.
- **A theme `<select>`, or a two-state light/dark switch.** A two-state switch
  cannot express "follow my OS", which is the correct default and the one most
  users want; a `<select>` is a heavier control for three mutually exclusive
  options. Rejected in favour of a three-button radiogroup, which is also what
  the desktop app uses.
- **Add an icon dependency (lucide or similar) for the five glyphs.** The
  standalone target ships as a static backendless bundle; a package for five
  icons is weight carried forever, and lucide's ISC licence would need its
  notice carried with it. Rejected — the glyphs are hand-authored in
  `src/web/icons.tsx` following the same drawing conventions (24-unit grid, 2px
  strokes, round caps, `currentColor`).
- **Tailwind, to match the desktop app's stack.** The brief was to copy the
  design, not the stack. Rejected; it would be a build-system change for a
  stylesheet change.

## Consequences

**Easier.** A user can pick a theme. The two products read as one family. Nine
tokens stop resolving to colours nobody chose, and two stop being invalid CSS.
The palette is measured, so the next change to it has a method rather than an
opinion.

**Harder / owed.**

- **`--dp-fg-faint` carries a usage rule that the CSS cannot enforce.** It is
  stated on the token and in the one place that was tempting to get wrong
  (`.dp-web-doc__para--footnote`, which can also be the `--active` paragraph and
  therefore sit on `bg-hover`). A future rule that puts faint text inside a
  hover fill will be below AA and nothing will fail.
- **The two dark blocks are duplicated by design** — one is the pre-paint
  `@media` block, one the explicit override — so they can drift. Guarded: the
  test asserts they declare exactly the same set.
- **RenderCore's preview canvas does not follow the theme.** `.rc-preview`
  paints `--rc-canvas: #e7e8ea` and `--rc-changebar: #808080` as literals in
  `src/shared/render-core/output/preview.ts`, and that file IS in the web
  bundle, so a paginated redline preview sits in a light-grey surround in dark
  mode. Arguably correct — paper is paper, and a redline must not change colour
  with a UI theme — but it is unresolved, it is behind RenderCore Phase 2's
  live-host gate, and it is recorded here rather than changed.
- **The Compare surface still overflows horizontally at ~1280px** (the
  "Revised" slot clips). Pre-existing, a `shell.css` grid, unrelated to tokens
  and untouched.
- **ADR-0012 is now partially superseded** and says so. Its task-pane decisions
  stand in full.
