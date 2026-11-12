/**
 * host-web — rendered list-number computer (spec/features/12 §4 hard-part #1).
 *
 * Word *renders* `"4.2.1"` / `"(b)"`; raw OOXML only carries `<w:numPr>` references
 * into `word/numbering.xml`. This module is the counter state machine that computes
 * those rendered strings for the web host.
 *
 * ── CONTRACT (pinned by tests/unit/host-web/numbering-render.todo.test.ts) ──────
 * Input: a parsed `NumberingScheme` and `parasInOrder` — every list paragraph in
 * DOCUMENT ORDER (already filtered: non-list paras excluded; see ooxml-loader).
 * Output: a Map from `paraId` to its rendered `{ listString, listLevel }`.
 *
 * Algorithm (per `(counter group, level)` counters — the group being the ABSTRACT
 * DEFINITION an instance resolves to, not its `w:numId` — walked in document order):
 *  1. For each paragraph, using its `(numId, ilvl)`:
 *       a. Resolve the effective level via `resolveLevel(scheme, numId, ilvl)`.
 *          If undefined (unknown numId/level), skip — emit nothing for that para.
 *       b. ADVANCE this level: increment its counter (first sighting starts at the
 *          level's `start`). Counters belong to the ABSTRACT DEFINITION the
 *          instance resolves to, not to the `w:numId` — two `<w:num>` over one
 *          `<w:abstractNum>` continue ONE sequence, which is the whole reason
 *          `<w:startOverride>` exists: it would be a no-op if a fresh instance
 *          restarted by itself. That override restarts the shared counter on the
 *          instance's FIRST TOUCH of the level (`counterGroupOf`, `takeRestart`) —
 *          touch and not advance, because reading a level as an ancestor token is
 *          as much a use of the instance as sitting at it, and treating only the
 *          latter made a restarted instance whose first paragraph is deeper render
 *          the counter it was restarting away from.
 *       c. RESET all DEEPER levels (ilvl+1…) of the same COUNTER GROUP to "uninitialised"
 *          so each re-enters at its own `start` next time it appears — RESPECTING
 *          `<w:lvlRestart>` (ECMA-376 §17.9.10, one-based: a level with
 *          `lvlRestart=k` restarts only when a level with one-based index ≤ k —
 *          i.e. ilvl ≤ k-1 — advances; `lvlRestart=0` ⇒ never auto-restarts;
 *          absent ⇒ restarts when ANY shallower level advances).
 *       d. SKIPPED shallower levels (a doc that jumps straight to ilvl=2 without a
 *          0 or 1) initialise to their `start` when first needed for a token.
 *  2. Build the per-level token array for levels 0..ilvl: `formatInt(counter_k,
 *     fmt_k)` where `fmt_k` is each ancestor level's `numFmt` — UNLESS the current
 *     level has `<w:isLgl/>`, in which case render ALL substituted levels as
 *     `decimal` (legal numbering), regardless of each level's own format.
 *     **The word ALL is a known DEFECT, not the contract** — see the fidelity
 *     bounds below. Word converts the ANCESTOR tokens and keeps the current
 *     level's own declared format: `numFmt="decimalZero"` under `<w:isLgl/>`
 *     renders `1.01`, and this renders `1.1`.
 *  3. `listString = applyLvlText(level.lvlText, perLevelTokens)`;
 *     `listLevel = ilvl`.
 *
 * Known fidelity bounds — and READ THE HISTORY BEFORE ADDING TO THIS LIST. Three
 * of the four ADR-0008 suppressions encountered were items filed here,
 * on the reasoning that a rendering bound cannot affect a comparison. It can:
 * Compare's `numberingKey` and its rendered-label comparison both resolve through
 * this module, so a bound recorded here is a bound on what Compare can SEE.
 * **A bound is not "confined to the renderer" while anything else shares its
 * resolver** (passes 22, 24 and 25).
 *
 *  - **`<w:isLgl/>` is applied one level too far** (step 2 above) — Word converts
 *    the ancestor tokens and keeps the own level's declared format. 08c's to fix,
 *    and the reason Compare's key carries the own level's DECLARED `numFmt`
 *    beside the effective one rather than mirroring this module's arithmetic:
 *    mirroring it hid a real edit outright (pass 22).
 *  - **Style-linked numbering is no longer on this list.** `resolveLevel` follows
 *    `<w:numStyleLink>` to the definition carrying the matching `<w:styleLink>`
 *    (pass 24), because while it did not, a redefinition of a list attached to a
 *    Word list style produced NO OP AT ALL. What remains genuinely upstream is a
 *    different thing that has been called by the same name: RenderCore's
 *    `parse-body.ts` does not yet read the `pStyle` → `numPr` chain at all (08c).
 *  - **Counters are no longer keyed per `numId`** (pass 25) — they belong to the
 *    abstract definition, which is what `<w:startOverride>` presupposes. While
 *    they were, two `<w:num>` over one `<w:abstractNum>` restarted independently:
 *    an insertion into the first group moved the second group's numbers in Word
 *    and reported NOTHING here (`summary.total` 1 against 3 for the identical
 *    edit on a single instance), and re-pointing paragraphs at a second instance
 *    invented a cascade of numbers Word never renders.
 *  - **Bullet glyphs and footnote-body lists** remain upstream/deferred, and both
 *    have now been MEASURED against Compare (pass 26) rather than assumed:
 *      · a bullet **glyph** written in `w:lvlText` is compared like any other
 *        template, so changing it reports two `formatting` ops;
 *      · a **footnote** paragraph's numbering is not compared at all, because the
 *        parts diff never runs `detectPropertyChanges` — already recorded in
 *        ADR-0042 § Open with its cost, and the measurement confirms it;
 *      · but a glyph changed through the level's `<w:rPr><w:rFonts>` (Symbol →
 *        Wingdings, `lvlText` untouched) produces `summary.total` 0. That is the
 *        general shape below, and it is not this module's to fix.
 *  - **`numbering.xml` is never diffed as a PART**, which `styles.xml` is. So a
 *    list-definition change is visible only through the fields Compare's
 *    per-paragraph key carries, and everything outside it — a level's `w:rPr`
 *    (bullet font, size, colour), its `w:pPr`/`w:ind` hanging gutter, `w:suff`,
 *    and any level no paragraph currently uses — produces no output at all, not
 *    even the §9.9-style *"definition changed"* note the same edit to `styles.xml`
 *    would produce. A detection limit, recorded in ADR-0042 § Open with its
 *    measurement rather than left for the next pass to rediscover.
 *
 * Host-NEUTRAL: no `Office.*`/`Word.*` token.
 */
import {
  type LevelDef,
  type NumberingScheme,
  resolveLevel,
  counterGroupOf,
  startOverrideFor,
} from './numbering-scheme.js';
import { applyLvlText, formatInt } from './numbering-format.js';

/**
 * A paragraph's numbering reference, extracted from `<w:pPr><w:numPr>` by the
 * loader. Emitted in document order, LIST PARAS ONLY (numId "0"/missing excluded).
 */
export interface ParaNumPr {
  /** The owning paragraph's canonical id (e.g. `body-para-12`). */
  paraId: string;
  /** `<w:numId w:val>` — the numbering instance. */
  numId: string;
  /** `<w:ilvl w:val>` — 0-based level (defaults to 0 when absent). */
  ilvl: number;
}

/** A rendered list number for one paragraph. */
export interface ListNumber {
  /** The rendered prefix, e.g. "4.2.1", "(b)". */
  listString: string;
  /** The 0-based list level (== the paragraph's ilvl). */
  listLevel: number;
  /**
   * The counter values behind the rendered token, for levels 0…ilvl — the
   * ordinals before `numFmt` turns 4 into `(d)` or `iv`. `[2, 1]` for `"2.1"`.
   *
   * Surfaced because a rendered label cannot be subtracted: telling a
   * renumbering cascade of +1 from one of +2 means comparing ordinals, and
   * `(d)` → `(e)` carries no arithmetic (Compare §6.1, ADR-0042).
   *
   * The whole PATH and not just the own-level counter, because in the dominant
   * SG/E&W house style (`lvlText` `%1.%2`) a sub-clause's number moves without
   * its own counter moving at all: insert one top-level clause and `1.1` → `2.1`
   * while the level-1 counter stays 1. An own-level ordinal reports a shift of
   * ZERO there, §6.1's collapse drops the member, and every sub-clause in the
   * document lands in the change list as its own uncollapsed row — measured at
   * six rows on a six-sub-clause probe. The path says which level moved.
   */
  path: readonly number[];
}

/**
 * Compute rendered list numbers for the given list paragraphs.
 * Paragraphs whose `(numId, ilvl)` does not resolve to a level definition are
 * absent from the result (the loader then leaves them unnumbered).
 */
export function renderListNumbers(
  scheme: NumberingScheme,
  parasInOrder: ParaNumPr[],
): Map<string, ListNumber> {
  const out = new Map<string, ListNumber>();
  // Current counter value per (COUNTER GROUP, ilvl) — the group being the
  // abstract definition the instance resolves to, NOT the `w:numId`. See
  // `counterGroupOf`. A missing ilvl entry means "uninitialised" — the level
  // (re-)enters at its resolved `start`.
  const counters = new Map<string, Map<number, number>>();
  // `${numId}:${ilvl}` seen already — a `<w:startOverride>` restarts the shared
  // counter on the instance's FIRST use of that level and never again.
  const instanceLevelSeen = new Set<string>();

  for (const para of parasInOrder) {
    const level = resolveLevel(scheme, para.numId, para.ilvl);
    if (!level) continue;

    const group = counterGroupOf(scheme, para.numId);
    let numCounters = counters.get(group);
    if (!numCounters) {
      numCounters = new Map<number, number>();
      counters.set(group, numCounters);
    }

    // ADVANCE this level — unless this instance RESTARTS it, which is what
    // `<w:startOverride>` means and the only reason Word writes one.
    const restart = takeRestart(scheme, instanceLevelSeen, para.numId, para.ilvl);
    const current = numCounters.get(para.ilvl);
    if (restart !== undefined) {
      numCounters.set(para.ilvl, restart);
    } else {
      numCounters.set(para.ilvl, current === undefined ? level.start : current + 1);
    }

    // RESET deeper levels of this counter group, respecting <w:lvlRestart>.
    for (const deeperIlvl of [...numCounters.keys()]) {
      if (deeperIlvl <= para.ilvl) continue;
      const deeperDef = resolveLevel(scheme, para.numId, deeperIlvl);
      if (restartsOnAdvanceOf(deeperDef, deeperIlvl, para.ilvl)) {
        numCounters.delete(deeperIlvl);
      }
    }

    // Per-level tokens for %1..%(ilvl+1); skipped shallower levels initialise to
    // their own `start` and KEEP that value (so a later advance continues from it).
    const tokens: string[] = [];
    const path: number[] = [];
    for (let k = 0; k <= para.ilvl; k++) {
      const kDef = k === para.ilvl ? level : resolveLevel(scheme, para.numId, k);
      // An ancestor token is this instance's first USE of that level just as
      // much as advancing it is, so the same restart applies. Without this the
      // two halves disagreed: a `<w:startOverride>` instance whose first
      // paragraph sits DEEPER than the overridden level read the previous
      // instance's counter for its ancestor token and restarted only when it
      // later reached that level itself — `1. 2. 2.a 1. 2.` where Word restarts
      // the whole instance, and `2.a 2.b` for an instance that never uses the
      // overridden level at all. One rule, both roles: the first time this
      // instance touches level k, its override (if any) sets the counter.
      let value = k === para.ilvl
        ? numCounters.get(k)
        : takeRestart(scheme, instanceLevelSeen, para.numId, k);
      if (value !== undefined && k !== para.ilvl) numCounters.set(k, value);
      if (value === undefined) value = numCounters.get(k);
      if (value === undefined) {
        value = kDef?.start ?? 1;
        numCounters.set(k, value);
      }
      const fmt = level.isLgl ? 'decimal' : (kDef?.numFmt ?? 'decimal');
      tokens.push(formatInt(value, fmt));
      path.push(value);
    }

    out.set(para.paraId, {
      listString: applyLvlText(level.lvlText, tokens),
      listLevel: para.ilvl,
      path,
    });
  }

  return out;
}

/**
 * The value this instance restarts level `ilvl` at, if it does — consumed ONCE.
 *
 * `<w:startOverride>` is a restart, not a per-paragraph start value: an instance
 * that leaves a level and returns to it does not restart a second time. So the
 * first touch by this instance takes the override and every later one advances
 * normally, and `seen` is what makes "first" mean first — whether the touch
 * ADVANCES the level (the paragraph sits at it) or merely READS it (the level is
 * an ancestor of the paragraph's own). Both are uses of the instance; treating
 * only the first as one made a restarted instance render its ancestor token from
 * the counter it was restarting away from.
 *
 * Returns undefined when this instance declares no override for the level, or
 * when it has already been consumed — the caller then advances or initialises as
 * usual. MUTATES `seen`, which is why it is `take`.
 */
function takeRestart(
  scheme: NumberingScheme,
  seen: Set<string>,
  numId: string,
  ilvl: number,
): number | undefined {
  const key = `${numId}:${ilvl}`;
  const first = !seen.has(key);
  seen.add(key);
  if (!first) return undefined;
  return startOverrideFor(scheme, numId, ilvl);
}

// ---------------------------------------------------------------------------
// Internals
// ---------------------------------------------------------------------------

/**
 * Does the level at `deeperIlvl` (definition `def`) restart when a paragraph at
 * `advancedIlvl` (< deeperIlvl) advances? ECMA-376 §17.9.10 `<w:lvlRestart>`,
 * one-based: val 0 = never; val k ≥ 1 = restart when the advancing level's
 * one-based index ≤ k (ilvl ≤ k-1); absent = restart on any shallower advance
 * (equivalent to val = deeperIlvl, the one-based index of the level above).
 */
function restartsOnAdvanceOf(
  def: LevelDef | undefined,
  deeperIlvl: number,
  advancedIlvl: number,
): boolean {
  const restart = def?.lvlRestart ?? deeperIlvl;
  if (restart === 0) return false;
  return advancedIlvl <= restart - 1;
}
