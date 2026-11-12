/**
 * host-web — thin `word/numbering.xml` IR + parser (spec/features/12 §4 hard-part #1).
 *
 * Word *renders* `"4.2.1"` / `"(b)"`; raw OOXML only has `<w:numPr>` references
 * (`numId` + `ilvl`) into `word/numbering.xml`. This module parses that part into a
 * THIN, OOXML-close intermediate representation and provides the pure level lookup
 * (`resolveLevel`) the (later) `renderListNumbers` counter state machine walks.
 *
 * The IR is deliberately close to the XML and NOT over-abstracted — a later session
 * is licensed to amend these types if the counter algorithm needs it.
 *
 * Counter-free: this module only *parses* and *resolves a level definition*. It
 * computes no running numbers — that is the deferred state machine (`render-list-
 * numbers.ts`).
 *
 * Host-NEUTRAL: no `Office.*`/`Word.*` token.
 *
 * numbering.xml shape (abbreviated):
 *   <w:numbering>
 *     <w:abstractNum w:abstractNumId="0">
 *       <w:lvl w:ilvl="0">
 *         <w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>
 *         <w:lvlRestart w:val="0"/>   (optional)  <w:isLgl/>  (optional)
 *       </w:lvl> …
 *     </w:abstractNum>
 *     <w:abstractNum w:abstractNumId="1">        (a list-STYLE definition)
 *       <w:numStyleLink w:val="HouseList"/>      delegates; holds no levels
 *     </w:abstractNum>                           -> the abstractNum whose
 *                                                   <w:styleLink> matches
 *     <w:num w:numId="1">
 *       <w:abstractNumId w:val="0"/>
 *       <w:lvlOverride w:ilvl="0">
 *         <w:startOverride w:val="5"/>            (optional)
 *         <w:lvl w:ilvl="0">…</w:lvl>             (optional inline level)
 *       </w:lvlOverride>
 *     </w:num>
 *   </w:numbering>
 */
import type { DocxBundle } from '@shared/ooxml/index.js';

import { type PoNode, parsePreserveOrder, tagOf, childrenOf, attrOf } from '@shared/ooxml/po-xml.js';

/** One level of an abstract numbering definition. Close to `<w:lvl>`. */
export interface LevelDef {
  /** 0-based level (`<w:lvl w:ilvl>`). */
  ilvl: number;
  /** `<w:numFmt w:val>` (e.g. "decimal", "lowerLetter", "lowerRoman", "bullet"). */
  numFmt: string;
  /** `<w:lvlText w:val>` template (e.g. "%1.", "(%2)"). */
  lvlText: string;
  /** `<w:start w:val>`; defaults to 1 when absent. */
  start: number;
  /** `<w:lvlRestart w:val>` — restart this level when a level ≤ this value advances. Absent if unset. */
  lvlRestart?: number;
  /**
   * `<w:isLgl/>` present — the "legal numbering" flag.
   *
   * `render-list-numbers.ts` renders ALL substituted levels as decimal under it,
   * **including the paragraph's own token, which is a known DEFECT and not the
   * contract**: Word converts the ANCESTOR tokens and keeps the own level's
   * declared format, so `numFmt="decimalZero"` under `<w:isLgl/>` renders `1.01`
   * where this build renders `1.1`. 08c's to fix; recorded in that module's
   * fidelity bounds, in `spec/features/16-render-core.md`, and in ADR-0042 (which
   * is why Compare's comparison key carries the own level's DECLARED format as
   * well as the folded one — a key that mirrors a renderer inherits its blind
   * spots, and that one hid a real edit outright).
   */
  isLgl: boolean;
}

/** An abstract numbering definition (`<w:abstractNum>`): its levels keyed by ilvl. */
export interface AbstractNum {
  abstractNumId: string;
  levels: Map<number, LevelDef>;
  /**
   * `<w:numStyleLink w:val>` (ECMA-376 §17.9.21) — this definition holds no
   * numbering properties of its own and DELEGATES to the numbering style named
   * here. Word writes this for every list built from a list style, which is how
   * a firm template usually carries its house numbering.
   */
  numStyleLink?: string;
  /**
   * `<w:styleLink w:val>` (§17.9.28) — the other end: this definition IS the one
   * backing the numbering style named here. `resolveLevel` follows a
   * `numStyleLink` to the definition whose `styleLink` matches.
   */
  styleLink?: string;
}

/** A concrete numbering instance (`<w:num>`) referencing an abstractNum, with per-level overrides. */
export interface NumInstance {
  numId: string;
  abstractNumId: string;
  /** `<w:lvlOverride>` by ilvl: a start-value override and/or a full inline level replacement. */
  lvlOverrides: Map<number, { startOverride?: number; lvl?: LevelDef }>;
}

/** The whole parsed `numbering.xml`: instances + abstract definitions. */
export interface NumberingScheme {
  nums: Map<string, NumInstance>;
  abstractNums: Map<string, AbstractNum>;
}

/** An empty scheme — returned when `numbering.xml` is absent or has no root. */
function emptyScheme(): NumberingScheme {
  return { nums: new Map(), abstractNums: new Map() };
}

/**
 * Parse `word/numbering.xml` from a bundle into the IR.
 * Absent part (or no `<w:numbering>` root) → empty scheme (every paragraph then
 * resolves to non-list, as today).
 */
export function parseNumberingScheme(bundle: DocxBundle): NumberingScheme {
  return parseNumberingSchemeXml(bundle.part('word/numbering.xml'));
}

/**
 * Parse `numbering.xml` from the RAW PART STRING. The bundle-taking
 * `parseNumberingScheme` delegates here; callers that already hold the part as a
 * string (Compare reads its parts through `io/parts.ts`, not through a
 * `DocxBundle`) use this one rather than synthesising a bundle to get at it.
 * Absent/empty part → empty scheme.
 */
export function parseNumberingSchemeXml(xml: string | null | undefined): NumberingScheme {
  if (!xml) return emptyScheme();

  const root = findElement(parsePreserveOrder(xml), 'w:numbering');
  if (!root) return emptyScheme();

  const scheme = emptyScheme();
  for (const child of childrenOf(root)) {
    const tag = tagOf(child);
    if (tag === 'w:abstractNum') {
      const a = parseAbstractNum(child);
      if (a) scheme.abstractNums.set(a.abstractNumId, a);
    } else if (tag === 'w:num') {
      const n = parseNumInstance(child);
      if (n) scheme.nums.set(n.numId, n);
    }
  }
  return scheme;
}

/**
 * Resolve the effective `LevelDef` for `(numId, ilvl)`, applying any `<w:lvlOverride>`
 * on top of the abstract level: an inline `<w:lvl>` replaces the level wholesale; a
 * `<w:startOverride>` overrides the start value. PURE lookup — no counters.
 * Returns undefined when the numId is unknown or the level is undefined.
 */
export function resolveLevel(
  scheme: NumberingScheme,
  numId: string,
  ilvl: number,
): LevelDef | undefined {
  const num = scheme.nums.get(numId);
  if (!num) return undefined;

  const override = num.lvlOverrides.get(ilvl);
  const base = override?.lvl ?? levelFrom(scheme, num.abstractNumId, ilvl);
  if (!base) return undefined;

  if (override?.startOverride !== undefined) {
    return { ...base, start: override.startOverride };
  }
  return base;
}

/**
 * One abstract definition's level, following `<w:numStyleLink>` when the
 * definition delegates to a numbering style (ECMA-376 §17.9.21 / §17.9.28).
 *
 * **Why this exists, measured.** A list built through a Word LIST STYLE — which
 * is how a firm template usually carries house numbering — puts the levels on a
 * second `<w:abstractNum>` carrying `<w:styleLink>`, and leaves the one the
 * `<w:num>` points at holding nothing but `<w:numStyleLink>`. Without the hop
 * every such paragraph resolved to NO level, which meant no rendered label and a
 * comparison key of the bare level. Compare then reported **nothing at all** for
 * a list redefinition that moves every number on the page:
 *
 *     decimal -> lowerRoman, 3 clauses, bodies identical
 *       levels on the abstractNum   labels 1. 2. 3. -> i. ii. iii.   total 3
 *       levels behind a list style  labels null     -> null          total 0
 *
 * and lost the whole renumbering cascade when a clause was inserted above it
 * (total 4 with 1 cascade, against 1 with none). ADR-0008 forbids that, and it is
 * the third such suppression this work has produced: a defect recorded as a
 * RENDERER bound is not confined to the renderer while the comparison key
 * resolves through the same module (ADR-0042).
 *
 * **The hop is resolved inside `numbering.xml`, not through `styles.xml`.**
 * §17.9.21 names a numbering STYLE, whose `<w:pPr><w:numPr>` points at a
 * `<w:num>` whose definition carries the levels — and that definition names the
 * same style back in `<w:styleLink>`, which Word always writes as the pair. So
 * the round trip is observable from this part alone, and this module keeps its
 * signature (it parses one part; it does not acquire a second input). A
 * `numStyleLink` with no matching `styleLink` anywhere resolves to undefined
 * exactly as before.
 *
 * **Additive by construction:** a definition that HAS the level keeps it, so no
 * resolution that worked before this hop changes. Only previously-unresolvable
 * ones move, which is what makes this safe for the renderer's gated consumers.
 * Word does not write both, so the precedence is unobservable in practice.
 *
 * Cycle-safe: `numStyleLink` chains are followed through a `seen` set.
 */
function levelFrom(
  scheme: NumberingScheme,
  abstractNumId: string,
  ilvl: number,
): LevelDef | undefined {
  const seen = new Set<string>();
  let current = scheme.abstractNums.get(abstractNumId);
  while (current !== undefined && !seen.has(current.abstractNumId)) {
    seen.add(current.abstractNumId);
    const own = current.levels.get(ilvl);
    if (own) return own;
    const link = current.numStyleLink;
    if (link === undefined) return undefined;
    current = findByStyleLink(scheme, link);
  }
  return undefined;
}

/**
 * The COUNTER GROUP an instance belongs to: the abstract definition its levels
 * are ultimately read from, delegation followed.
 *
 * Counters belong to the definition, not to the `<w:num>`. That is not a reading
 * of Word's behaviour but of the format's own design: `<w:lvlOverride>`'s
 * `<w:startOverride>` exists to restart a level for one instance, and it would be
 * a no-op if every fresh `<w:num>` restarted by itself. Word writes a second
 * `<w:num>` over the same `<w:abstractNum>` WITH a `startOverride` when the user
 * asks to restart, and WITHOUT one when the user asks to continue.
 *
 * Keying counters by `numId` instead broke both directions, measured through the
 * whole Compare pipeline:
 *
 *   - a clause inserted into the first of two groups sharing one definition moved
 *     the second group's numbers in Word and produced NO OP AT ALL here —
 *     `summary.total` 1 against 3 for the identical edit on a single instance;
 *   - and re-pointing paragraphs to a second instance of the same definition
 *     reported a renumbering cascade — with the numbers it invented — where Word
 *     renders no change.
 *
 * Falls back to the instance's own id when nothing resolves, so an unknown
 * `numId` still gets its own counters rather than sharing a bucket with every
 * other unknown one. The `num:` prefix keeps that namespace apart from the
 * abstract ids.
 */
export function counterGroupOf(scheme: NumberingScheme, numId: string): string {
  const num = scheme.nums.get(numId);
  if (!num) return `num:${numId}`;
  const seen = new Set<string>();
  let current = scheme.abstractNums.get(num.abstractNumId);
  while (current !== undefined && !seen.has(current.abstractNumId)) {
    seen.add(current.abstractNumId);
    // A delegating definition holds no levels: the sequence belongs to the one
    // the list style points at, so two instances reaching it by different routes
    // are one list.
    if (current.levels.size > 0 || current.numStyleLink === undefined) {
      return current.abstractNumId;
    }
    current = findByStyleLink(scheme, current.numStyleLink);
  }
  return num.abstractNumId === '' ? `num:${numId}` : num.abstractNumId;
}

/**
 * The `<w:startOverride>` this instance declares for a level, if any — the value
 * the shared counter RESTARTS at when the instance first TOUCHES that level.
 *
 * Touch, not *use*: a paragraph sitting at the level and a paragraph reading it
 * as an ancestor token are both uses of the instance, and the caller
 * (`takeRestart`) consumes this once across both. Reading "use" as only the first
 * of those is a measured defect — see that function.
 *
 * `resolveLevel` also folds this into the level's `start`, which covers the case
 * where the shared counter has not been initialised yet. This accessor is for the
 * case that per-`numId` keying used to give for free and shared keying must do
 * explicitly: an instance that restarts a level the previous instance already
 * advanced.
 */
export function startOverrideFor(
  scheme: NumberingScheme,
  numId: string,
  ilvl: number,
): number | undefined {
  return scheme.nums.get(numId)?.lvlOverrides.get(ilvl)?.startOverride;
}

/** The abstract definition backing a numbering style, by its `<w:styleLink>`. */
function findByStyleLink(scheme: NumberingScheme, styleLink: string): AbstractNum | undefined {
  for (const abstract of scheme.abstractNums.values()) {
    if (abstract.styleLink === styleLink) return abstract;
  }
  return undefined;
}

// ---------------------------------------------------------------------------
// Internals
// ---------------------------------------------------------------------------

function parseAbstractNum(node: PoNode): AbstractNum | undefined {
  const abstractNumId = attrOf(node, '@_w:abstractNumId');
  if (abstractNumId === undefined) return undefined;
  const levels = new Map<number, LevelDef>();
  let numStyleLink: string | undefined;
  let styleLink: string | undefined;
  for (const child of childrenOf(node)) {
    const tag = tagOf(child);
    if (tag === 'w:lvl') {
      const lvl = parseLevel(child);
      if (lvl) levels.set(lvl.ilvl, lvl);
    } else if (tag === 'w:numStyleLink') {
      numStyleLink = attrOf(child, '@_w:val') ?? undefined;
    } else if (tag === 'w:styleLink') {
      styleLink = attrOf(child, '@_w:val') ?? undefined;
    }
  }
  return {
    abstractNumId,
    levels,
    ...(numStyleLink !== undefined ? { numStyleLink } : {}),
    ...(styleLink !== undefined ? { styleLink } : {}),
  };
}

function parseLevel(node: PoNode): LevelDef | undefined {
  const ilvl = parseIntAttr(node, '@_w:ilvl');
  if (ilvl === undefined) return undefined;

  let numFmt = 'decimal'; // OOXML default when <w:numFmt> is absent
  let lvlText = '';
  let start = 1;
  let lvlRestart: number | undefined;
  let isLgl = false;

  for (const child of childrenOf(node)) {
    switch (tagOf(child)) {
      case 'w:numFmt': {
        const v = attrOf(child, '@_w:val');
        if (v !== undefined) numFmt = v;
        break;
      }
      case 'w:lvlText': {
        const v = attrOf(child, '@_w:val');
        if (v !== undefined) lvlText = v;
        break;
      }
      case 'w:start': {
        const v = parseIntVal(attrOf(child, '@_w:val'));
        if (v !== undefined) start = v;
        break;
      }
      case 'w:lvlRestart': {
        const v = parseIntVal(attrOf(child, '@_w:val'));
        if (v !== undefined) lvlRestart = v;
        break;
      }
      case 'w:isLgl':
        isLgl = true;
        break;
    }
  }

  return {
    ilvl,
    numFmt,
    lvlText,
    start,
    isLgl,
    ...(lvlRestart !== undefined ? { lvlRestart } : {}),
  };
}

function parseNumInstance(node: PoNode): NumInstance | undefined {
  const numId = attrOf(node, '@_w:numId');
  if (numId === undefined) return undefined;

  let abstractNumId = '';
  const lvlOverrides = new Map<number, { startOverride?: number; lvl?: LevelDef }>();

  for (const child of childrenOf(node)) {
    const tag = tagOf(child);
    if (tag === 'w:abstractNumId') {
      const v = attrOf(child, '@_w:val');
      if (v !== undefined) abstractNumId = v;
    } else if (tag === 'w:lvlOverride') {
      const ov = parseLvlOverride(child);
      if (ov) lvlOverrides.set(ov.ilvl, ov.value);
    }
  }

  return { numId, abstractNumId, lvlOverrides };
}

function parseLvlOverride(
  node: PoNode,
): { ilvl: number; value: { startOverride?: number; lvl?: LevelDef } } | undefined {
  const ilvl = parseIntAttr(node, '@_w:ilvl');
  if (ilvl === undefined) return undefined;

  let startOverride: number | undefined;
  let lvl: LevelDef | undefined;
  for (const child of childrenOf(node)) {
    const tag = tagOf(child);
    if (tag === 'w:startOverride') {
      const v = parseIntVal(attrOf(child, '@_w:val'));
      if (v !== undefined) startOverride = v;
    } else if (tag === 'w:lvl') {
      lvl = parseLevel(child);
    }
  }

  return {
    ilvl,
    value: {
      ...(startOverride !== undefined ? { startOverride } : {}),
      ...(lvl !== undefined ? { lvl } : {}),
    },
  };
}

/** Recursively find the first element with the given tag (skips the `<?xml?>` decl). */
function findElement(nodes: PoNode[], tag: string): PoNode | undefined {
  for (const node of nodes) {
    if (tagOf(node) === tag) return node;
    const found = findElement(childrenOf(node), tag);
    if (found) return found;
  }
  return undefined;
}

/** Parse an integer attribute; undefined when absent or non-numeric. */
function parseIntAttr(node: PoNode, name: string): number | undefined {
  return parseIntVal(attrOf(node, name));
}

function parseIntVal(v: string | undefined): number | undefined {
  if (v === undefined) return undefined;
  const n = parseInt(v, 10);
  return Number.isNaN(n) ? undefined : n;
}
