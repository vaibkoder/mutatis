/**
 * Non-text property changes — the ops that would otherwise not exist.
 * Spec: spec/features/08-compare.md §6 row 8, §9.4, §9.9 · ADR-0042, ADR-0045 § Amendment.
 *
 * The structural diff pairs paragraphs on `hash`, which is prefix-stripped
 * visible TEXT. A paragraph that was outdented — restyled, re-levelled, its
 * number changed from `(vi)` to `(d)` — has the same text on both sides, so it
 * pairs as `equal` and no classifier ever sees it. The user accepts a redline
 * and the change summary reports nothing. Under ADR-0008 that is suppression by
 * omission, and it is the defect this pass closes.
 *
 * ── The rule this pass obeys ─────────────────────────────────────────────────
 * TEXT IDENTITY ANCHORS. NON-TEXT IDENTITY ANNOTATES. This pass reads the
 * pairing the structural diff already made and never changes it: it converts an
 * `equal` op into a `property` op and touches nothing else. No pair is broken,
 * no range is re-cut, and a `property` op authors exactly as `equal` does, so
 * the redline XML is byte-identical. Only `ops` and the summary grow.
 *
 * ONE exception (ADR-0045 § Amendment): a footnote or endnote reference,
 * picture, text box or embedded object added, removed or moved with the words
 * untouched. That is a change to the content, and the redline marks it.
 * ─────────────────────────────────────────────────────────────────────────────
 *
 * Pure: returns a new op list; does not mutate the input ops or the records.
 */
import type { DiffOp, ParaRecord, PropertyChange, ZeroTextObject } from '../model.js';
import { zeroTextObjects } from '../preprocess/inline-units.js';

/**
 * The comparable numbering identity; null/undefined both mean "not numbered".
 *
 * `ParaNumbering.key` and NOT `numId`: see that field's doc comment. A `w:num`
 * id is document-local bookkeeping, and comparing it across two packages
 * reported six changes on a document with none.
 */
function identityKey(rec: ParaRecord): string {
  return rec.numbering ? rec.numbering.key : '';
}

/**
 * The zero-text objects of a paragraph (ADR-0045 § Amendment). A table record
 * has none here: the table path is its own (ADR-0027) and is not touched.
 */
function objectsOf(rec: ParaRecord): ZeroTextObject[] {
  return rec.table ? [] : zeroTextObjects(rec.rawXml, rec.numberingPrefix);
}

/** Same sorts at the same places, in the same order. */
function sameObjects(a: readonly ZeroTextObject[], b: readonly ZeroTextObject[]): boolean {
  return a.length === b.length && a.every((x, i) => x.sort === b[i]!.sort && x.at === b[i]!.at && x.within === b[i]!.within);
}

/**
 * Which properties differ across a matched pair, or null when none do.
 *
 * `numberLabel` is compared only when BOTH sides resolved one. A label that
 * exists on one side and not the other is not evidence of a renumbering: it is
 * evidence that one side's level definition did not resolve, and reporting that
 * as a change would invent one. The identity comparison still covers the real
 * case, because a paragraph that gained or lost numbering changed identity.
 *
 * `objects`: a footnote or endnote reference, picture, text box or embedded
 * object added, removed or moved, with the words untouched (ADR-0045 §
 * Amendment). The text is identical, so this is the only pass that can see it.
 */
export function comparePropertyPair(orig: ParaRecord, rev: ParaRecord): PropertyChange | null {
  const styleId = orig.styleId !== rev.styleId;
  const numberingIdentity = identityKey(orig) !== identityKey(rev);
  const bothLabelled =
    typeof orig.numbering?.label === 'string' && typeof rev.numbering?.label === 'string';
  const numberLabel = bothLabelled && orig.numbering?.label !== rev.numbering?.label;
  const origObjects = objectsOf(orig);
  const revObjects = objectsOf(rev);
  const objects = !sameObjects(origObjects, revObjects);

  if (!styleId && !numberingIdentity && !numberLabel && !objects) return null;
  const change: PropertyChange = {
    styleId,
    numberingIdentity,
    numberLabel,
    // `ilvl` rides along because a label is not always available: a `numId`
    // missing from `numbering.xml`, or an `ilvl` past the levels the
    // `abstractNum` defines, resolves to no label on either side — and then
    // "— → —" is all the reader gets for a real re-levelling. The level is
    // always known, because it is what the style chain resolved.
    from: {
      styleId: orig.styleId,
      numberLabel: orig.numbering?.label ?? null,
      ilvl: orig.numbering?.ilvl ?? null,
    },
    to: {
      styleId: rev.styleId,
      numberLabel: rev.numbering?.label ?? null,
      ilvl: rev.numbering?.ilvl ?? null,
    },
  };
  if (objects) {
    change.objects = true;
    change.from.objects = origObjects;
    change.to.objects = revObjects;
  }

  // A PURE cascade: this paragraph did not change; its number moved because a
  // sibling did. Carry what §6.1's collapse needs, because `cascade.ts` derives
  // a member from `wordDiffs` and a property op has none — without this, one
  // inserted clause puts one uncollapsed row in the change list for every clause
  // that follows it, which is the exact pain §6.1 exists to prevent.
  //
  // Never when an object changed too: that paragraph DID change, and a cascade
  // member is hidden from the list by *Collapse renumbering* (§8.2).
  if (!styleId && !numberingIdentity && numberLabel && !objects) {
    const shift = counterShift(
      orig.numbering?.path,
      rev.numbering?.path,
      orig.numbering?.renders,
    );
    if (shift !== null) {
      change.cascade = {
        from: orig.numbering?.label as string,
        to: rev.numbering?.label as string,
        shiftedIlvl: shift.ilvl,
        shift: shift.by,
      };
    }
  }
  return change;
}

/**
 * WHICH counter moved, and by how much — the first level at which two counter
 * paths disagree.
 *
 * Not the paragraph's own counter. Under `%1.%2`, the SG/E&W house style, a
 * sub-clause renumbers when its PARENT does: insert one top-level clause and
 * `1.1` → `2.1`, `1.2` → `2.2`, … while the level-1 counter never leaves 1.
 * Subtracting own-level ordinals there yields a shift of zero for every one of
 * them; §6.1's collapse drops zero-shift members, so the flood it exists to
 * prevent arrived in full — measured at six uncollapsed rows and zero cascades
 * on a six-sub-clause probe, with the inserted clause's own row alone in a
 * group of one below the threshold.
 *
 * Reading the first differing level fixes both halves at once. All seven rows
 * of that probe now report `(level 0, +1)`, so they group together and collapse
 * as the single fact they are: the top-level numbering moved by one.
 *
 * -- …but only among the levels the LABEL PRINTS (`renders`) ------------------
 * §6.1's header names a level and a shift, and the rows beneath it show the two
 * labels. Scanning the whole path makes the two contradict each other whenever
 * the level that moved is one the template never substitutes. Measured on a
 * clause list split by a new top-level clause with the sub-level templated
 * `(%2)` — the own-level-only shape, and the shape of this suite's own shared
 * fixture:
 *
 *     O: 1. | (a) | (b) | (c) | (d) | (e)
 *     R: 1. | (a) | (b) | 2. | (a) | (b) | (c)
 *     header: "Clause numbering (level 1) shifted (+1) — 3 paragraphs"
 *     rows:   (c) → (a)   (d) → (b)   (e) → (c)
 *
 * Every row moved MINUS TWO at the level it prints; the `+1` is the parent
 * counter, which `(%2)` renders nowhere. Under `%1.%2` the same split reads
 * `1.c → 2.a` and the `+1` is on the page, which is why this went unseen — the
 * two shapes disagree, and only one of them was ever probed.
 *
 * So the scan runs over `renders` (ascending), not over every level of the path.
 * `%1.%2` is unaffected: it prints both levels, so the shallowest differing one
 * is still level 0. `(%2)` now reports `(level 1, -2)`, which is what its rows
 * show. The rule is the one `numberingKey` already follows — what this
 * paragraph's number is rendered from, and nothing else.
 *
 * Returns null when nothing moved (the label changed for some other reason — a
 * `lvlText` edit, say, which is an identity change and does not reach here) or
 * when either side has no path, because a cascade asserted without an arithmetic
 * basis is a guess.
 */
function counterShift(
  origPath: readonly number[] | null | undefined,
  revPath: readonly number[] | null | undefined,
  renders: readonly number[] | null | undefined,
): { ilvl: number; by: number } | null {
  if (!origPath || !revPath) return null;
  // Differing depths mean the paragraph was re-levelled, which changes `key` and
  // so never reaches here; guard anyway rather than compare across levels.
  if (origPath.length !== revPath.length) return null;

  // `renders` comes from ONE side because this branch runs only when the two
  // identities are equal, and the template is part of the identity — so both
  // sides print the same levels. A null or empty set falls through to the full
  // scan below rather than losing the collapse: a label that moved with no
  // printed counter behind it cannot happen (it would be a definition change,
  // which is an identity change and never reaches here), and over-reporting a
  // cascade is the safe direction — the ops all exist either way, collapsed or
  // not (ADR-0008).
  const scan = renders !== null && renders !== undefined && renders.length > 0
    ? renders
    : origPath.map((_, i) => i);

  for (const i of scan) {
    const from = origPath[i];
    const to = revPath[i];
    if (from === undefined || to === undefined) continue;
    if (from !== to) return { ilvl: i, by: to - from };
  }
  return null;
}

/**
 * Convert every `equal` op whose two paragraphs differ in a non-text property
 * into a `property` op. Everything else passes through untouched.
 *
 * Only ONE-TO-ONE `equal` ops are examined. An `equal` op always pairs a single
 * paragraph with a single paragraph (the structural diff emits one per matched
 * record); a range-spanning op would be a merge or split, which is a text
 * change and already has its own op.
 *
 * Classification is left to the taxonomy (`classifyClasses`), which reads
 * `propertyChange` — this pass decides WHAT differs, not what it MEANS.
 */
export function detectPropertyChanges(
  ops: DiffOp[],
  original: ParaRecord[],
  revised: ParaRecord[],
): DiffOp[] {
  const out: DiffOp[] = [];
  for (const op of ops) {
    if (op.type !== 'equal' || !op.origRange || !op.revRange) {
      out.push(op);
      continue;
    }
    if (op.origRange.end !== op.origRange.start || op.revRange.end !== op.revRange.start) {
      out.push(op);
      continue;
    }
    const orig = original[op.origRange.start];
    const rev = revised[op.revRange.start];
    if (orig === undefined || rev === undefined) {
      out.push(op);
      continue;
    }

    const propertyChange = comparePropertyPair(orig, rev);
    if (propertyChange === null) {
      out.push(op);
      continue;
    }

    out.push({ ...op, type: 'property', propertyChange });
  }
  return out;
}

