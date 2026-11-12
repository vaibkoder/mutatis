/**
 * Feature 08 — Compare · renumbering-cascade collapse (Task B3).
 *
 * Inserting one clause shifts every downstream cross-reference, producing
 * hundreds of "Clause 4.2 → Clause 4.3" changes that bury the one substantive
 * edit. This module groups consistent shifts into a single CascadeAnnotation
 * and tags each member op with `collapsedInto` so the NAVIGATION LIST can show
 * them as one row.
 *
 * ── ADR-0008 (NON-NEGOTIABLE) ────────────────────────────────────────────────
 * Collapse is PRESENTATION ONLY. `computeCascades` MUST NOT remove any DiffOp
 * from the returned `ops` array. Every input op is present in the output;
 * `collapsedInto` is a hint, never a drop. The classifier categorises; it never
 * suppresses. This is enforced by construction here (we map over the input and
 * never filter it) and asserted explicitly by the suppression-guard test.
 * ─────────────────────────────────────────────────────────────────────────────
 *
 * Spec: spec/features/08-compare.md §6.1.
 */

import type { CompareConfig } from '@shared/config/index.js';
import type { CascadeAnnotation, DiffOp, WordDiff } from '../model.js';

/** A parsed reference token: its kind ("Clause"/"Schedule"/"paragraph") and trailing number. */
interface ParsedRef {
  /** Leading reference word, e.g. "Clause". null if not a recognised reference. */
  tokenKind: string | null;
  /** The trailing integer used for shift computation, e.g. 3 in "4.3" / "Schedule 3". null if none. */
  trailingNumber: number | null;
  /** The reference text, trimmed, for example-ref formatting (e.g. "4.2"). */
  refBody: string;
}

/**
 * Parse a clauseRef token into its kind and trailing number.
 * Examples: "Clause 4.2" → { kind:'Clause', trailing:2, body:'4.2' };
 *           "Schedule 3" → { kind:'Schedule', trailing:3, body:'3' };
 *           "4.3"        → { kind:null, trailing:3, body:'4.3' }.
 */
function parseRef(text: string): ParsedRef {
  const trimmed = text.trim();
  // Leading reference word (Clause / Schedule / paragraph / etc.).
  const wordMatch = /^([A-Za-z]+)\b\s*(.*)$/.exec(trimmed);
  let tokenKind: string | null = null;
  let body = trimmed;
  if (wordMatch) {
    const word = wordMatch[1] ?? '';
    const rest = wordMatch[2] ?? '';
    if (/[A-Za-z]/.test(word) && /\d/.test(rest)) {
      tokenKind = word;
      body = rest.trim();
    }
  }
  // Trailing integer of the reference body (last run of digits).
  const numMatches = body.match(/\d+/g);
  const lastNum = numMatches && numMatches.length > 0 ? numMatches[numMatches.length - 1] : undefined;
  const trailingNumber = lastNum !== undefined ? Number.parseInt(lastNum, 10) : null;
  return { tokenKind, trailingNumber, refBody: body };
}

/** The non-equal clauseRef del/ins pair of a crossRefRenumber op, if present. */
function refPair(op: DiffOp): { del: WordDiff; ins: WordDiff; delAt: number } | null {
  const diffs = op.wordDiffs ?? [];
  const delAt = diffs.findIndex((wd) => wd.atomic === 'clauseRef' && wd.op === 'del');
  const del = delAt >= 0 ? diffs[delAt] : undefined;
  const ins = diffs.find((wd) => wd.atomic === 'clauseRef' && wd.op === 'ins');
  if (!del || !ins) return null;
  return { del, ins, delAt };
}

/**
 * The reference noun governing a value token — `clause` in `clause 4.2`.
 *
 * The tokenizer emits the noun as an ordinary word so that a renumbering marks
 * only the number and leaves the noun, identical on both sides, outside the mark
 * (P1). The kind therefore has to be read back off the unchanged text just
 * before the mark. Returns the profile's own spelling of the token, so `Clause`
 * and `clause` land in one cascade rather than two.
 */
function refKindBefore(
  diffs: readonly WordDiff[],
  index: number,
  config: CompareConfig,
): string | null {
  for (let i = index - 1; i >= 0; i--) {
    const wd = diffs[i];
    if (wd === undefined || wd.op === 'ins') continue;
    const word = /(\p{L}+)\s*$/u.exec(wd.text)?.[1];
    if (word === undefined) return null;
    return config.referenceTokens.find((t) => t.toLowerCase() === word.toLowerCase()) ?? null;
  }
  return null;
}

interface CascadeMember {
  op: DiffOp;
  tokenKind: string;
  shift: number;
  example: string;
}

/** Derive the (tokenKind, shift, example) for a crossRefRenumber op, or null. */
/**
 * A cascade member derived from a `property` op (ADR-0042) — the clause NUMBER
 * moved, not a reference to it.
 *
 * `deriveMember` below reads `op.wordDiffs`, and a property op has none: its two
 * paragraphs are word-for-word identical. Without this branch, inserting one
 * sub-clause into a list of thirty puts THIRTY uncollapsed rows in the change
 * list and produces zero cascades — measured, not supposed. The shift comes
 * from the counter path because `(d)` → `(e)` cannot be subtracted, and the
 * level it shifted AT is folded into the token kind so two levels cascading at
 * once stay two groups.
 *
 * That level is the one whose counter moved, NOT the paragraph's own: `1.1` →
 * `2.1` is a level-1 paragraph carried by a level-0 insertion, and grouping it
 * under its own level would scatter one insertion's cascade across every level
 * it touched. See `PropertyChange.cascade.shiftedIlvl`.
 */
function derivePropertyMember(op: DiffOp, config: CompareConfig): Omit<CascadeMember, 'op'> | null {
  const cascade = op.propertyChange?.cascade;
  if (!cascade || cascade.shift === 0) return null;
  // The unit noun is the jurisdiction's, and the jurisdiction's noun lives in
  // Config (§7.4, CONVENTIONS § Terminology) — a firm drafting statutes may set
  // it to something else. `Clause` is the SG/E&W default, not a literal.
  //
  // `primaryUnit` and NOT `referenceTokens[0]`: that list is a DETECTION set —
  // every noun whose cross-references should be recognised in prose — and its
  // order carries no meaning. Reading position 0 silently renamed this heading
  // to "paragraph numbering" for a profile that listed the nouns another way.
  return {
    tokenKind: `${config.primaryUnit} numbering (level ${cascade.shiftedIlvl + 1})`,
    shift: cascade.shift,
    example: `${cascade.from} → ${cascade.to}`,
  };
}

function deriveMember(op: DiffOp, config: CompareConfig): Omit<CascadeMember, 'op'> | null {
  if (op.type === 'property') return derivePropertyMember(op, config);

  const pair = refPair(op);
  if (!pair) return null;
  const oldRef = parseRef(pair.del.text);
  const newRef = parseRef(pair.ins.text);
  if (oldRef.trailingNumber === null || newRef.trailingNumber === null) return null;

  // tokenKind: prefer an explicit leading word; then the noun sitting unchanged
  // in front of the mark; then the other side.
  const tokenKind =
    oldRef.tokenKind ?? refKindBefore(op.wordDiffs ?? [], pair.delAt, config) ?? newRef.tokenKind;
  if (tokenKind === null) return null;

  const shift = newRef.trailingNumber - oldRef.trailingNumber;
  const example = formatExample(tokenKind, oldRef, newRef);
  return { tokenKind, shift, example };
}

/** "Clause 4.2 → 4.3" style example string. */
function formatExample(tokenKind: string, oldRef: ParsedRef, newRef: ParsedRef): string {
  const left = `${tokenKind} ${oldRef.refBody}`.trim();
  return `${left} → ${newRef.refBody}`.trim();
}

/** Stable group key for a (tokenKind, shift) pair. */
function groupKey(tokenKind: string, shift: number): string {
  return `${tokenKind}\u0000${shift}`;
}

/**
 * Group crossRefRenumber ops by (tokenKind, shift); for each group at or above
 * `config.renumberThreshold`, emit one CascadeAnnotation and tag each member op
 * with `collapsedInto`.
 *
 * Returns a NEW ops array of the SAME LENGTH as the input (ADR-0008: never
 * drops an op) plus the list of cascade annotations.
 */
export function computeCascades(
  ops: DiffOp[],
  config: CompareConfig,
): { ops: DiffOp[]; cascades: CascadeAnnotation[] } {
  // 1. Derive cascade metadata for each crossRefRenumber op, keyed by op id.
  const memberByOpId = new Map<string, Omit<CascadeMember, 'op'>>();
  const groups = new Map<
    string,
    {
      tokenKind: string;
      shift: number;
      opIds: string[];
      examples: string[];
      source: 'crossReference' | 'numbering';
    }
  >();

  for (const op of ops) {
    if (op.classification !== 'crossRefRenumber') continue;
    const derived = deriveMember(op, config);
    if (!derived) continue;
    memberByOpId.set(op.id, derived);
    const key = groupKey(derived.tokenKind, derived.shift);
    let group = groups.get(key);
    if (!group) {
      group = {
        tokenKind: derived.tokenKind,
        shift: derived.shift,
        opIds: [],
        examples: [],
        source: op.type === 'property' ? 'numbering' : 'crossReference',
      };
      groups.set(key, group);
    }
    group.opIds.push(op.id);
    if (group.examples.length < 3) group.examples.push(derived.example);
  }

  // 2. Emit a CascadeAnnotation for each group that meets the threshold; record
  //    which op ids are collapsed and into which cascade id.
  const cascades: CascadeAnnotation[] = [];
  const collapsedOpToCascade = new Map<string, string>();
  let cascadeIndex = 0;

  // Iterate groups in first-appearance order for deterministic cascade ids.
  for (const group of groups.values()) {
    if (group.opIds.length < config.renumberThreshold) continue;
    const id = `cascade-${cascadeIndex}`;
    cascadeIndex += 1;
    cascades.push({
      id,
      tokenKind: group.tokenKind,
      shift: group.shift,
      affectedCount: group.opIds.length,
      exampleRefs: group.examples.slice(0, 3),
      source: group.source,
    });
    for (const opId of group.opIds) {
      collapsedOpToCascade.set(opId, id);
    }
  }

  // 3. Map over the FULL op list — every op is preserved. Members of a collapsed
  //    group get `collapsedInto`; all others pass through unchanged.
  const resultOps: DiffOp[] = ops.map((op) => {
    const cascadeId = collapsedOpToCascade.get(op.id);
    if (cascadeId === undefined) {
      return { ...op };
    }
    return { ...op, collapsedInto: cascadeId };
  });

  return { ops: resultOps, cascades };
}
