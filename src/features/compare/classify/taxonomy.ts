/**
 * Feature 08 — Compare · classifier taxonomy (Task B1).
 *
 * Priority-ordered, rule-based classification of a single DiffOp into the §6
 * taxonomy. A change MAY match several categories (`DiffOp.classes`); the
 * highest-priority match is the primary (`DiffOp.classification`).
 *
 * INVARIANT (ADR-0008): classification only LABELS a change. Nothing here
 * removes a DiffOp, mutates its `type`/`wordDiffs`, or suppresses a change.
 * No ML, no semantic-equivalence, no fuzzy matching — regex + OOXML inspection
 * only.
 *
 * Spec: spec/features/08-compare.md §6.
 */

import type { ChangeClass, DiffOp, ParaRecord, WordDiff } from '../model.js';

/**
 * The §6 taxonomy in priority order (index 0 = highest priority). The primary
 * classification of an op is the first entry in this list that it matches.
 */
export const PRIORITY_ORDER: readonly ChangeClass[] = [
  'structural', // 1
  'obligationLanguage', // 2
  'numericalValue', // 3
  'definedTerm', // 4
  'crossRefRenumber', // 5
  'substantiveText', // 6
  'whitespacePunct', // 7
  'formatting', // 8
] as const;

/**
 * Modal obligation keywords (§6, priority 2). Whole-word, case-insensitive.
 * Longer phrases first so "shall not" is preferred over "shall" when building
 * the alternation, though detection is independent per keyword.
 */
const OBLIGATION_MODALS: readonly string[] = [
  'shall not',
  'may not',
  'shall',
  'must',
  'may',
  'will',
  'should',
];

/** Atomic kinds that signal a numerical-value change (§6, priority 3). */
const NUMERIC_ATOMICS = new Set<string>(['amount', 'date', 'percent', 'duration']);

/** Title-Case multi-word phrase (two or more Title-Case words) — a defined-term signal. */
const TITLE_CASE_MULTIWORD = /\b[A-Z][a-z]+(?:\s+[A-Z][a-z]+)+\b/;

/**
 * A paragraph that DEFINES a term, rather than merely using one: the quoted
 * term followed by a definitional verb. Deliberately narrow — the broad
 * unquoted-candidate recognition of ADR-0013 belongs to feature 01, and
 * borrowing it here would re-introduce the over-matching this replaced.
 */
// NB the curly-quote limb is unreachable in practice: `normaliseText`
// straightens curly double quotes before any classifier sees the text
// (`preprocess/normalise.ts`). Kept so the pattern stays correct if that
// changes, but it is not a tested path — see ADR-0033.
const DEFINITION_PARA =
  /(?:"[^"]+"|\u201C[^\u201D]+\u201D)\s+(?:means\b|shall mean\b|has the meaning\b|shall have the meaning\b)/i;

/** Whole-word modal match, case-insensitive. */
function containsModal(text: string): boolean {
  for (const modal of OBLIGATION_MODALS) {
    const re = new RegExp(`\\b${modal.replace(/\s+/g, '\\s+')}\\b`, 'i');
    if (re.test(text)) return true;
  }
  return false;
}

/** The WordDiffs of an op that represent an actual change (ins or del). */
function changedDiffs(op: DiffOp): WordDiff[] {
  return (op.wordDiffs ?? []).filter((wd) => wd.op !== 'eq');
}

/** Whether the op has any non-equal (ins/del) word-text delta with real text. */
function hasTextDelta(op: DiffOp): boolean {
  return changedDiffs(op).some((wd) => wd.text.length > 0);
}

/** Resolve the ParaRecords this op touches, from both sides. */
function opParas(op: DiffOp, original: ParaRecord[], revised: ParaRecord[]): ParaRecord[] {
  const out: ParaRecord[] = [];
  if (op.origRange) {
    for (let i = op.origRange.start; i <= op.origRange.end; i++) {
      const p = original[i];
      if (p) out.push(p);
    }
  }
  if (op.revRange) {
    for (let i = op.revRange.start; i <= op.revRange.end; i++) {
      const p = revised[i];
      if (p) out.push(p);
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Per-category detectors. Each returns true if the op matches that category.
// ---------------------------------------------------------------------------

function isStructural(op: DiffOp): boolean {
  return op.type === 'insert' || op.type === 'delete' || op.type === 'move';
}

function isObligationLanguage(op: DiffOp): boolean {
  // A modal keyword is added or removed in one of the changed runs.
  return changedDiffs(op).some((wd) => containsModal(wd.text));
}

function isNumericalValue(op: DiffOp, paras: ParaRecord[]): boolean {
  // A changed atomic value, OR any value change inside a table cell.
  const atomicHit = changedDiffs(op).some(
    (wd) => wd.atomic !== undefined && NUMERIC_ATOMICS.has(wd.atomic),
  );
  if (atomicHit) return true;
  const inCell = paras.some((p) => p.isTableCell);
  return inCell && hasTextDelta(op);
}

/**
 * §6 row 4 — "Change **touches** a quoted term or a Title-Case multi-word term,
 * **or alters the definition paragraph** of one." Both limbs are about the
 * change or the kind of paragraph, never about whether some unrelated term is
 * mentioned nearby.
 *
 * The distinction is load-bearing, not pedantic. `TITLE_CASE_MULTIWORD` matches
 * a sentence-initial capital followed by a capitalised word, so `The Vendor
 * shall …` — how most clauses in an SG/E&W contract open — matches. Testing the
 * PARAGRAPH therefore made almost every edit a defined-term change and, because
 * row 4 outranks row 5, meant a pure renumbering never classified as
 * `crossRefRenumber` and so never joined a renumbering cascade (§6 "Collapsed").
 * Identical edits classified differently depending on prose they never touched.
 */
function isDefinedTerm(op: DiffOp, paras: ParaRecord[]): boolean {
  const changed = changedDiffs(op);
  // Limb 1a: a changed WordDiff is a quoted defined-term atomic …
  if (changed.some((wd) => wd.atomic === 'definedTerm')) return true;
  // … 1b: or the changed TEXT itself carries a Title-Case multi-word term.
  if (changed.some((wd) => TITLE_CASE_MULTIWORD.test(wd.text))) return true;
  // Limb 2: or the touched paragraph is the one that DEFINES a term, where any
  // edit at all is an edit to the definition.
  return paras.some((p) => DEFINITION_PARA.test(p.normText));
}

function isCrossRefRenumber(op: DiffOp): boolean {
  // ALL changed WordDiffs are clauseRef tokens (and there is at least one).
  const changed = changedDiffs(op);
  if (changed.length === 0) return false;
  return changed.every((wd) => wd.atomic === 'clauseRef');
}

function isSubstantiveText(op: DiffOp): boolean {
  // Any remaining run-text insert/delete/replace.
  return hasTextDelta(op);
}

/**
 * Whitespace/punctuation (§6 priority 7): there is a real text delta, but under
 * aggressive normalisation (lowercased, all punctuation stripped) the two
 * matched sides are identical — only spacing/punctuation differs.
 */
function isWhitespacePunct(
  op: DiffOp,
  original: ParaRecord[],
  revised: ParaRecord[],
): boolean {
  if (!op.origRange || !op.revRange) return false;
  if (!hasTextDelta(op)) return false;
  const origAgg = aggressiveForRange(op.origRange, original);
  const revAgg = aggressiveForRange(op.revRange, revised);
  if (origAgg === null || revAgg === null) return false;
  return origAgg === revAgg;
}

function isFormatting(op: DiffOp): boolean {
  // Pure formatting: zero text delta. The diff still exists (the op was
  // emitted) but it is style-only (<w:rPr>/<w:pPr>).
  return !hasTextDelta(op);
}

/** Concatenate aggressively-normalised text across a record range, or null. */
function aggressiveForRange(
  range: { start: number; end: number },
  records: ParaRecord[],
): string | null {
  const parts: string[] = [];
  for (let i = range.start; i <= range.end; i++) {
    const p = records[i];
    if (!p) return null;
    parts.push(p.normTextAggressive);
  }
  return parts.join(' ');
}

// ---------------------------------------------------------------------------
// Public: collect all matching classes for an op, in priority order.
// ---------------------------------------------------------------------------

/**
 * Return every taxonomy class the op matches, ordered by priority (highest
 * first). The first entry is the primary classification. Always returns at
 * least one class (never empty — never drops the op).
 */
export function classifyClasses(
  op: DiffOp,
  original: ParaRecord[],
  revised: ParaRecord[],
): ChangeClass[] {
  // ADR-0042 — a `property` op has identical text on both sides, so every
  // detector below (all of which read the word diff or the text) would find
  // nothing and fall through to the empty-set default. It carries its own
  // evidence instead.
  if (op.type === 'property') return classifyProperty(op);

  const paras = opParas(op, original, revised);
  const matched = new Set<ChangeClass>();

  if (isStructural(op)) matched.add('structural');
  if (isObligationLanguage(op)) matched.add('obligationLanguage');
  if (isNumericalValue(op, paras)) matched.add('numericalValue');
  if (isDefinedTerm(op, paras)) matched.add('definedTerm');
  if (isCrossRefRenumber(op)) matched.add('crossRefRenumber');
  if (isSubstantiveText(op)) matched.add('substantiveText');
  if (isWhitespacePunct(op, original, revised)) matched.add('whitespacePunct');
  if (isFormatting(op)) matched.add('formatting');

  if (matched.size === 0) {
    // A diff op that matched nothing above is, by §6, formatting-only (a
    // style-only delta). Never drop it — always carry a class.
    matched.add('formatting');
  }

  return PRIORITY_ORDER.filter((c) => matched.has(c));
}

/**
 * §6 mapping for a non-text property change (ADR-0042). Two answers, both
 * already written in the spec, and the split between them is what keeps a long
 * document readable:
 *
 * - The paragraph itself was restyled or re-levelled — `w:pStyle` changed, or
 *   its numbering IDENTITY did (level + level-definition shape — never
 *   `(numId, ilvl)`, which is document-local). §6 row 8 lists *"style id"*
 *   inside formatting-only
 *   and §9.9 says per-paragraph `styleId` changes are likewise formatting-only.
 *   → `formatting`.
 * - Only the RENDERED label moved, identity untouched: this paragraph did not
 *   change at all; its number shifted because a sibling did. §9.4 calls that a
 *   `crossRefRenumber`/formatting signal. → `crossRefRenumber`, which is the
 *   class §6.1's cascade collapse keys on, so inserting one clause cannot bury
 *   the real edit under one op per clause that followed it.
 * - A zero-text object was added, removed or moved (ADR-0045 § Amendment) →
 *   `substantiveText`, alongside either of the above; never a cascade.
 *
 * `formatting` rides along in `classes` for the cascade case too: the change is
 * still formatting in nature, and `classes` is every applicable label. The
 * signature/execution-block rule (§7.2) strips `formatting` and forces
 * substantive afterwards, which is §9.9's *"except where they alter a
 * signature/structural paragraph"* carve-out arriving for free.
 */
function classifyProperty(op: DiffOp): ChangeClass[] {
  const change = op.propertyChange;
  const matched = new Set<ChangeClass>();
  // ADR-0045 § Amendment — a footnote or endnote reference, picture, text box
  // or embedded object added, removed or moved. That is a run inserted or
  // deleted (§6 row 6), and §9.8 already makes a footnote change substantive.
  // Never `formatting`, which *Hide formatting-only* hides by default.
  if (change?.objects === true) matched.add('substantiveText');
  const identityChanged = change === undefined || change.styleId || change.numberingIdentity;
  const numberOnly = change !== undefined && !identityChanged && change.numberLabel;
  if (identityChanged || numberOnly) matched.add('formatting');
  if (numberOnly && change.objects !== true) matched.add('crossRefRenumber');
  if (matched.size === 0) matched.add('formatting'); // never empty — never drops the op
  return PRIORITY_ORDER.filter((c) => matched.has(c));
}
