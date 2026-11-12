/**
 * Change-set serialiser (feature 11 — spec §4.4, rule F1: before/after, never markup).
 *
 * Projects Compare's `DiffOp[]` into plain `{original, revised}` string pairs by
 * traversing `wordDiffs` (del+eq = original, ins+eq = revised). The consumer is
 * an LLM, not a diff-viewer: no strikethrough, no `<w:del>`, no caret notation.
 */
import type { DiffOp, WordDiff } from '@features/compare/model.js';
import type { ChangeSetExport, ChangeEntry } from '../types.js';

export interface ChangeSetDeps {
  baseline: string;
  revised: string;
  /** Best-effort human citation for a change (part label + clause number). */
  clauseFor: (op: DiffOp) => string;
  /** Whole-paragraph original text, for pure deletes/moves lacking wordDiffs. */
  origParaText?: (op: DiffOp) => string;
  /** Whole-paragraph revised text, for pure inserts/moves lacking wordDiffs. */
  revParaText?: (op: DiffOp) => string;
  /** Target citation for a move op (resolves moveTargetId). */
  moveTarget?: (op: DiffOp) => string;
}

/**
 * The op types this export carries. Rule F1 is before/after TEXT, so the two
 * types with no text delta are out: 'equal', and 'property' (ADR-0042 — a
 * matched pair whose style or numbering changed but whose words did not).
 * Emitting a property change here would mean an entry whose `original` and
 * `revised` are the same string, which tells an LLM nothing and reads as a
 * mistake. Representing it properly needs a field this export's schema does not
 * have (`ChangeEntry.type` is a published contract), so it is feature 11's to
 * add, not Compare's to smuggle in. The change itself is not lost: it is in
 * `ops`, in the change summary's formatting count, and in the exported report.
 */
const EXPORTED_TYPES = new Set<DiffOp['type']>(['insert', 'delete', 'replace', 'move']);

export function serialiseChangeSet(ops: readonly DiffOp[], deps: ChangeSetDeps): ChangeSetExport {
  const changes: ChangeEntry[] = [];
  for (const op of ops) {
    if (!EXPORTED_TYPES.has(op.type)) continue; // skip 'equal' / 'property'
    if (op.collapsedInto) continue; // already represented by its collapse target
    changes.push(toChangeEntry(op, deps));
  }
  return { baseline: deps.baseline, revised: deps.revised, changes };
}

function toChangeEntry(op: DiffOp, deps: ChangeSetDeps): ChangeEntry {
  const entry: ChangeEntry = {
    id: op.id,
    type: op.type as ChangeEntry['type'],
    classification: op.classification,
    clause: deps.clauseFor(op),
  };

  const orig = origText(op, deps);
  const rev = revText(op, deps);

  switch (op.type) {
    case 'insert':
      entry.revised = rev;
      break;
    case 'delete':
      entry.original = orig;
      break;
    case 'replace':
      entry.original = orig;
      entry.revised = rev;
      break;
    case 'move':
      entry.original = orig;
      entry.revised = rev;
      if (deps.moveTarget) entry.move_target = deps.moveTarget(op);
      break;
  }

  if (op.lowConfidence) entry.low_confidence = true;
  return entry;
}

function joinSide(wordDiffs: WordDiff[] | undefined, keep: ReadonlyArray<WordDiff['op']>): string | null {
  if (!wordDiffs) return null;
  return wordDiffs.filter(d => keep.includes(d.op)).map(d => d.text).join('');
}

function origText(op: DiffOp, deps: ChangeSetDeps): string {
  const fromDiffs = joinSide(op.wordDiffs, ['del', 'eq']);
  if (fromDiffs !== null) return fromDiffs;
  return deps.origParaText?.(op) ?? '';
}

function revText(op: DiffOp, deps: ChangeSetDeps): string {
  const fromDiffs = joinSide(op.wordDiffs, ['ins', 'eq']);
  if (fromDiffs !== null) return fromDiffs;
  return deps.revParaText?.(op) ?? '';
}
