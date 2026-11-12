/**
 * Compare → AI change-set export (feature 11 / ADR-0021), shared by the
 * Word-vs-Word and Text-vs-text panels.
 *
 * Builds the "Change set with before/after" envelope (export #3) from the active
 * DiffOps + the compared ParaRecord arrays. `wordDiffs` are populated only for
 * `replace` ops, so pure delete/insert/move carry their before/after text as
 * whole paragraphs sliced from the compared records (origRange→Original,
 * revRange→Revised) — without that those entries would serialise with empty
 * before/after, defeating the feature's before/after guarantee (rule F1).
 *
 * A header, footer or note op (ADR-0047) indexes its OWN part's records, so
 * `parts` (the driver's `parts` signal, keyed by `partId`) resolves its text
 * and gives its locator ("Footer", "Footnote 3 (added)"); an op not in it is in
 * the main text.
 *
 * Pure serialisation over already-computed data — zero network (ADR-0001/0021),
 * never markup, never suppresses (ADR-0008). Every op with a TEXT delta is included; 'equal' and 'property' (ADR-0042) are excluded because rule F1 is before/after text and they have none — see `serialise/change-set.ts`. The canonical
 * home for this builder; the Text-vs-text panel consumes it directly. The
 * Word-vs-Word panel does too.
 *
 * Spec: spec/features/11-ai-export.md §4.4 (export #3).
 */
import type { AiExportEnvelope } from '@features/ai-export/index.js';
import { buildExport, partLabel } from '@features/ai-export/index.js';
import { getStoredJurisdiction } from '@shared/config/index.js';

import type { DiffOp, ParaRecord } from './model.js';
import type { PartView } from './ui-types.js';

/** Whole-paragraph text for an op's (inclusive) index range, joined across paras. */
export function rangeText(paras: ParaRecord[], range: { start: number; end: number } | null): string {
  if (!range) return '';
  const out: string[] = [];
  for (let i = range.start; i <= range.end; i++) {
    const rec = paras[i];
    if (rec) out.push(rec.normText);
  }
  return out.join('\n');
}

/**
 * Build the change-set export envelope from the active diff ops, or `null` when
 * there is nothing to export (no ops). Shared by both Compare panels.
 */
export function buildChangeSet(
  ops: DiffOp[],
  originalParas: ParaRecord[],
  revisedParas: ParaRecord[],
  parts: ReadonlyMap<string, PartView> = new Map(),
): AiExportEnvelope | null {
  if (ops.length === 0) return null;
  const byId = new Map(ops.map((op) => [op.id, op]));
  const records = (op: DiffOp) => parts.get(op.partId) ?? { originalParas, revisedParas };
  const where = (op: DiffOp) => parts.get(op.partId)?.location.label ?? partLabel(op.partId);
  return buildExport(
    'compare',
    {
      ops,
      baseline: 'Original',
      revised: 'Revised',
      origParaText: (op) => rangeText(records(op).originalParas, op.origRange),
      revParaText: (op) => rangeText(records(op).revisedParas, op.revRange),
      moveTarget: (op) => {
        const pair = op.moveTargetId ? byId.get(op.moveTargetId) : undefined;
        return pair ? where(pair) : '';
      },
      clauseFor: where,
    },
    {
      generated: new Date().toISOString(),
      document: 'Compared documents',
      jurisdiction: getStoredJurisdiction() === 'SG' ? 'SG' : 'EW',
      index: null,
      orderedParaIds: [],
      paraText: () => '',
    },
  );
}
