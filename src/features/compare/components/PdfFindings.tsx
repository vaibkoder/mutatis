/**
 * PDF findings panel body.
 *
 * Renders a completed PDF-vs-PDF comparison: the change SUMMARY, the navigable
 * change LIST (reused from the Word path), and — because a PDF has no openable
 * redline document — an in-pane {@link SideBySideDiff} for the selected change.
 *
 * ADR-0008: the summary total is always the full count; the list filters affect
 * DISPLAY only; low-confidence changes (fuzzy-anchored paragraphs, approximate
 * table cells) are BADGED, never hidden. Pure presentational — the driver owns
 * the ops, summary, ParaRecords and filter signals; this only renders them.
 */
import { useState } from 'preact/hooks';
import type { ChangeSummary, DiffOp, ParaRecord } from '../model.js';
import { SummaryPanel } from './SummaryPanel.js';
import { ChangeList, cascadeLabelMap, type FilterFlags } from './ChangeList.js';
import { SideBySideDiff } from './SideBySideDiff.js';

interface PdfFindingsProps {
  ops: DiffOp[];
  summary: ChangeSummary;
  filters: FilterFlags;
  /** Original-side ParaRecords, to resolve plain text for pure delete/move rows. */
  originalParas: ParaRecord[];
  /** Revised-side ParaRecords, to resolve plain text for pure insert/move rows. */
  revisedParas: ParaRecord[];
}

export function PdfFindings({
  ops,
  summary,
  filters,
  originalParas,
  revisedParas,
}: PdfFindingsProps) {
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const selected = ops.find(o => o.id === selectedId) ?? null;

  return (
    <div class="dp-compare__pdf-findings">
      <SummaryPanel summary={summary} />
      <ChangeList
        ops={ops}
        filters={filters}
        cascadeLabels={cascadeLabelMap(summary.renumberingCascades)}
        onGoToChange={id => setSelectedId(id)}
      />
      {selected && (
        <SideBySideDiff
          op={selected}
          originalText={resolveText(selected, 'original', originalParas)}
          revisedText={resolveText(selected, 'revised', revisedParas)}
        />
      )}
    </div>
  );
}

/** Resolve the plain text of an op's paragraph on one side (for non-replace ops). */
function resolveText(
  op: DiffOp,
  side: 'original' | 'revised',
  paras: ParaRecord[],
): string | undefined {
  const range = side === 'original' ? op.origRange : op.revRange;
  if (range === null) return undefined;
  return paras[range.start]?.normText;
}
