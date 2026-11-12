import { useState } from 'react';
import { AlertTriangle, Info, ListTree, Loader2, XCircle } from 'lucide-react';
import type { CompareOutcome } from '../compare/protocol.js';
import type { RedlineTab } from '../tabs/model.js';
import { ChangeSummaryBar } from './ChangeSummaryBar.js';
import { RedlineDocument } from './RedlineDocument.js';
import { ExportForAiButton } from './ExportForAiButton.js';
import { SavePdfButton } from './SavePdfButton.js';
import { StructurePanel } from './StructurePanel.js';

export function RedlineView({ tab, onCancel }: { tab: RedlineTab; onCancel: () => void }) {
  const [showStructure, setShowStructure] = useState(false);
  switch (tab.status.state) {
    case 'running':
      return (
        <Centred>
          <Loader2 size={22} className="animate-spin text-accent" aria-hidden />
          <p className="text-[13px] text-text">{tab.status.step}</p>
          <div
            className="h-1 w-56 overflow-hidden rounded-full bg-border"
            role="progressbar"
            aria-valuenow={tab.status.progress}
            aria-valuemin={0}
            aria-valuemax={100}
          >
            <div
              className="h-full bg-accent transition-[width] duration-200"
              style={{ width: `${tab.status.progress}%` }}
            />
          </div>
          <p className="text-[11px] text-text-faint">
            {tab.originalName} → {tab.revisedName}
          </p>
          <button
            type="button"
            onClick={onCancel}
            className="mt-1 rounded-md border border-border px-2.5 py-1 text-[11px] text-text-muted transition-colors hover:border-border-strong hover:text-text"
          >
            Cancel
          </button>
        </Centred>
      );

    case 'failed':
      return (
        <Centred>
          <AlertTriangle size={22} className="text-danger" aria-hidden />
          <p className="text-[13px] font-medium text-text">This comparison failed</p>
          <p className="max-w-prose text-center text-[12px] leading-relaxed text-text-muted">
            {tab.status.message}
          </p>
          <p className="text-[11px] text-text-faint">
            {tab.originalName} → {tab.revisedName}
          </p>
        </Centred>
      );

    case 'cancelled':
      return (
        <Centred>
          <XCircle size={22} className="text-text-faint" aria-hidden />
          <p className="text-[13px] text-text-muted">Comparison cancelled.</p>
        </Centred>
      );

    case 'ready': {
      const outcome = tab.result as CompareOutcome | null;
      if (!outcome) {
        return (
          <Centred>
            <p className="text-[12px] text-text-muted">No result was produced.</p>
          </Centred>
        );
      }
      return (
        <div className="flex h-full flex-col">
          <div className="flex shrink-0 items-center gap-2 border-b border-border px-4 py-1.5">
            <p className="min-w-0 truncate text-[11px] text-text-muted">
              {tab.originalName} → {tab.revisedName}
            </p>
            <div className="ml-auto flex items-center gap-2">
              {outcome.structure ? (
                <button
                  type="button"
                  aria-pressed={showStructure}
                  onClick={() => setShowStructure((v) => !v)}
                  title="Show each document's clause structure beside the redline"
                  className={
                    'flex shrink-0 items-center gap-1.5 rounded border px-2.5 py-1 text-[11px] font-medium transition-colors ' +
                    (showStructure
                      ? 'border-accent text-accent'
                      : 'border-border text-text hover:border-border-strong')
                  }
                >
                  <ListTree size={12} aria-hidden />
                  Structure
                </button>
              ) : null}
              <ExportForAiButton
                context={outcome.context}
                baseName={tab.revisedName.replace(/\.[^.]+$/, '')}
              />
              <SavePdfButton
                model={outcome.model}
                title={`${tab.originalName} → ${tab.revisedName}`}
                author="Mutatis Desktop"
                fileName={`${tab.revisedName.replace(/\.[^.]+$/, '')} (redline).pdf`}
              />
            </div>
          </div>
          <ChangeSummaryBar
            summary={outcome.summary}
            quarantinedOriginal={outcome.quarantinedOriginal}
            quarantinedRevised={outcome.quarantinedRevised}
          />
          {tab.notes && tab.notes.length > 0 ? (
            <div
              role="note"
              className="flex shrink-0 gap-2 border-b border-border bg-surface-sunken px-4 py-2 text-[11px] leading-relaxed text-text-muted"
            >
              <Info size={13} className="mt-0.5 shrink-0 text-[var(--redline-move)]" aria-hidden />
              <ul className="flex min-w-0 flex-col gap-1">
                {tab.notes.map((note) => (
                  <li key={note}>{note}</li>
                ))}
              </ul>
            </div>
          ) : null}
          <div className="flex min-h-0 flex-1">
            <div className="min-w-0 flex-1 overflow-y-auto bg-surface">
              <RedlineDocument model={outcome.model} />
            </div>
            {showStructure && outcome.structure ? (
              <StructurePanel
                structure={outcome.structure}
                originalName={tab.originalName}
                revisedName={tab.revisedName}
                onClose={() => setShowStructure(false)}
              />
            ) : null}
          </div>
        </div>
      );
    }
  }
}

function Centred({ children }: { children: React.ReactNode }) {
  return (
    <div className="flex h-full flex-col items-center justify-center gap-3 px-8">{children}</div>
  );
}
