import { AlertTriangle, Bot, Download, Loader2, XCircle } from 'lucide-react';
import { RedlineDocument } from '../redline/RedlineDocument.js';
import { ChangeSummaryBar } from '../redline/ChangeSummaryBar.js';
import { SavePdfButton } from '../redline/SavePdfButton.js';
import type { SelfRedlineOutcome } from '../compare/protocol.js';
import type { SelfRedlineTab } from '../tabs/model.js';

/**
 * The assistant's own changes, against the document it edited.
 *
 * This exists because of one rule: every AI edit lands as a tracked change the
 * user reviews. The edits below have already been applied to a COPY — nothing on
 * disk has moved — and what is shown is the comparison engine's redline of that
 * copy against the original. The same engine, the same renderer, the same
 * markup. An AI edit is not a special kind of change that gets its own,
 * friendlier presentation.
 *
 * Refused edits are shown with equal weight. A model that quoted a clause
 * slightly wrong produced nothing, and the user needs to know that as much as
 * they need to see what did change.
 */
export function SelfRedlineView({
  tab,
  onSave,
}: {
  tab: SelfRedlineTab;
  onSave: (outcome: SelfRedlineOutcome, targetName: string) => void;
}) {
  if (tab.status.state === 'running') {
    return (
      <Centred>
        <Loader2 size={20} className="animate-spin text-accent" aria-hidden />
        <p className="text-[12px] text-text-muted">{tab.status.step}</p>
      </Centred>
    );
  }
  if (tab.status.state === 'failed') {
    return (
      <Centred>
        <AlertTriangle size={22} className="text-danger" aria-hidden />
        <p className="text-[13px] font-medium text-text">No edit was applied</p>
        <p className="max-w-prose text-center text-[12px] text-text-muted">{tab.status.message}</p>
      </Centred>
    );
  }
  if (tab.status.state === 'cancelled' || !tab.result) {
    return (
      <Centred>
        <XCircle size={22} className="text-text-faint" aria-hidden />
        <p className="text-[12px] text-text-muted">This edit was cancelled.</p>
      </Centred>
    );
  }

  const outcome = tab.result as SelfRedlineOutcome;

  return (
    <div className="flex h-full min-h-0 flex-col">
      <div className="flex shrink-0 items-center gap-2 border-b border-border bg-surface-sunken px-4 py-2">
        <Bot size={14} className="text-accent" aria-hidden />
        <div className="min-w-0">
          <p className="truncate text-[12px] font-medium text-text">
            The assistant&rsquo;s changes to {tab.targetName}
          </p>
          <p className="text-[11px] text-text-muted">
            {outcome.applied.length} applied
            {outcome.refused.length > 0 ? `, ${outcome.refused.length} refused` : ''} — review them
            below, then save the edited file if you want it.
          </p>
        </div>
        <div className="ml-auto">
          <SavePdfButton
            model={outcome.outcome.model}
            title={`Assistant's changes to ${tab.targetName}`}
            author="Mutatis assistant"
            fileName={`${tab.targetName.replace(/\.docx$/i, '')} (assistant redline).pdf`}
          />
        </div>
        <button
          type="button"
          onClick={() => onSave(outcome, tab.targetName)}
          className="flex shrink-0 items-center gap-1.5 rounded bg-accent px-2.5 py-1.5 text-[11px] font-medium text-accent-text transition-opacity hover:opacity-90"
        >
          <Download size={12} aria-hidden />
          Save edited .docx
        </button>
      </div>


      {outcome.refused.length > 0 && (
        <div className="shrink-0 border-b border-border bg-danger/10 px-4 py-2">
          <p className="mb-1 flex items-center gap-1.5 text-[11px] font-medium text-danger">
            <AlertTriangle size={12} aria-hidden />
            {outcome.refused.length} edit{outcome.refused.length === 1 ? ' was' : 's were'} refused
            and {outcome.refused.length === 1 ? 'is' : 'are'} NOT in the document
          </p>
          <ul className="flex flex-col gap-1">
            {outcome.refused.map((edit, index) => (
              <li key={`${edit.find}-${index}`} className="text-[11px] leading-relaxed text-text">
                <span className="font-mono text-text-muted">&ldquo;{truncate(edit.find)}&rdquo;</span>
                {' — '}
                {edit.refusal}
              </li>
            ))}
          </ul>
        </div>
      )}

      {outcome.applied.some((edit) => edit.reason !== undefined) && (
        <div className="shrink-0 border-b border-border px-4 py-2">
          <p className="mb-1 text-[11px] font-medium text-text-muted">Why, in the assistant&rsquo;s words</p>
          <ul className="flex flex-col gap-1">
            {outcome.applied.map((edit, index) =>
              edit.reason !== undefined ? (
                <li key={`${edit.find}-${index}`} className="text-[11px] leading-relaxed text-text-muted">
                  <span className="font-mono">&ldquo;{truncate(edit.find)}&rdquo;</span> — {edit.reason}
                </li>
              ) : null,
            )}
          </ul>
        </div>
      )}

      <ChangeSummaryBar
        summary={outcome.outcome.summary}
        quarantinedOriginal={outcome.outcome.quarantinedOriginal}
        quarantinedRevised={outcome.outcome.quarantinedRevised}
      />

      <div className="base-ui-disable-scrollbar min-h-0 flex-1 overflow-auto">
        <RedlineDocument model={outcome.outcome.model} />
      </div>
    </div>
  );
}

function truncate(value: string): string {
  return value.length > 70 ? `${value.slice(0, 67)}…` : value;
}

function Centred({ children }: { children: React.ReactNode }) {
  return (
    <div className="flex h-full flex-col items-center justify-center gap-3 px-8">{children}</div>
  );
}
