import { AlertTriangle, Loader2 } from 'lucide-react';
import { useState } from 'react';
import { LandingPage, type LandingMode } from '../landing/LandingPage.js';
import { emptyPasteSide } from '../paste/pasted.js';
import type { PasteState } from '../paste/PasteCompare.js';
import { PairingGrid } from '../pairing/PairingGrid.js';
import { RedlineView } from '../redline/RedlineView.js';
import { DeviationMatrixView } from '../ai/DeviationMatrixView.js';
import { SelfRedlineView } from '../edits/SelfRedlineView.js';
import { saveEditedDocument } from '../edits/save.js';
import type { DeviationMatrix } from '../ai/deviation-matrix.js';
import type { PairingResult } from '../pairing/build-pairing.js';
import { TabStrip } from '../tabs/TabStrip.js';
import { activeTab, type TabId, type TabState } from '../tabs/model.js';
import type { DocSide, DocumentSet, LoadedDoc } from '../state/documents.js';

interface WorkPaneProps {
  documents: DocumentSet;
  tabs: TabState;
  onAdd: (side: DocSide, docs: readonly LoadedDoc[]) => void;
  onRemove: (side: DocSide, id: string) => void;
  onCompare: () => void;
  onSelectTab: (id: TabId) => void;
  onCloseTab: (id: TabId) => void;
  onCancelTab: (id: TabId) => void;
  onRepoint: (tabId: TabId, rowId: string, b: string | null) => void;
  onConfirmPairing: (tabId: TabId) => void;
  /** Run a one-to-one comparison of two already-built documents (pasted text). */
  onComparePair: (original: LoadedDoc, modified: LoadedDoc) => void;
}

export function WorkPane({
  documents,
  tabs,
  onAdd,
  onRemove,
  onCompare,
  onSelectTab,
  onCloseTab,
  onCancelTab,
  onRepoint,
  onConfirmPairing,
  onComparePair,
}: WorkPaneProps) {
  const current = activeTab(tabs);
  // Held here, not in LandingPage: that unmounts whenever a redline tab is in
  // front, and pasted text must survive a look at the result and a return.
  const [mode, setMode] = useState<LandingMode>('files');
  const [paste, setPaste] = useState<PasteState>({ original: emptyPasteSide, modified: emptyPasteSide });

  return (
    <div className="flex h-full min-w-0 flex-col bg-surface">
      <TabStrip state={tabs} onSelect={onSelectTab} onClose={onCloseTab} />

      <div className="min-h-0 flex-1">
        {current.kind === 'landing' ? (
          <LandingPage
            documents={documents}
            onAdd={onAdd}
            onRemove={onRemove}
            onCompare={onCompare}
            mode={mode}
            onModeChange={setMode}
            paste={paste}
            onPasteChange={setPaste}
            onComparePasted={onComparePair}
          />
        ) : current.kind === 'pairing' ? (
          current.status === 'building' ? (
            <Centred>
              <Loader2 size={22} className="animate-spin text-accent" aria-hidden />
              <p className="text-[13px] text-text">Reading documents and proposing pairs</p>
              <p className="max-w-prose text-center text-[11px] leading-relaxed text-text-faint">
                Every document on both sides is read and sketched before anything is
                proposed, so the grid can show real content similarity rather than guessing
                from filenames.
              </p>
            </Centred>
          ) : current.status === 'failed' ? (
            <Centred>
              <AlertTriangle size={22} className="text-danger" aria-hidden />
              <p className="text-[13px] font-medium text-text">Pairing failed</p>
              <p className="max-w-prose text-center text-[12px] text-text-muted">
                {current.message}
              </p>
            </Centred>
          ) : (
            <PairingGrid
              proposal={(current.result as PairingResult).proposal}
              onRepoint={(rowId, b) => onRepoint(current.id, rowId, b)}
              onConfirm={() => onConfirmPairing(current.id)}
              busy={false}
            />
          )
        ) : current.kind === 'matrix' ? (
          <DeviationMatrixView matrix={current.result as DeviationMatrix} />
        ) : current.kind === 'self-redline' ? (
          <SelfRedlineView tab={current} onSave={saveEditedDocument} />
        ) : (
          <RedlineView tab={current} onCancel={() => onCancelTab(current.id)} />
        )}
      </div>
    </div>
  );
}

function Centred({ children }: { children: React.ReactNode }) {
  return (
    <div className="flex h-full flex-col items-center justify-center gap-3 px-8">{children}</div>
  );
}
