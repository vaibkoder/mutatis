import { ArrowRight } from 'lucide-react';
import { PastePanel } from './PastePanel.js';
import { buildPastedPair, sideBlocks, type PasteSide } from './pasted.js';
import { hasContent } from './model.js';
import type { LoadedDoc } from '../state/documents.js';

export interface PasteState {
  readonly original: PasteSide;
  readonly modified: PasteSide;
}

interface PasteCompareProps {
  state: PasteState;
  onChange: (next: PasteState) => void;
  onCompare: (original: LoadedDoc, modified: LoadedDoc) => void;
}

/** Two paste boxes and a Compare button: a one-to-one comparison of pasted text. */
export function PasteCompare({ state, onChange, onCompare }: PasteCompareProps) {
  const ready = hasContent(sideBlocks(state.original)) && hasContent(sideBlocks(state.modified));

  return (
    <div className="flex min-h-0 flex-1 flex-col">
      <div className="flex min-h-0 flex-1 gap-3 px-3 pb-3">
        <PastePanel
          title="Original"
          hint="Paste the original text here. Copying from Word, Google Docs or a web page keeps its formatting."
          side={state.original}
          onChange={(original) => onChange({ ...state, original })}
        />
        <PastePanel
          title="Modified"
          hint="Paste the modified text here."
          side={state.modified}
          onChange={(modified) => onChange({ ...state, modified })}
        />
      </div>

      <footer className="flex h-12 shrink-0 items-center gap-3 border-t border-border bg-surface-raised px-4">
        <span className="text-[12px] text-text-muted">
          {ready ? 'Two pasted texts, compared one-to-one' : 'Paste text into both boxes'}
        </span>
        <button
          type="button"
          disabled={!ready}
          onClick={() => {
            const pair = buildPastedPair(state.original, state.modified);
            if (pair) onCompare(pair.original, pair.modified);
          }}
          className={
            'ml-auto flex items-center gap-1.5 rounded-md px-3 py-1.5 text-[12px] font-medium transition-colors ' +
            (ready ? 'bg-accent text-accent-text hover:opacity-90' : 'cursor-not-allowed bg-border text-text-faint')
          }
        >
          Compare pasted text
          <ArrowRight size={13} aria-hidden />
        </button>
      </footer>
    </div>
  );
}
