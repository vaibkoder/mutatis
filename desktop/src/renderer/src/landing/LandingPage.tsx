import { ArrowRight, ClipboardPaste, Files } from 'lucide-react';
import { PasteCompare, type PasteState } from '../paste/PasteCompare.js';
import { DropZone } from './DropZone.js';
import {
  comparisonShape,
  describeShape,
  type DocSide,
  type DocumentSet,
  type LoadedDoc,
} from '../state/documents.js';

interface LandingPageProps {
  documents: DocumentSet;
  onAdd: (side: DocSide, docs: readonly LoadedDoc[]) => void;
  onRemove: (side: DocSide, id: string) => void;
  onCompare: () => void;
  mode: LandingMode;
  onModeChange: (mode: LandingMode) => void;
  paste: PasteState;
  onPasteChange: (next: PasteState) => void;
  onComparePasted: (original: LoadedDoc, modified: LoadedDoc) => void;
}

export type LandingMode = 'files' | 'paste';

const MODES: ReadonlyArray<{
  value: LandingMode;
  label: string;
  detail: string;
  Icon: typeof Files;
}> = [
  { value: 'files', label: 'Compare files', detail: 'Upload .docx, .pdf, .txt or .md files', Icon: Files },
  { value: 'paste', label: 'Compare pasted text', detail: 'Paste from Word, Google Docs, the web or Markdown', Icon: ClipboardPaste },
];

/**
 * The switch between comparing files and comparing pasted text. Deliberately
 * large and labelled: it decides what the whole page below it is, so it must not
 * be a small toggle that is easy to miss.
 */
function ModeSwitch({ mode, onChange }: { mode: LandingMode; onChange: (mode: LandingMode) => void }) {
  return (
    <div role="group" aria-label="What to compare" className="flex shrink-0 px-3 pt-3">
      {MODES.map(({ value, label, detail, Icon }, index) => {
        const selected = mode === value;
        return (
          <button
            key={value}
            type="button"
            aria-pressed={selected}
            onClick={() => onChange(value)}
            className={
              // The chosen mode takes three quarters of the row and the other one
              // quarter; flex-grow animates, so the switch glides. Motion is off
              // for anyone who has asked their system for reduced motion.
              // One joined control: outer corners only, and the second button overlaps
              // the first by its border width so a single divider remains. The
              // selected one sits on top so its accent border is never covered.
              'min-w-0 overflow-hidden border-2 p-0 text-left ' +
              (index === 0 ? 'rounded-l-lg rounded-r-none ' : '-ml-0.5 rounded-l-none rounded-r-lg ') +
              (selected ? 'relative z-10 ' : '') +
              'transition-[flex-grow,background-color,border-color,color] duration-300 ease-out motion-reduce:transition-none ' +
              (selected
                ? 'grow-3 basis-0 border-accent bg-accent text-accent-text'
                : 'grow basis-0 border-border bg-surface-raised text-text hover:border-border-strong')
            }
          >
            <span className="flex items-center gap-3 px-4 py-3">
              <Icon size={22} aria-hidden className="shrink-0" />
              <span className="flex min-w-0 flex-col">
                <span className="truncate text-[14px] font-semibold">{label}</span>
                <span className={'truncate text-[11px] ' + (selected ? 'opacity-85' : 'text-text-muted')}>
                  {detail}
                </span>
              </span>
            </span>
          </button>
        );
      })}
    </div>
  );
}

export function LandingPage({
  documents,
  onAdd,
  onRemove,
  onCompare,
  mode,
  onModeChange,
  paste,
  onPasteChange,
  onComparePasted,
}: LandingPageProps) {
  const shape = comparisonShape(documents);
  const ready = shape !== 'none';

  if (mode === 'paste') {
    return (
      <div className="flex h-full flex-col overflow-hidden">
        <ModeSwitch mode={mode} onChange={onModeChange} />
        <div className="h-3 shrink-0" />
        <PasteCompare state={paste} onChange={onPasteChange} onCompare={onComparePasted} />
      </div>
    );
  }

  return (
    <div className="flex h-full flex-col overflow-hidden">
      <ModeSwitch mode={mode} onChange={onModeChange} />
      <div className="flex min-h-0 flex-1 gap-3 p-3">
        <DropZone
          title="Original"
          hint="Drop the original document here"
          docs={documents.original}
          onAdd={(docs) => onAdd('original', docs)}
          onRemove={(id) => onRemove('original', id)}
        />
        <DropZone
          title="Modified"
          hint="Drop the modified document here"
          docs={documents.modified}
          onAdd={(docs) => onAdd('modified', docs)}
          onRemove={(id) => onRemove('modified', id)}
        />
      </div>

      <footer className="flex h-12 shrink-0 items-center gap-3 border-t border-border bg-surface-raised px-4">
        <span className="text-[12px] text-text-muted">{describeShape(shape)}</span>
        <button
          type="button"
          disabled={!ready}
          onClick={onCompare}
          title={
            shape === 'many-to-many'
              ? 'You will be asked to confirm how the documents pair up before anything runs'
              : undefined
          }
          className={
            'ml-auto flex items-center gap-1.5 rounded-md px-3 py-1.5 text-[12px] font-medium transition-colors ' +
            (ready
              ? 'bg-accent text-accent-text hover:opacity-90'
              : 'cursor-not-allowed bg-border text-text-faint')
          }
        >
          {shape === 'many-to-many' ? 'Pair and compare' : 'Compare'}
          <ArrowRight size={13} aria-hidden />
        </button>
      </footer>
    </div>
  );
}
