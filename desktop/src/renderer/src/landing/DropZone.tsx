import { useCallback, useId, useRef, useState } from 'react';
import { FilePlus2, FileWarning, Loader2, X } from 'lucide-react';
import {
  acceptFiles,
  describeBytes,
  overCapReason,
  type LoadedDoc,
  type RejectedFile,
} from '../state/documents.js';
import { convertUploads } from '../paste/convert-uploads.js';

interface DropZoneProps {
  title: string;
  hint: string;
  docs: readonly LoadedDoc[];
  onAdd: (docs: readonly LoadedDoc[]) => void;
  onRemove: (id: string) => void;
}

export function DropZone({ title, hint, docs, onAdd, onRemove }: DropZoneProps) {
  const [over, setOver] = useState(false);
  const [rejected, setRejected] = useState<readonly RejectedFile[]>([]);
  const [reading, setReading] = useState(false);
  const inputId = useId();
  // Counter, not a boolean: dragenter/dragleave fire for every descendant, so a
  // boolean flickers off the moment the pointer crosses an inner element.
  const depth = useRef(0);

  const ingest = useCallback(
    (files: readonly File[]) => {
      const result = acceptFiles(files, docs);
      setRejected(result.rejected);
      if (result.accepted.length === 0) return;
      // .txt, .md and .pdf are converted to .docx first: the engine reads only
      // OOXML. Reading a PDF takes a moment, so say so; one that cannot be read
      // is reported here with the reason rather than failing later.
      setReading(result.accepted.some((d) => d.kind === 'pdf'));
      void convertUploads(result.accepted)
        .then(({ docs: ready, failed }) => {
          if (ready.length > 0) onAdd(ready);
          if (failed.length > 0) setRejected((prev) => [...prev, ...failed]);
        })
        .finally(() => setReading(false));
    },
    [docs, onAdd],
  );

  return (
    <section
      onDragEnter={(e) => {
        e.preventDefault();
        depth.current += 1;
        setOver(true);
      }}
      onDragOver={(e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'copy';
      }}
      onDragLeave={(e) => {
        e.preventDefault();
        depth.current -= 1;
        if (depth.current <= 0) setOver(false);
      }}
      onDrop={(e) => {
        e.preventDefault();
        depth.current = 0;
        setOver(false);
        ingest(Array.from(e.dataTransfer.files));
      }}
      className={
        'flex min-h-0 flex-1 flex-col rounded-lg border-2 border-dashed transition-colors ' +
        (over ? 'border-accent bg-accent-soft' : 'border-border bg-surface-raised')
      }
    >
      <div className="flex shrink-0 items-baseline gap-2 px-4 pt-3">
        <h2 className="text-[12px] font-semibold tracking-wide uppercase text-text-muted">
          {title}
        </h2>
        {docs.length > 0 ? (
          <span className="text-[11px] text-text-faint">
            {docs.length} {docs.length === 1 ? 'document' : 'documents'}
          </span>
        ) : null}
      </div>

      {docs.length === 0 ? (
        <div className="flex flex-1 flex-col items-center justify-center gap-3 px-6 text-center">
          <FilePlus2 size={22} className="text-text-faint" aria-hidden />
          <p className="text-[13px] text-text-muted">{hint}</p>
          <p className="text-[11px] text-text-faint">.docx · .pdf · .txt · .md</p>
          <BrowseButton id={inputId} onFiles={ingest} />
        </div>
      ) : (
        <>
          <ul className="min-h-0 flex-1 overflow-y-auto px-2 py-2">
            {docs.map((doc) => (
              <li key={doc.id} className="group rounded px-2 py-1.5 hover:bg-border/40">
                <div className="flex items-baseline gap-2">
                  <span className="min-w-0 flex-1 truncate text-[13px]" title={doc.name}>
                    {doc.name}
                  </span>
                  <span className="shrink-0 text-[11px] tabular-nums text-text-faint">
                    {describeBytes(doc.byteLength)}
                  </span>
                  <button
                    type="button"
                    onClick={() => onRemove(doc.id)}
                    aria-label={`Remove ${doc.name}`}
                    className="shrink-0 rounded p-0.5 text-text-faint opacity-0 transition hover:bg-border hover:text-text group-hover:opacity-100 focus-visible:opacity-100"
                  >
                    <X size={13} aria-hidden />
                  </button>
                </div>
                {doc.overCap ? (
                  <p className="mt-0.5 flex items-start gap-1 text-[11px] leading-snug text-danger">
                    <FileWarning size={12} className="mt-px shrink-0" aria-hidden />
                    <span>{overCapReason(doc)}</span>
                  </p>
                ) : null}
              </li>
            ))}
          </ul>
          <div className="shrink-0 border-t border-border px-3 py-2">
            <BrowseButton id={inputId} onFiles={ingest} compact />
          </div>
        </>
      )}

      {reading ? (
        <p role="status" className="flex shrink-0 items-center gap-2 border-t border-border px-3 py-2 text-[11px] text-text-muted">
          <Loader2 size={12} className="animate-spin text-accent" aria-hidden />
          Reading PDF…
        </p>
      ) : null}

      {rejected.length > 0 ? (
        <div className="shrink-0 border-t border-border px-3 py-2">
          <p className="pb-1 text-[11px] font-medium text-danger">
            {rejected.length} {rejected.length === 1 ? 'file was' : 'files were'} not added
          </p>
          <ul className="flex flex-col gap-0.5">
            {rejected.map((r) => (
              <li key={r.name} className="text-[11px] leading-snug text-text-muted">
                <span className="font-medium">{r.name}</span> — {r.reason}
              </li>
            ))}
          </ul>
        </div>
      ) : null}
    </section>
  );
}

function BrowseButton({
  id,
  onFiles,
  compact = false,
}: {
  id: string;
  onFiles: (files: readonly File[]) => void;
  compact?: boolean;
}) {
  return (
    <>
      <input
        id={id}
        type="file"
        multiple
        accept=".docx,.pdf,.txt,.md"
        className="sr-only"
        onChange={(e) => {
          onFiles(Array.from(e.target.files ?? []));
          // Reset so choosing the same file twice still fires a change event.
          e.target.value = '';
        }}
      />
      <label
        htmlFor={id}
        className={
          'cursor-pointer rounded-md border border-border bg-surface text-text-muted transition-colors hover:border-border-strong hover:text-text ' +
          (compact ? 'px-2 py-1 text-[11px]' : 'px-3 py-1.5 text-[12px]')
        }
      >
        Browse files
      </label>
    </>
  );
}
