import { FileText, ListChecks } from 'lucide-react';
import type { DocumentSet, LoadedDoc } from '../state/documents.js';
import type { TabState } from '../tabs/model.js';
import { TaskPlan } from './TaskPlan.js';
import { describeBytes } from '../state/documents.js';

/**
 * Right column: the documents in play, plus the task-plan checklist.
 *
 * The checklist is the thing that makes "comparing document 37 of 100" legible,
 * so it lives here from the start even though nothing drives it until phase 3.
 */
export function ContextRail({ documents, tabs }: { documents: DocumentSet; tabs: TabState }) {
  const total = documents.original.length + documents.modified.length;

  return (
    <div className="flex h-full flex-col border-l border-border bg-surface-sunken">
      <div className="flex h-9 shrink-0 items-center gap-2 border-b border-border px-3">
        <FileText size={13} className="text-text-muted" aria-hidden />
        <span className="text-[11px] font-semibold tracking-wide uppercase text-text-muted">
          Documents
        </span>
        {total > 0 ? (
          <span className="ml-auto rounded bg-border px-1.5 py-0.5 text-[10px] font-medium text-text-muted">
            {total}
          </span>
        ) : null}
      </div>

      <div className="flex-1 overflow-y-auto">
        {total === 0 ? (
          <p className="px-3 py-4 text-[11px] leading-relaxed text-text-faint">
            Documents you add appear here.
          </p>
        ) : (
          <>
            <DocGroup title="Original" docs={documents.original} />
            <DocGroup title="Modified" docs={documents.modified} />
          </>
        )}
      </div>

      <div className="shrink-0 border-t border-border">
        <div className="flex h-9 items-center gap-2 px-3">
          <ListChecks size={13} className="text-text-muted" aria-hidden />
          <span className="text-[11px] font-semibold tracking-wide uppercase text-text-muted">
            Task plan
          </span>
        </div>
        <TaskPlan tabs={tabs} />
      </div>
    </div>
  );
}

function DocGroup({ title, docs }: { title: string; docs: readonly LoadedDoc[] }) {
  if (docs.length === 0) return null;
  return (
    <section className="px-2 py-2">
      <h3 className="px-1 pb-1 text-[10px] font-semibold tracking-wide uppercase text-text-faint">
        {title}
      </h3>
      <ul className="flex flex-col gap-0.5">
        {docs.map((doc) => (
          <li
            key={doc.id}
            className="flex items-baseline gap-2 rounded px-1 py-1 hover:bg-border/50"
          >
            <span className="min-w-0 flex-1 truncate text-[12px]" title={doc.name}>
              {doc.name}
            </span>
            <span className="shrink-0 text-[10px] tabular-nums text-text-faint">
              {describeBytes(doc.byteLength)}
            </span>
          </li>
        ))}
      </ul>
    </section>
  );
}
