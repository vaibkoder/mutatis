import { useState } from 'react';
import { Check, Copy, Download, X } from 'lucide-react';
import type { StructureOutline } from '../structure/outline.js';
import { entryDetail, entryTitle, outlineToJson, outlineToMarkdown } from '../structure/outline.js';
import { downloadText } from './export-for-ai.js';

const BTN =
  'flex shrink-0 items-center gap-1.5 rounded border border-border px-2 py-0.5 text-[11px] font-medium text-text transition-colors hover:border-border-strong';

type Side = 'revised' | 'original';

/**
 * The clause structure of either document, beside the redline and never in front
 * of it: the redline keeps its full width minus this column, and nothing here
 * filters or hides a change. Copy and save give the outline as Markdown or JSON.
 */
export function StructurePanel({
  structure,
  originalName,
  revisedName,
  onClose,
}: {
  structure: { original: StructureOutline | null; revised: StructureOutline | null };
  originalName: string;
  revisedName: string;
  onClose: () => void;
}) {
  const [side, setSide] = useState<Side>('revised');
  const [copied, setCopied] = useState<boolean | 'failed'>(false);
  const outline = structure[side];
  const name = side === 'revised' ? revisedName : originalName;
  const baseName = name.replace(/\.[^.]+$/, '');

  return (
    <aside
      className="flex w-[360px] shrink-0 flex-col border-l border-border bg-surface"
      aria-label="Contract structure"
      data-structure-panel
    >
      <div className="flex shrink-0 items-center gap-2 border-b border-border px-3 py-1.5">
        <p className="text-[12px] font-semibold text-text">Structure</p>
        <div className="flex overflow-hidden rounded border border-border text-[11px]" role="tablist">
          {(['revised', 'original'] as const).map((s) => (
            <button
              key={s}
              type="button"
              role="tab"
              aria-selected={side === s}
              onClick={() => setSide(s)}
              className={
                'px-2 py-0.5 transition-colors ' +
                (side === s ? 'bg-accent text-white' : 'text-text-muted hover:text-text')
              }
            >
              {s === 'revised' ? 'Revised' : 'Original'}
            </button>
          ))}
        </div>
        <button
          type="button"
          onClick={onClose}
          className="ml-auto rounded p-0.5 text-text-muted hover:text-text"
          title="Close the structure panel"
          aria-label="Close the structure panel"
        >
          <X size={14} aria-hidden />
        </button>
      </div>

      {outline ? (
        <>
          <div className="flex shrink-0 flex-wrap items-center gap-1.5 border-b border-border px-3 py-1.5">
            <button
              type="button"
              className={BTN}
              title="Copy the outline as Markdown"
              onClick={() => {
                void (async () => {
                  try {
                    await navigator.clipboard.writeText(outlineToMarkdown(outline, name));
                    setCopied(true);
                  } catch {
                    setCopied('failed');
                  }
                  setTimeout(() => setCopied(false), 2000);
                })();
              }}
            >
              {copied === true ? <Check size={12} aria-hidden /> : <Copy size={12} aria-hidden />}
              {copied === true ? 'Copied' : copied === 'failed' ? 'Copy blocked — use Save' : 'Copy outline'}
            </button>
            <button
              type="button"
              className={BTN}
              title="Save the outline as a Markdown file"
              onClick={() =>
                downloadText(outlineToMarkdown(outline, name), 'text/markdown', `${baseName} (structure).md`)
              }
            >
              <Download size={12} aria-hidden />
              .md
            </button>
            <button
              type="button"
              className={BTN}
              title="Save the outline as a JSON file"
              onClick={() =>
                downloadText(outlineToJson(outline, name), 'application/json', `${baseName} (structure).json`)
              }
            >
              <Download size={12} aria-hidden />
              .json
            </button>
          </div>
          <p className="shrink-0 border-b border-border px-3 py-1 text-[10px] leading-snug text-text-faint">
            {summaryLine(outline)} Read from the layout — check a citation before relying on it.
          </p>
          <ol className="min-h-0 flex-1 select-text overflow-y-auto px-2 py-2 text-[11.5px] leading-snug">
            {outline.entries.map((e) => {
              const detail = entryDetail(e);
              return (
                <li
                  key={e.id}
                  className="flex gap-1.5 rounded px-1 py-0.5 hover:bg-surface-sunken"
                  style={{ paddingLeft: `${4 + Math.min(e.depth, 8) * 12}px` }}
                  data-kind={e.kind}
                >
                  <span
                    className={
                      'shrink-0 ' +
                      (e.kind === 'heading'
                        ? 'italic text-text-muted'
                        : e.kind === 'definition'
                          ? 'text-text'
                          : 'font-semibold text-text')
                    }
                  >
                    {entryTitle(e)}
                  </span>
                  {detail ? <span className="min-w-0 truncate text-text-muted">{detail}</span> : null}
                </li>
              );
            })}
            {outline.entries.length === 0 ? (
              <li className="px-1 text-text-muted">No structure found in this document.</li>
            ) : null}
          </ol>
        </>
      ) : (
        <p className="px-3 py-3 text-[12px] text-text-muted">
          This document&rsquo;s structure could not be read. The redline is unaffected.
        </p>
      )}
    </aside>
  );
}

function summaryLine(outline: StructureOutline): string {
  const c = outline.counts;
  const bits: string[] = [];
  const clauses = (c['clause'] ?? 0) + (c['subclause'] ?? 0);
  if (clauses) bits.push(`${clauses} clauses`);
  if (c['definition']) bits.push(`${c['definition']} definitions`);
  const containers = (c['schedule'] ?? 0) + (c['annex'] ?? 0) + (c['appendix'] ?? 0);
  if (containers) bits.push(`${containers} schedules/annexes`);
  return bits.length ? `${bits.join(' · ')}.` : '';
}
