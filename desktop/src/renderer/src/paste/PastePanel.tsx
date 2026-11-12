import { useMemo, type ClipboardEvent } from 'react';
import { ClipboardPaste, Eraser } from 'lucide-react';
import { BlocksPreview } from './BlocksPreview.js';
import { htmlToBlocks } from './html-to-blocks.js';
import { looksLikeMarkdown } from './markdown-to-blocks.js';
import { hasContent, hasFormatting, wordCount, type PasteFormat } from './model.js';
import { emptyPasteSide, sideBlocks, type PasteSide } from './pasted.js';

interface PastePanelProps {
  title: string;
  hint: string;
  side: PasteSide;
  onChange: (next: PasteSide) => void;
}

const FORMATS: ReadonlyArray<{ value: PasteFormat; label: string; title: string }> = [
  { value: 'formatted', label: 'Formatted', title: 'Keep the formatting of what you pasted (Word, Google Docs, web pages)' },
  { value: 'markdown', label: 'Markdown', title: 'Read the text as Markdown: # headings, **bold**, lists, tables' },
  { value: 'plain', label: 'Plain text', title: 'Compare the characters exactly as they are, no formatting' },
];

export function PastePanel({ title, hint, side, onChange }: PastePanelProps) {
  const blocks = useMemo(() => sideBlocks(side), [side]);
  const filled = hasContent(blocks);
  const words = wordCount(blocks);
  const formatted = side.format === 'formatted' && side.html !== null;

  const handlePaste = (e: ClipboardEvent) => {
    const html = e.clipboardData.getData('text/html');
    const text = e.clipboardData.getData('text/plain');
    if (html) {
      // Only keep the HTML if it says something a plain paste would lose.
      // Copying from a code editor, for instance, sends styled spans and nothing else.
      const parsed = htmlToBlocks(html);
      if (hasFormatting(parsed)) {
        e.preventDefault();
        onChange({ text, html, format: 'formatted' });
        return;
      }
    }
    if (side.text.trim() === '' || formatted) {
      e.preventDefault();
      onChange({ text, html: null, format: looksLikeMarkdown(text) ? 'markdown' : 'plain' });
    }
    // Otherwise it is an edit in the middle of existing text: let the browser do it.
  };

  return (
    <section className="flex min-h-0 min-w-0 flex-1 flex-col rounded-lg border-2 border-border bg-surface-raised focus-within:border-accent">
      <div className="flex shrink-0 flex-wrap items-center gap-x-3 gap-y-1.5 px-4 pt-3 pb-2">
        <h2 className="text-[12px] font-semibold tracking-wide uppercase text-text-muted">{title}</h2>
        <div role="group" aria-label={`${title} format`} className="ml-auto flex gap-1">
          {FORMATS.map((f) => {
            const selected = side.format === f.value;
            const disabled = f.value === 'formatted' && side.html === null;
            return (
              <button
                key={f.value}
                type="button"
                aria-pressed={selected}
                disabled={disabled}
                title={disabled ? 'Only available when you paste formatted text' : f.title}
                onClick={() => onChange({ ...side, format: f.value })}
                className={
                  'rounded-md border px-2.5 py-1 text-[11px] font-medium transition-colors ' +
                  (selected
                    ? 'border-accent bg-accent-soft text-text'
                    : disabled
                      ? 'cursor-not-allowed border-border text-text-faint opacity-50'
                      : 'border-border text-text-muted hover:border-border-strong hover:text-text')
                }
              >
                {f.label}
              </button>
            );
          })}
        </div>
      </div>

      {formatted ? (
        <div
          tabIndex={0}
          aria-label={`${title} — pasted text with formatting. Paste again to replace it.`}
          onPaste={handlePaste}
          className="mx-3 min-h-0 flex-1 overflow-y-auto rounded border border-border bg-surface px-3 py-2 outline-none"
        >
          <BlocksPreview blocks={blocks} />
        </div>
      ) : (
        <div className="flex min-h-0 flex-1 flex-col gap-2 px-3">
          <textarea
            value={side.text}
            onPaste={handlePaste}
            onChange={(e) => onChange({ ...side, text: e.target.value, html: null })}
            placeholder={hint}
            aria-label={`${title} text`}
            spellCheck={false}
            className={
              'w-full resize-none rounded border border-border bg-surface px-3 py-2 font-mono text-[12px] leading-snug text-text outline-none placeholder:font-sans placeholder:text-text-faint focus:border-accent ' +
              (side.format === 'markdown' && filled ? 'h-1/2 shrink-0' : 'min-h-0 flex-1')
            }
          />
          {side.format === 'markdown' && filled ? (
            <div className="min-h-0 flex-1 overflow-y-auto rounded border border-border bg-surface px-3 py-2">
              <BlocksPreview blocks={blocks} />
            </div>
          ) : null}
        </div>
      )}

      <div className="flex h-10 shrink-0 items-center gap-3 px-4 text-[11px] text-text-muted">
        {filled ? (
          <>
            <span className="tabular-nums">
              {blocks.length} {blocks.length === 1 ? 'paragraph' : 'paragraphs'} · {words} {words === 1 ? 'word' : 'words'}
            </span>
            {formatted ? <span className="text-redline-ins">Formatting kept</span> : null}
            {side.format === 'markdown' ? <span>Read as Markdown</span> : null}
            <button
              type="button"
              onClick={() => onChange(emptyPasteSide)}
              className="ml-auto flex items-center gap-1 rounded px-1.5 py-0.5 hover:bg-border hover:text-text"
            >
              <Eraser size={12} aria-hidden /> Clear
            </button>
          </>
        ) : (
          <span className="flex items-center gap-1.5 text-text-faint">
            <ClipboardPaste size={12} aria-hidden />
            Click the box, then paste (Ctrl+V or ⌘V)
          </span>
        )}
      </div>
    </section>
  );
}
