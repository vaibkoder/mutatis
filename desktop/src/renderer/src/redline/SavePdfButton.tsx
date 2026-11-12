import { useState } from 'react';
import { FileText, Loader2 } from 'lucide-react';
import type { RenderModel } from '@shared/render-core/types.js';

/**
 * Typeset a redline to PDF and save it.
 *
 * Everything the PDF needs — ~1 MB of vendored font faces, the HarfBuzz wasm
 * module and a PDF writer — is behind a dynamic import, so a session that never
 * asks for a PDF never loads any of it.
 *
 * The result is paginated by the SAME `LayoutResult` the engine's own preview
 * uses, so what a reader sees in Acrobat is what the engine laid out. The fonts
 * are the four metric clones embedded at build time; nothing is fetched at
 * runtime, which matters for a document that may be confidential.
 */
export function SavePdfButton({
  model,
  title,
  author,
  fileName,
}: {
  model: RenderModel;
  title: string;
  author: string;
  fileName: string;
}) {
  const [state, setState] = useState<'idle' | 'working' | { error: string }>('idle');

  return (
    <div className="flex items-center gap-2">
      <button
        type="button"
        data-pdf-button
        disabled={state === 'working'}
        onClick={() => {
          setState('working');
          void (async () => {
            try {
              const { renderPdfRedline, savePdf } = await import('../edits/pdf-redline.js');
              const bytes = await renderPdfRedline(model, { title, author });
              savePdf(bytes, fileName);
              setState('idle');
            } catch (error) {
              setState({ error: error instanceof Error ? error.message : String(error) });
            }
          })();
        }}
        className="flex shrink-0 items-center gap-1.5 rounded border border-border px-2.5 py-1 text-[11px] font-medium text-text transition-colors hover:border-border-strong disabled:opacity-50"
      >
        {state === 'working' ? (
          <Loader2 size={12} className="animate-spin" aria-hidden />
        ) : (
          <FileText size={12} aria-hidden />
        )}
        {state === 'working' ? 'Typesetting…' : 'Save PDF redline'}
      </button>
      {typeof state === 'object' && (
        <span className="text-[11px] text-danger">PDF failed: {state.error}</span>
      )}
    </div>
  );
}
