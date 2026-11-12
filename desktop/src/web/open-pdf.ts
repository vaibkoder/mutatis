import { HARDENED_DOC_OPTIONS } from '@features/compare/io/pdf/pdfjs-config.js';
import { ensureReadableStreamAsyncIteration } from '@features/compare/io/pdf/readable-stream-async-iter.js';
import type { PdfDocumentProxyLike, PdfOpener } from '@features/compare/io/pdf/pdfjs-types.js';
import PdfWorker from './pdf-worker.ts?worker&inline';

/**
 * Web build's replacement for renderer/src/pdf/open-pdf.ts (aliased in
 * vite.web.config.ts).
 *
 * From file:// there is no server to fetch a PDF.js worker, character maps or
 * fonts from, so the worker is carried inside the HTML (as a classic worker,
 * like the compare worker — a module worker will not start from file://) and
 * nothing is fetched. The same hardened options as the engine's own opener are
 * applied, so PDF.js still has no wasm, no worker fetch and no system fonts.
 *
 * Not available here: the predefined character maps and the standard-font data.
 * Text in a PDF that embeds its fonts reads normally; a PDF using a non-embedded
 * CJK font with a predefined CMap may read incompletely, and the paragraph
 * reconstruction can be slightly off for non-embedded standard fonts.
 */
export const openPdf: PdfOpener = async (data) => {
  ensureReadableStreamAsyncIteration();
  const pdfjs = await import('pdfjs-dist');
  // PDF.js's declaration types `port` as null; at runtime it takes a Worker.
  const worker = new pdfjs.PDFWorker({ port: new PdfWorker() as unknown as null });
  const task = pdfjs.getDocument({ data, ...HARDENED_DOC_OPTIONS, worker });
  const doc = (await task.promise) as unknown as PdfDocumentProxyLike;
  const destroy = doc.destroy.bind(doc);
  return Object.assign(doc, {
    destroy: async () => {
      await destroy();
      worker.destroy();
    },
  });
};
