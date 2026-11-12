import { HARDENED_DOC_OPTIONS } from '@features/compare/io/pdf/pdfjs-config.js';
import type { PdfDocumentProxyLike, PdfOpener } from '@features/compare/io/pdf/pdfjs-types.js';

/**
 * PDF.js's Node build, opened with the SAME hardened options as the app's opener.
 * The app's opener needs a browser (a worker URL); this one runs PDF.js in-process
 * so the tests can read real PDFs.
 */
export const nodePdfOpener: PdfOpener = async (data) => {
  const { getDocument } = await import('pdfjs-dist/legacy/build/pdf.mjs');
  const task = getDocument({ data, ...HARDENED_DOC_OPTIONS, verbosity: 0 });
  return (await task.promise) as unknown as PdfDocumentProxyLike;
};
