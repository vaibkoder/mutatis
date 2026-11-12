import { openPdfDocument } from '@features/compare/io/pdf/pdfjs-loader.js';
import type { PdfOpener } from '@features/compare/io/pdf/pdfjs-types.js';

/**
 * The one place a PDF is opened.
 *
 * Electron uses the engine's own hardened opener: PDF.js with no wasm, no worker
 * fetch and no system fonts, whose worker, character maps and fonts are served
 * from this app at /pdf-assets/ (scripts/pdfjs-assets-plugin.mjs). The single-file
 * web build has no server to serve them from, so vite.web.config.ts aliases this
 * module to src/web/open-pdf.ts, which carries the worker inside the HTML.
 */
export const openPdf: PdfOpener = (data) => openPdfDocument(data);
