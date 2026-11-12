/**
 * Node PDF opener for tests — the test-side {@link PdfOpener}.
 *
 * Mirrors the production loader's hardened options but resolves the worker +
 * standard-font assets from the installed `pdfjs-dist` package on disk (the
 * browser loader serves them from `/pdf-assets/`). TEST-ONLY: filesystem paths
 * are fine here; this never ships. Lets the extraction/reconstruction tests run
 * headless against the committed fixtures.
 */
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';

import type {
  PdfDocumentProxyLike,
  PdfImageOps,
  PdfOpener,
} from '../../../src/features/compare/io/pdf/pdfjs-types.js';
import { HARDENED_DOC_OPTIONS } from '../../../src/features/compare/io/pdf/pdfjs-config.js';

const require = createRequire(import.meta.url);

let pdfjsPromise: Promise<typeof import('pdfjs-dist')> | null = null;

async function loadPdfjs(): Promise<typeof import('pdfjs-dist')> {
  if (pdfjsPromise === null) {
    const pdfjsPath = require.resolve('pdfjs-dist/legacy/build/pdf.mjs');
    const workerPath = require.resolve('pdfjs-dist/legacy/build/pdf.worker.mjs');
    pdfjsPromise = import(pdfjsPath).then((mod: typeof import('pdfjs-dist')) => {
      mod.GlobalWorkerOptions.workerSrc = workerPath;
      return mod;
    });
  }
  return pdfjsPromise;
}

/** Open raw PDF bytes with the hardened, offline options (Node asset paths). */
export const nodePdfOpener: PdfOpener = async (data: Uint8Array): Promise<PdfDocumentProxyLike> => {
  const pdfjs = await loadPdfjs();
  const fontsDir = require.resolve('pdfjs-dist/package.json').replace('package.json', 'standard_fonts/');
  const cmapsDir = require.resolve('pdfjs-dist/package.json').replace('package.json', 'cmaps/');
  const task = pdfjs.getDocument({
    data,
    ...HARDENED_DOC_OPTIONS,
    standardFontDataUrl: fontsDir,
    cMapUrl: cmapsDir,
    cMapPacked: true,
  });
  const doc = await task.promise;
  return doc as unknown as PdfDocumentProxyLike;
};

/** The D3 OPS subset, sourced from the real pdfjs.OPS map (Node test seam). */
export async function loadPdfImageOps(): Promise<PdfImageOps> {
  const pdfjs = await loadPdfjs();
  const OPS = pdfjs.OPS;
  return {
    save: OPS.save,
    restore: OPS.restore,
    transform: OPS.transform,
    paintImageXObject: OPS.paintImageXObject,
    paintInlineImageXObject: OPS.paintInlineImageXObject,
    paintImageMaskXObject: OPS.paintImageMaskXObject,
  };
}

/** Read a committed fixture as bytes. */
export function readFixtureBytes(path: string): Uint8Array {
  return new Uint8Array(readFileSync(path));
}
