/**
 * Browser/Vite PDF.js opener — the production seam that
 * opens raw PDF bytes into a hardened, fully-offline document proxy.
 *
 * Confidentiality (ADR-0001/0003) is enforced on TWO axes here:
 *  - every `getDocument` call spreads {@link HARDENED_DOC_OPTIONS} (no wasm, no
 *    worker-fetch, no system fonts, no font-face); and
 *  - every asset URL (worker / cMaps / standard fonts / icc / wasm) is built
 *    from a SAME-ORIGIN local base and passed through {@link assertAssetUrlsLocal}
 *    BEFORE `getDocument` sees it — an off-origin URL throws, never fetches.
 *
 * The assets are vendored to `<app>/pdf-assets/` by the Vite plugin in
 * `vite.config.ts` (served same-origin in dev, copied into `dist/` on build), so
 * the default base is the relative `pdf-assets/` path under the app origin.
 *
 * `pdfjs-dist` is imported DYNAMICALLY so the ~1 MB bundle is code-split and only
 * pulled when the user actually runs a PDF compare (Architecture perf guidance).
 * This file is never imported by tests — extraction takes an injected opener.
 */
import {
  HARDENED_DOC_OPTIONS,
  assertAssetUrlsLocal,
  assertSameOriginOrLocal,
} from './pdfjs-config.js';
import { ensureReadableStreamAsyncIteration } from './readable-stream-async-iter.js';
import type { PdfDocumentProxyLike, PdfImageOps } from './pdfjs-types.js';

/** Where the vendored PDF.js assets are served (same-origin, root-relative). */
const DEFAULT_ASSET_BASE = '/pdf-assets/';

interface AssetUrls {
  workerSrc: string;
  cMapUrl: string;
  standardFontDataUrl: string;
  iccUrl: string;
  wasmUrl: string;
}

/** Resolve the vendored asset URLs against the app origin, then guard them. */
function resolveAssetUrls(base: string): AssetUrls {
  // Resolve relative to the document base so the result is an absolute
  // same-origin URL (in the browser); the guard re-checks the origin.
  const root = absolutise(base);
  const urls: AssetUrls = {
    workerSrc: `${root}pdf.worker.min.mjs`,
    cMapUrl: `${root}cmaps/`,
    standardFontDataUrl: `${root}standard_fonts/`,
    iccUrl: `${root}iccs/`,
    wasmUrl: `${root}wasm/`,
  };
  assertAssetUrlsLocal({ ...urls });
  return urls;
}

/** Turn a relative base into an absolute same-origin URL when possible. */
function absolutise(base: string): string {
  const docBase = (globalThis as { document?: { baseURI?: string } }).document?.baseURI;
  if (typeof docBase === 'string') {
    const u = new URL(base, docBase);
    return u.href.endsWith('/') ? u.href : `${u.href}/`;
  }
  return base.endsWith('/') ? base : `${base}/`;
}

/**
 * Open PDF bytes into a hardened document proxy. `assetBase` is overridable for
 * non-standard hosting; it must be same-origin (guarded).
 */
export async function openPdfDocument(
  data: Uint8Array,
  assetBase: string = DEFAULT_ASSET_BASE,
): Promise<PdfDocumentProxyLike> {
  const assets = resolveAssetUrls(assetBase);

  // WebKit (Office-on-the-web on macOS) lacks ReadableStream async iteration,
  // which PDF.js's getTextContent() relies on — shim it before any PDF.js work.
  ensureReadableStreamAsyncIteration();

  // Dynamic import ⇒ code-split. The bundler resolves 'pdfjs-dist' to the ESM
  // build; the worker is served from our vendored, same-origin asset folder.
  const pdfjs = await import('pdfjs-dist');
  assertSameOriginOrLocal(assets.workerSrc, 'workerSrc');
  pdfjs.GlobalWorkerOptions.workerSrc = assets.workerSrc;

  const loadingTask = pdfjs.getDocument({
    data,
    ...HARDENED_DOC_OPTIONS,
    cMapUrl: assets.cMapUrl,
    cMapPacked: true,
    standardFontDataUrl: assets.standardFontDataUrl,
    iccUrl: assets.iccUrl,
    wasmUrl: assets.wasmUrl,
  });
  const doc = await loadingTask.promise;
  return doc as unknown as PdfDocumentProxyLike;
}

/**
 * The D3 image-presence OPS subset, sourced from the dynamically-imported pdfjs
 * `OPS` map. Lets `images.ts` stay free of the heavy bundle (it takes these as a
 * parameter). Reuses the same code-split chunk as {@link openPdfDocument}.
 */
export async function loadPdfImageOps(): Promise<PdfImageOps> {
  const pdfjs = await import('pdfjs-dist');
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
