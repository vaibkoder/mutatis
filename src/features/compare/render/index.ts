/**
 * Compare → RenderCore adapter (feature 16 / ADR-0024, Phase 2 B2/B3).
 *
 * The host-free glue between Compare's `RenderSource` (surfaced by the driver in
 * B1) and the RenderCore engine. It:
 *   - `buildRenderInput` — maps a `RenderSource` to the engine's `RenderInput`,
 *     translating Compare's `classByRevId` into RenderCore's annotation types;
 *   - `extractEmbeddedFaces` — pulls de-obfuscated embedded `.odttf` faces from
 *     the source bundle (embedded wins over the OFL clones);
 *   - `renderRedlinePdf` — runs the layout off the main thread (worker) and paints
 *     the shared `LayoutResult` as a vector PDF (`Uint8Array`);
 *   - `mountRedlinePreview` — runs the same layout and mounts the paginated DOM
 *     preview into a host element.
 *
 * Host-free by discipline: imports only `@shared/*` (the engine + OOXML), never
 * `Office/Word`. The Web Worker is created via `new Worker(new URL(...))` (a
 * same-origin module worker — no `fetch`); it is **injectable** so unit tests run
 * the real pipeline in-process without a real Worker.
 *
 * Never-suppress (ADR-0008): class annotations are styling only; the layout
 * renders every native revision regardless of class. The editable `.docx`/host-doc
 * redline is untouched — this path only ever reads the surfaced `RenderSource`.
 */
import {
  createRenderClient, createFontResolver, renderPdf, renderPreview,
  renderToLayout, createHarfBuzzShaper,
  type RenderClient, type RenderInput, type FaceInput, type LayoutResult,
  type RenderTheme, type PreviewHandle, type PdfMetadata, type WorkerLike,
} from '@shared/render-core/index.js';
import { loadRenderFaces, registerPreviewFonts } from '@shared/render-core/assets/faces.js';
import type { RenderSource } from '../ui-types.js';
import { toRcClassMap } from './classes.js';
import { extractEmbeddedFaces } from './embedded-fonts.js';

export { toRcChangeClass, toRcClassMap } from './classes.js';
export { extractEmbeddedFaces } from './embedded-fonts.js';

/** Options shared by both render entry points. */
export interface RedlineRenderOptions {
  theme?: RenderTheme;
  onProgress?: (pct: number) => void;
  /**
   * Injected render client. Defaults to a shared worker-backed client (one Web
   * Worker per session, reused across renders). Tests inject an in-process fake.
   */
  client?: RenderClient;
}

export interface RedlinePdfOptions extends RedlineRenderOptions {
  meta?: PdfMetadata;
}

export interface RedlinePreviewOptions extends RedlineRenderOptions {
  /** Document whose `FontFace`s the clone display fonts are registered on. */
  ownerDocument?: Document;
}

/** Map a driver `RenderSource` to the engine `RenderInput` (classes → annotations). */
export function buildRenderInput(src: RenderSource): RenderInput {
  const classByRevId = src.classByRevId ? toRcClassMap(src.classByRevId) : undefined;
  return {
    authoredBodyXml: src.authoredBodyXml,
    ...(src.stylesXml !== undefined ? { stylesXml: src.stylesXml } : {}),
    ...(src.footnotesXml !== undefined ? { footnotesXml: src.footnotesXml } : {}),
    ...(src.numbering !== undefined ? { numbering: src.numbering } : {}),
    ...(classByRevId !== undefined ? { classByRevId } : {}),
  };
}

/** Faces for a render: vendored clones + any embedded `.odttf` (embedded wins). */
function facesFor(src: RenderSource): FaceInput[] {
  return [...extractEmbeddedFaces(src.bundle), ...loadRenderFaces()];
}

/**
 * A main-thread render client: runs the (host-free) engine in-process instead of
 * over a Worker. Same parse→layout→paginate→decorate pipeline, just synchronous —
 * the only cost is that a large render blocks the UI thread briefly. Used where a
 * module Worker can't be constructed (e.g. a page opened from disk over `file://`,
 * which blocks `new Worker(...)`); the engine's font bytes still arrive in-message
 * and HarfBuzz's wasm is a bundled (data-URL) asset, so there is no `fetch` either.
 */
function createInProcessClient(): RenderClient {
  let nextId = 1;
  return {
    render(input, faces, opts = {}): { id: number; layout: Promise<LayoutResult> } {
      const id = nextId++;
      const layout = (async (): Promise<LayoutResult> => {
        const shaper = await createHarfBuzzShaper();
        const fonts = createFontResolver(faces);
        const result = renderToLayout(input, fonts, shaper, opts.theme ? { theme: opts.theme } : {});
        opts.onProgress?.(100);
        return result;
      })();
      return { id, layout };
    },
    cancel(): void {},
    terminate(): void {},
  };
}

/** Build a module Worker for off-thread render, or null if the host forbids it. */
function tryCreateWorkerClient(): RenderClient | null {
  // The standalone single-file build forces in-process (no Worker chunk to fetch);
  // the define makes this branch statically true so the Worker code tree-shakes out.
  if ((import.meta as { env?: Record<string, string | undefined> }).env?.VITE_DP_INPROCESS === '1') {
    return null;
  }
  try {
    const worker = new Worker(
      new URL('../../../workers/render-core.worker.ts', import.meta.url),
      { type: 'module' },
    );
    // The DOM `Worker`'s `onmessage` is typed against `MessageEvent`; the engine's
    // `WorkerLike` narrows the message data. Structurally compatible at runtime —
    // the protocol is the contract — so bridge them through the seam type.
    return createRenderClient(worker as unknown as WorkerLike);
  } catch {
    return null;
  }
}

// One render client per session (lazy — never created under tests, which always
// inject a client, nor at import time, so non-render flows pay nothing). Prefers
// the off-main-thread Worker; falls back to the in-process engine where a Worker
// is unavailable (e.g. `file://`) so the render still works.
let sharedClient: RenderClient | null = null;
function getClient(opts: RedlineRenderOptions): RenderClient {
  if (opts.client) return opts.client;
  if (!sharedClient) {
    sharedClient = tryCreateWorkerClient() ?? createInProcessClient();
  }
  return sharedClient;
}

/** Run the off-main-thread layout for a render source + faces. */
async function computeLayout(
  src: RenderSource, faces: FaceInput[], opts: RedlineRenderOptions,
): Promise<LayoutResult> {
  const job = getClient(opts).render(buildRenderInput(src), faces, {
    ...(opts.theme !== undefined ? { theme: opts.theme } : {}),
    ...(opts.onProgress !== undefined ? { onProgress: opts.onProgress } : {}),
  });
  return job.layout;
}

/**
 * Render the compare redline to a programmatic vector PDF (`Uint8Array`). Layout
 * runs in the worker; the main thread only paints. The PDF paginates identically
 * to the preview (one shared `LayoutResult`).
 */
export async function renderRedlinePdf(
  src: RenderSource, opts: RedlinePdfOptions = {},
): Promise<Uint8Array> {
  const faces = facesFor(src);
  const layout = await computeLayout(src, faces, opts);
  // The SAME faces array backs the PDF subset embed (matches the layout's metrics).
  const fonts = createFontResolver(faces);
  return renderPdf(layout, fonts, opts.meta ?? {});
}

/**
 * Mount the paginated redline preview into `el`. Async — layout runs in the
 * worker. Returns the `PreviewHandle` (caller may `refresh()`/`destroy()`).
 */
export async function mountRedlinePreview(
  el: HTMLElement, src: RenderSource, opts: RedlinePreviewOptions = {},
): Promise<PreviewHandle> {
  // Register the clone faces for DISPLAY (metrics still come from the bytes).
  registerPreviewFonts(opts.ownerDocument ?? el.ownerDocument);
  const faces = facesFor(src);
  const layout = await computeLayout(src, faces, opts);
  const handle = renderPreview(layout, opts.theme !== undefined ? { theme: opts.theme } : {});
  el.appendChild(handle.root);
  return handle;
}

// ---------------------------------------------------------------------------
// Local download (the read-only rendered-redline snapshot artefact, ADR-0020).
// Host-free: a same-document anchor + object URL — NO network, no `fetch`. We do
// NOT reuse host-web's `triggerDownload` (that lives behind the Office/web host
// seam); this path must stay importable by both hosts from `@features/compare`.
// ---------------------------------------------------------------------------

/**
 * The download filename for the rendered-redline snapshot (ADR-0020 framing:
 * a read-only snapshot that complements — never replaces — the editable `.docx`).
 * The date is injected so the function stays deterministic for tests.
 */
export function redlinePdfFilename(date: Date): string {
  const iso = date.toISOString().split('T')[0] ?? 'export';
  return `mutatis-rendered-redline-${iso}.pdf`;
}

/**
 * Trigger a local download of PDF bytes. Guarded so it is a no-op where the DOM
 * download primitives are absent (worker / SSR). Confidentiality (ADR-0001/0003):
 * an in-memory Blob + object URL — the bytes never leave the machine.
 */
export function downloadPdf(bytes: Uint8Array, filename: string): void {
  if (typeof document === 'undefined' || typeof URL.createObjectURL !== 'function') return;
  // Copy into a fresh ArrayBuffer-backed view so the Blob part type is exact
  // (mirrors host-web's writeBack; avoids the SharedArrayBuffer-union mismatch).
  const blob = new Blob([new Uint8Array(bytes)], { type: 'application/pdf' });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  // Anchor must be in the DOM for programmatic .click() in some engines
  // (mirrors host-web's triggerDownload — the project's established pattern).
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}
