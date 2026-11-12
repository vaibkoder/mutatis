/**
 * Cross-engine (Chromium + WebKit) headless safety invariants for the render
 * path (feature 16 §5 acceptance 9).
 *
 * These guard the structural WebKit lessons the project has already paid for.
 * They are the *headless* half of acceptance 9; the live visual open in Safari /
 * Preview / Office-on-web-Mac is a manual smoke-test gate
 * (open the exported PDF and preview in each engine by hand). Set up, not faked.
 */
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { downloadPdf, redlinePdfFilename } from '@features/compare/render/index.js';
import { registerPreviewFonts, CLONE_DISPLAY_FAMILIES } from '@shared/render-core/assets/faces.js';

const root = resolve(__dirname, '../../..');
const readSrc = (rel: string): string => readFileSync(resolve(root, rel), 'utf8');
/** Strip block + line comments so doc-comments may legitimately name a banned API. */
const code = (rel: string): string =>
  readSrc(rel).replace(/\/\*[\s\S]*?\*\//g, '').replace(/\/\/[^\n]*/g, '');

describe('cross-engine: wasm is same-origin, never fetch-streamed (acceptance 9)', () => {
  it('the shaper never uses WebAssembly.instantiateStreaming (WebKit MIME-strict)', () => {
    const shape = code('src/shared/render-core/layout/shape.ts');
    expect(/instantiateStreaming/.test(shape)).toBe(false);
  });

  it('the wasm is loaded same-origin via import.meta.url, not an off-origin/CDN URL', () => {
    const shape = readSrc('src/shared/render-core/layout/shape.ts');
    // The harfbuzz wasm rides the Vite same-origin asset mechanism (new URL(..., import.meta.url));
    // no http(s)/CDN literal may reference it.
    expect(/https?:\/\/[^\s'"]*harfbuzz/i.test(shape)).toBe(false);
  });
});

describe('cross-engine: render path does not stream (no WebKit async-iter dep)', () => {
  // The webkit-readablestream-async-iter polyfill exists for PDF.js text
  // extraction (PDF-vs-PDF compare) — a DIFFERENT surface. The render path must
  // not depend on ReadableStream async iteration, so it needs no such polyfill.
  const renderPaths = [
    'src/workers/render-core.worker.ts',
    'src/shared/render-core/layout/shape.ts',
    'src/shared/render-core/output/pdf.ts',
    'src/shared/render-core/output/preview.ts',
    'src/features/compare/render/index.ts',
  ];
  for (const rel of renderPaths) {
    it(`${rel} uses no ReadableStream async-iteration`, () => {
      const c = code(rel);
      expect(/for\s+await/.test(c)).toBe(false);
      expect(/ReadableStream/.test(c)).toBe(false);
      expect(/\.getReader\s*\(/.test(c)).toBe(false);
    });
  }
});

describe('cross-engine: Blob download path is WebKit-safe (acceptance 9)', () => {
  let createObjectURL: ReturnType<typeof vi.fn>;
  let revokeObjectURL: ReturnType<typeof vi.fn>;
  let clickedWhileMounted: boolean;

  beforeEach(() => {
    createObjectURL = vi.fn(() => 'blob:mock-url');
    revokeObjectURL = vi.fn();
    // jsdom does not implement these; install deterministic spies.
    (URL as unknown as { createObjectURL: unknown }).createObjectURL = createObjectURL;
    (URL as unknown as { revokeObjectURL: unknown }).revokeObjectURL = revokeObjectURL;
    clickedWhileMounted = false;
    // Capture whether the anchor is attached to the DOM at .click() time —
    // WebKit (Safari) ignores programmatic clicks on detached anchors.
    vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(function (this: HTMLAnchorElement) {
      clickedWhileMounted = document.body.contains(this);
    });
  });
  afterEach(() => {
    vi.restoreAllMocks();
    delete (URL as unknown as { createObjectURL?: unknown }).createObjectURL;
    delete (URL as unknown as { revokeObjectURL?: unknown }).revokeObjectURL;
  });

  it('mounts the anchor before clicking, then removes it and revokes the URL', () => {
    const bytes = new Uint8Array([0x25, 0x50, 0x44, 0x46]); // %PDF
    downloadPdf(bytes, redlinePdfFilename(new Date('2026-06-15T00:00:00Z')));

    expect(createObjectURL).toHaveBeenCalledTimes(1);
    const blob = createObjectURL.mock.calls[0]![0] as Blob;
    expect(blob.type).toBe('application/pdf');
    expect(clickedWhileMounted).toBe(true); // attached at click time (WebKit-safe)
    expect(document.querySelector('a[download]')).toBeNull(); // cleaned up
    expect(revokeObjectURL).toHaveBeenCalledWith('blob:mock-url');
  });

  it('the snapshot filename carries the ISO date and .pdf extension', () => {
    expect(redlinePdfFilename(new Date('2026-06-15T00:00:00Z')))
      .toBe('mutatis-rendered-redline-2026-06-15.pdf');
  });
});

describe('cross-engine: preview FontFace registration is fetch-free + quote-safe', () => {
  it('registers display faces from BufferSource bytes — never a url() (no fetch)', () => {
    const ctorArgs: Array<{ family: string; source: unknown }> = [];
    const added: unknown[] = [];
    class FakeFontFace {
      family: string;
      source: unknown;
      constructor(family: string, source: unknown) {
        this.family = family;
        this.source = source;
        ctorArgs.push({ family, source });
      }
      load(): Promise<void> { return Promise.resolve(); }
    }
    const fakeDoc = { fonts: { add: (f: unknown) => added.push(f) } } as unknown as Document;
    const g = globalThis as unknown as { FontFace?: unknown };
    const prev = g.FontFace;
    g.FontFace = FakeFontFace as unknown;
    try {
      registerPreviewFonts(fakeDoc);
    } finally {
      if (prev === undefined) delete g.FontFace; else g.FontFace = prev;
    }

    expect(ctorArgs.length).toBeGreaterThan(0);
    for (const { source } of ctorArgs) {
      // Source is the raw bytes (a typed array), NOT a `url(...)` string → no fetch.
      expect(source instanceof Uint8Array || ArrayBuffer.isView(source as ArrayBufferView)).toBe(true);
      expect(typeof source).not.toBe('string');
    }
    // The multi-word clone families are registered (CSS consumers quote them).
    const families = new Set(ctorArgs.map((a) => a.family));
    for (const fam of CLONE_DISPLAY_FAMILIES) expect(families.has(fam)).toBe(true);
    expect([...families].some((f) => f.includes(' '))).toBe(true); // e.g. "Liberation Sans"
  });
});
