/**
 * A1 confidentiality guard — LAYER 1 (static).
 *
 * Asserts the hardened PDF.js options are set and the asset-URL guard rejects
 * every off-origin shape while accepting local/same-origin ones. Also scans the
 * loader/config source for any hardcoded off-origin CDN string — a literal CDN
 * URL must never reach production. The runtime zero-network proof (LAYER 2) is
 * the Playwright spec `tests/integration/pdf-network.spec.ts` (user-run) plus the
 * build-bundle scan in `scripts/check-pdf-offline.mjs`.
 */
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

import {
  HARDENED_DOC_OPTIONS,
  assertSameOriginOrLocal,
  assertAssetUrlsLocal,
  OffOriginAssetError,
} from '../../../src/features/compare/io/pdf/pdfjs-config.js';

const HERE = dirname(fileURLToPath(import.meta.url));
const SRC = resolve(HERE, '../../../src/features/compare/io/pdf');

describe('A1 — PDF.js offline hardening (static guard)', () => {
  it('hardened getDocument options disable every outbound surface', () => {
    expect(HARDENED_DOC_OPTIONS.useWasm).toBe(false);
    expect(HARDENED_DOC_OPTIONS.useWorkerFetch).toBe(false);
    expect(HARDENED_DOC_OPTIONS.useSystemFonts).toBe(false);
    expect(HARDENED_DOC_OPTIONS.disableFontFace).toBe(true);
    expect(HARDENED_DOC_OPTIONS.isEvalSupported).toBe(false);
  });

  it('rejects off-origin asset URLs', () => {
    const offOrigin = [
      'https://cdn.jsdelivr.net/npm/pdfjs-dist/cmaps/',
      'http://unpkg.com/pdfjs-dist/standard_fonts/',
      'https://mozilla.github.io/pdf.js/web/',
      '//evil.example.com/pdf.worker.js',
      'https://fonts.googleapis.com/x',
    ];
    for (const url of offOrigin) {
      expect(() => assertSameOriginOrLocal(url, 'test'), url).toThrow(OffOriginAssetError);
    }
  });

  it('accepts local / in-memory / root-relative asset URLs', () => {
    const local = [
      '/pdf-assets/cmaps/',
      '/pdf-assets/pdf.worker.min.mjs',
      './pdf-assets/standard_fonts/',
      'pdf-assets/wasm/',
      'blob:https://localhost/abc-123',
      '', // unset (feature disabled) — no fetch
    ];
    for (const url of local) {
      expect(() => assertSameOriginOrLocal(url, 'test'), url).not.toThrow();
    }
  });

  it('treats absolute http(s) URLs as off-origin when no page origin (Node)', () => {
    // In tests there is no `location`; production never hardcodes an absolute URL.
    expect(() => assertSameOriginOrLocal('https://localhost:3000/pdf-assets/cmaps/', 'x')).toThrow(
      OffOriginAssetError,
    );
  });

  it('assertAssetUrlsLocal throws on the first off-origin member', () => {
    expect(() =>
      assertAssetUrlsLocal({
        cMapUrl: '/pdf-assets/cmaps/',
        workerSrc: 'https://cdn.jsdelivr.net/x',
      }),
    ).toThrow(OffOriginAssetError);
  });

  it('loader + config source contain no hardcoded off-origin CDN string', () => {
    const sources = ['pdfjs-loader.ts', 'pdfjs-config.ts'].map((f) =>
      readFileSync(resolve(SRC, f), 'utf8'),
    );
    const cdnHosts = [
      'jsdelivr',
      'unpkg.com',
      'cdnjs',
      'mozilla.github.io',
      'googleapis.com',
      'gstatic.com',
    ];
    for (const src of sources) {
      for (const host of cdnHosts) {
        expect(src.includes(host), `source must not reference ${host}`).toBe(false);
      }
      // No absolute http(s) literal at all in the loader/config.
      expect(/https?:\/\/[a-z0-9.-]+/i.test(stripComments(src))).toBe(false);
    }
  });
});

/** Remove // and block comments so example URLs in docs don't trip the scan. */
function stripComments(src: string): string {
  return src.replace(/\/\*[\s\S]*?\*\//g, '').replace(/\/\/.*$/gm, '');
}
