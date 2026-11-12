/**
 * PDF.js confidentiality hardening — the offline WALL.
 *
 * This module is PURE (no `pdfjs-dist` import) so it can be unit-tested without
 * a browser or the heavy PDF.js bundle. It owns two things:
 *
 *  1. {@link HARDENED_DOC_OPTIONS} — the fixed, security-critical `getDocument`
 *     flags every PDF open MUST carry. Document content never leaves the WebView
 *     (ADR-0001/0003); these flags eliminate PDF.js's only outbound surfaces.
 *  2. {@link assertSameOriginOrLocal} — a runtime guard the loader calls on every
 *     asset URL (worker / cMaps / standard fonts / wasm / icc) BEFORE handing it
 *     to PDF.js, throwing {@link OffOriginAssetError} on anything off-origin.
 *
 * Together with the build-time bundle scan and the Playwright zero-network spec
 * (see the A1 guard tests), this makes "PDF.js runs 100% offline" a wall, not a
 * convention — mirroring the three custom invariant lints.
 *
 * Note on `isEvalSupported`: PDF.js v5 REMOVED the eval code-path entirely
 * (the CVE-2024-4367 vector), so the version pin (≥ 5.1.91) already eliminates
 * it. We still pass the flag for defence-in-depth + intent; unknown flags are
 * ignored by `getDocument`.
 */

/** Thrown when a PDF.js asset URL would resolve to a non-local origin. */
export class OffOriginAssetError extends Error {
  constructor(
    readonly label: string,
    readonly url: string,
  ) {
    super(
      `PDF.js ${label} asset URL "${url}" is off-origin. ` +
        `All PDF.js assets MUST be bundled locally and served same-origin ` +
        `(ADR-0001/0003 confidentiality). Never point an asset URL at a CDN.`,
    );
    this.name = 'OffOriginAssetError';
  }
}

/**
 * The security-critical `getDocument` flags. Spread into every PDF open.
 *  - `useWasm:false`         — no WebAssembly image decoder ⇒ no `wasmUrl` fetch.
 *  - `useWorkerFetch:false`  — the worker never `fetch()`es cMaps/fonts itself;
 *                              the (guarded) main thread controls all asset I/O.
 *  - `useSystemFonts:false`  — never reach into the host's installed fonts.
 *  - `disableFontFace:true`  — text-only diffing never needs glyph rendering, so
 *                              skip @font-face entirely (no font network/CSS).
 *  - `isEvalSupported:false` — defence-in-depth (path removed upstream in v5).
 *  - `stopAtErrors:false`    — keep extracting good pages even if one is broken;
 *                              the per-page guards (A2) decide what is trustworthy.
 */
export const HARDENED_DOC_OPTIONS = {
  useWasm: false,
  useWorkerFetch: false,
  useSystemFonts: false,
  disableFontFace: true,
  isEvalSupported: false,
  stopAtErrors: false,
} as const;

/** Schemes that carry in-memory data (never a network request). Always allowed. */
const IN_MEMORY_SCHEMES = ['blob:', 'data:'];

/**
 * Guard one asset URL. Throws {@link OffOriginAssetError} unless the URL is:
 *  - an in-memory scheme (blob:/data:), or
 *  - a relative/root-relative path (resolves same-origin by definition), or
 *  - an absolute http(s) URL whose origin === the current page origin.
 *
 * Protocol-relative (`//host/…`) and any cross-origin absolute URL are rejected.
 * In a non-browser context (tests, where there is no `location`), absolute
 * http(s) URLs are rejected outright — production code never hardcodes one.
 */
export function assertSameOriginOrLocal(url: string, label: string): void {
  const u = url.trim();
  if (u === '') return; // empty ⇒ feature disabled (e.g. iccUrl unset) — no fetch.

  const lower = u.toLowerCase();
  if (IN_MEMORY_SCHEMES.some((s) => lower.startsWith(s))) return;

  // Protocol-relative: //host/path — inherits scheme but is an explicit host.
  if (u.startsWith('//')) throw new OffOriginAssetError(label, url);

  // Root-relative or explicitly relative ⇒ same-origin by construction.
  if (u.startsWith('/') || u.startsWith('./') || u.startsWith('../')) return;

  // Anything with an absolute scheme must be same-origin as the page.
  const hasScheme = /^[a-z][a-z0-9+.-]*:/i.test(u);
  if (!hasScheme) return; // bare relative path (e.g. "pdf-assets/cmaps/")

  const origin = currentOrigin();
  if (origin === null) {
    // No page origin (test/Node) — we never hardcode an absolute URL in prod.
    throw new OffOriginAssetError(label, url);
  }
  let parsed: URL;
  try {
    parsed = new URL(u);
  } catch {
    throw new OffOriginAssetError(label, url);
  }
  if (parsed.origin !== origin) throw new OffOriginAssetError(label, url);
}

/** Guard a whole set of asset URLs at once (keyed by label for clear errors). */
export function assertAssetUrlsLocal(urls: Record<string, string>): void {
  for (const [label, url] of Object.entries(urls)) {
    assertSameOriginOrLocal(url, label);
  }
}

/** The current page origin, or null when running outside a browser. */
function currentOrigin(): string | null {
  // Avoid a hard reference to `location` so this stays import-safe in Node.
  const loc = (globalThis as { location?: { origin?: string } }).location;
  return typeof loc?.origin === 'string' && loc.origin !== 'null' ? loc.origin : null;
}
