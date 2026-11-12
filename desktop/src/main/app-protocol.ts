import { protocol, net, session } from 'electron';
import { pathToFileURL } from 'node:url';
import path from 'node:path';

/**
 * Serve the renderer from a custom `app://` scheme instead of `file://`.
 *
 * This is load-bearing and NOT optional. Three separate things break under
 * `file://`, and only a privileged custom scheme fixes all three:
 *
 *   a) Chromium blocks ESM module scripts on `file://` (opaque origin), so
 *      dynamic `import()` and the Vite bundle fail outright.
 *   b) `file://` has no response headers, so a real CSP is impossible — you are
 *      stuck with a `<meta>` tag, which cannot express `frame-ancestors` and is
 *      weaker in general.
 *   c) `localStorage` on a `file://` origin is unreliable, which breaks saved
 *      panel layouts.
 *
 * Setting Vite `base:'./'` is necessary but NOT sufficient. Do both.
 */

export const APP_SCHEME = 'app';

/**
 * The renderer's Content-Security-Policy, delivered as a real response header.
 *
 * `style-src 'self'` with no 'unsafe-inline' is achievable because we are on
 * Base UI, not Radix:
 *   - Base UI's <CSPProvider disableStyleElements> suppresses its only injected
 *     <style> element (the native-scrollbar-hiding rules used by
 *     ScrollArea.Viewport and Select's popup).
 *   - Base UI's remaining inline styling is set via CSSOM (element.style.*),
 *     and `style-src-attr` only polices style attributes parsed from HTML. A
 *     client-rendered app with no SSR never trips it.
 *   - Radix cannot do this: react-remove-scroll (Dialog, AlertDialog, Popover,
 *     Select, DropdownMenu, ContextMenu, Menubar) injects a <style> element and
 *     exposes no nonce prop.
 *
 * `connect-src 'self'` is deliberate and is a second wall behind the lint rule:
 * the renderer must never reach a provider directly. All provider HTTP happens
 * in the main process, which CSP does not govern.
 */
const CSP = [
  "default-src 'none'",
  // 'wasm-unsafe-eval' — NOT 'unsafe-eval'. It permits WebAssembly compilation
  // and nothing else: no eval(), no new Function(), no inline script, no remote
  // script. The PDF redline needs it because RenderCore shapes text with
  // HarfBuzz, which is wasm; without it WebAssembly.compile throws a CSP
  // violation and the PDF path fails at the last step.
  "script-src 'self' 'wasm-unsafe-eval'",
  "style-src 'self'",
  "img-src 'self' data: blob:",
  "font-src 'self' data:",
  "connect-src 'self'",
  "object-src 'none'",
  "base-uri 'none'",
  "form-action 'none'",
  "frame-ancestors 'none'",
].join('; ');

/**
 * Must run BEFORE app-ready. `standard: true` makes relative URLs resolve the
 * normal way; `secure: true` makes it a secure context (so crypto.subtle,
 * localStorage and service workers behave).
 */
export function registerAppScheme(): void {
  protocol.registerSchemesAsPrivileged([
    {
      scheme: APP_SCHEME,
      privileges: {
        standard: true,
        secure: true,
        supportFetchAPI: true,
        stream: true,
        codeCache: true,
      },
    },
  ]);
}

/** True when `child` is inside `parent` — blocks `..` traversal out of dist. */
function isInside(parent: string, child: string): boolean {
  const rel = path.relative(parent, child);
  return rel !== '' && !rel.startsWith('..') && !path.isAbsolute(rel);
}

/**
 * Wire `app://` to the built renderer directory, and attach the CSP header.
 * `rendererDist` is the absolute path to the built renderer output.
 */
export function serveRendererOverAppScheme(rendererDist: string): void {
  const root = path.resolve(rendererDist);

  protocol.handle(APP_SCHEME, async (request) => {
    const url = new URL(request.url);

    // `app://./index.html` parses with host '.', pathname '/index.html'.
    const decoded = decodeURIComponent(url.pathname);
    const candidate = path.resolve(root, '.' + decoded);

    // SPA fallback: a path with no file extension is a route, not an asset.
    const target =
      isInside(root, candidate) && path.extname(candidate) !== ''
        ? candidate
        : path.join(root, 'index.html');

    if (!isInside(root, target) && target !== path.join(root, 'index.html')) {
      return new Response('Forbidden', { status: 403 });
    }

    return net.fetch(pathToFileURL(target).toString());
  });

  // A real CSP header on every app:// response. `onHeadersReceived` is the only
  // way to set this for a scheme we serve ourselves.
  session.defaultSession.webRequest.onHeadersReceived((details, callback) => {
    callback({
      responseHeaders: {
        ...details.responseHeaders,
        'Content-Security-Policy': [CSP],
        'X-Content-Type-Options': ['nosniff'],
      },
    });
  });
}

export const APP_ENTRY_URL = `${APP_SCHEME}://./index.html`;
