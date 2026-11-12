/**
 * Asserts out/web is what the single-file web build promises: ONE file, no
 * external references, a CSP that cannot reach the network, and no URL anywhere
 * in it that we have not looked at. Run at the end of `npm run build:web`.
 */
import { readdirSync, readFileSync, statSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { reporter } from './cdp.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const dir = path.join(root, 'out/web');
const r = reporter();

const entries = readdirSync(dir);
r.check('out/web contains exactly one file, Mutatis.html', entries.length === 1 && entries[0] === 'Mutatis.html', entries.join(', '));
const file = path.join(dir, 'Mutatis.html');
const html = readFileSync(file, 'utf8');

// ---- markup: strip script/style bodies, then look at what is left ----------
const markup = html
  .replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, (m) => m.slice(0, m.indexOf('>') + 1) + '</script>')
  .replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, '<style></style>');

const references = [...markup.matchAll(/\s(src|href|action|poster|data)\s*=\s*["']?([^"'\s>]+)/gi)].map((m) => m[2]);
const external = references.filter((value) => !/^(data:|blob:|#)/.test(value));
r.check('no src/href/action pointing at anything but data:, blob: or #', external.length === 0, external.slice(0, 3).join(', '));
r.check('no <link> or external <script src>', !/<link\b/i.test(markup) && !/<script\b[^>]*\bsrc=/i.test(markup));

// ---- CSP -------------------------------------------------------------------
const csp = /<meta[^>]+http-equiv=["']?Content-Security-Policy["']?[^>]*content="([^"]*)"/i.exec(markup)?.[1]
  ?? /<meta[^>]+content="([^"]*)"[^>]+http-equiv=["']?Content-Security-Policy/i.exec(markup)?.[1]
  ?? null;
r.check('a CSP <meta> is present', csp !== null);
if (csp !== null) {
  const directives = Object.fromEntries(
    csp.split(';').map((d) => d.trim()).filter(Boolean).map((d) => {
      const [name, ...sources] = d.split(/\s+/);
      return [name, sources];
    }),
  );
  r.check("default-src is 'none'", directives['default-src']?.join(' ') === "'none'");
  const reachesNetwork = Object.entries(directives).filter(([, sources]) =>
    sources.some((s) => /^(https?:|wss?:|\*|ftp:)/.test(s) || s === 'http:' || s === 'https:'),
  );
  r.check('no directive allows http:, https:, ws: or *', reachesNetwork.length === 0, reachesNetwork.map(([n]) => n).join(', '));
  r.check('connect-src allows data: only', directives['connect-src']?.join(' ') === 'data:', directives['connect-src']?.join(' '));
  r.check("object-src, base-uri and form-action are locked", directives['object-src']?.join() === "'none'" && directives['base-uri']?.join() === "'none'" && directives['form-action']?.join() === "'none'");
  r.check("no 'unsafe-eval' (only 'wasm-unsafe-eval')", !/'unsafe-eval'/.test(csp));
}

// ---- URLs in the bundle ----------------------------------------------------
// Everything below is a string constant in a library (XML/SVG namespaces, error
// message links, a licence comment) that is never requested. Anything NEW fails
// the build so it gets looked at rather than shipped.
const KNOWN = [
  /^https?:\/\/www\.w3\.org\//,
  /^https:\/\/react\.dev\/errors/,
  /^https:\/\/base-ui\.com\/production-error/,
  /^https:\/\/github\.com\/Hopding\/pdf-lib/,
  // OOXML namespace identifiers in the XML the paste feature writes. Names, not requests.
  /^http:\/\/schemas\.openxmlformats\.org\//,
  // PDF.js: XML namespace names for XFA forms, XMP metadata and XDP/XFDF, which it
  // matches against; placeholder domains in its own comments/defaults
  // (example.com, foo.bar); and the template in its link normaliser, which only
  // prefixes a scheme onto a string. None is requested — the CSP forbids it and
  // ui-check-web.mjs asserts zero network requests while reading PDFs.
  /^http:\/\/ns\.adobe\.com\//,
  /^http:\/\/www\.xfa\.org\/schema\//,
  /^http:\/\/example\.com$/,
  /^https:\/\/foo\.bar$/,
  /^http:\/\/\$\{[a-z]\}$/,
];
const urls = [...new Set([...html.matchAll(/https?:\/\/[^\s"'`)<>\\]+/g)].map((m) => m[0]))];
const unknown = urls.filter((u) => !KNOWN.some((k) => k.test(u)));
r.check('every http(s) URL in the file is a known inert string', unknown.length === 0, unknown.slice(0, 5).join(', '));
r.check('no absolute-URL fetch or XHR call', !/\b(fetch|open)\(\s*["'`]https?:/.test(html));
r.check('harfbuzz wasm is inlined, not addressed by URL', !/new URL\(["']harfbuzz\.wasm["']\s*,\s*import\.meta\.url\)/.test(html) && html.includes('data:application/wasm;base64,'));
r.check('worker is inlined', !/new Worker\(new URL\(/.test(html) && /new Worker\(/.test(html));

console.log(`\n  Mutatis.html: ${(statSync(file).size / 1024 / 1024).toFixed(2)} MB`);
process.exit(r.finish() === 0 ? 0 : 1);
