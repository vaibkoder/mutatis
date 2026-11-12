/**
 * THE EGRESS AUDIT — phase 4.
 *
 * Two non-negotiables:
 *   1. API keys never appear in the renderer, in preload, or in any VITE_* var.
 *   2. All provider HTTP happens in the Electron main process.
 *
 * Both are claims about the SHIPPED artefact, so this runs against the packaged
 * app — `release/**​/Mutatis Desktop.app` — and never against the dev server.
 * Dev-server behaviour is not proof: the dev renderer is served over http with a
 * different module graph and a different CSP path.
 *
 * What may leave the machine is stated plainly in the app: an AI chat turn the user
 * sent, which carries the file names and changed wording of the open comparison.
 * This audit checks the paths by which anything else could leave.
 *
 * Part A is static: unpack the asar and read what actually shipped.
 * Part B is live: launch the packaged binary, attach over the DevTools protocol,
 * and interrogate the real renderer — its globals, its CSP, and every network
 * request it is able to make.
 *
 * Run:  node scripts/egress-audit.mjs
 */
import { execFileSync } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, readdirSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import {
  connect,
  evaluate,
  findPackagedApp,
  launchPackaged,
  waitForPage,
} from './cdp.mjs';

const ROOT = path.resolve(import.meta.dirname, '..');
const results = [];
let failures = 0;

function check(name, ok, detail = '') {
  results.push({ name, ok, detail });
  if (!ok) failures += 1;
  const mark = ok ? 'PASS' : 'FAIL';
  console.log(`  ${mark}  ${name}${detail ? ` — ${detail}` : ''}`);
}

const APP = findPackagedApp(ROOT);
if (!APP) {
  console.error('No packaged app found. Run `npm run dist` first.');
  process.exit(2);
}
console.log(`Auditing packaged build: ${path.relative(ROOT, APP)}\n`);

// ---------------------------------------------------------------- Part A ----

console.log('Part A — static audit of the shipped asar');

const work = mkdtempSync(path.join(tmpdir(), 'mutatis-audit-'));
// @electron/asar ships with electron-builder, so this resolves locally rather
// than fetching a package mid-audit.
execFileSync(
  process.execPath,
  [
    path.join(ROOT, 'node_modules/@electron/asar/bin/asar.js'),
    'extract',
    path.join(APP, 'Contents/Resources/app.asar'),
    work,
  ],
  { cwd: ROOT, stdio: 'inherit' },
);

function readAll(dir) {
  const out = [];
  const walk = (current) => {
    for (const entry of readdirSync(current, { withFileTypes: true })) {
      const full = path.join(current, entry.name);
      if (entry.isDirectory()) walk(full);
      else if (/\.(js|cjs|mjs|html|css)$/.test(entry.name)) {
        out.push({ file: path.relative(work, full), text: readFileSync(full, 'utf8') });
      }
    }
  };
  walk(dir);
  return out;
}

const rendererFiles = readAll(path.join(work, 'out/renderer'));
const preloadFiles = readAll(path.join(work, 'out/preload'));
const mainFiles = readAll(path.join(work, 'out/main'));

const rendererText = rendererFiles.map((f) => f.text).join('\n');
const preloadText = preloadFiles.map((f) => f.text).join('\n');
const mainText = mainFiles.map((f) => f.text).join('\n');

check('the renderer bundle actually shipped', rendererFiles.length > 0,
  `${rendererFiles.length} files`);

// 1. No provider endpoint reachable from the renderer.
const PROVIDER_HOSTS = [
  'api.deepseek.com',
  'dashscope.aliyuncs.com',
  'dashscope-intl.aliyuncs.com',
  'api.openai.com',
  'api.anthropic.com',
];
for (const host of PROVIDER_HOSTS) {
  check(`renderer contains no "${host}"`, !rendererText.includes(host));
  check(`preload contains no "${host}"`, !preloadText.includes(host));
}

// 2. The absence above would be worthless if the provider code had simply
//    vanished from the build. It must still be present, reachable ONLY from
//    main. `externalizeDepsPlugin` keeps it in node_modules rather than inlining
//    it, so this looks there and then confirms main is what imports it.
function sdkText(name) {
  const dir = path.join(work, 'node_modules', name);
  return existsSync(dir) ? readAll(dir).map((f) => f.text).join('\n') : '';
}

for (const [pkg, host] of [
  ['@ai-sdk/deepseek', 'api.deepseek.com'],
  ['@ai-sdk/alibaba', 'dashscope-intl.aliyuncs.com'],
]) {
  check(`${pkg} shipped (the call moved, it did not vanish)`, sdkText(pkg).includes(host));
  check(`main is what imports ${pkg}`, mainText.includes(pkg));
  check(`the renderer bundle never imports ${pkg}`, !rendererText.includes(pkg));
  check(`preload never imports ${pkg}`, !preloadText.includes(pkg));
}

// 3. No key-shaped literal or key-named env read anywhere in renderer/preload.
const KEY_PATTERNS = [
  [/\bsk-[A-Za-z0-9_-]{12,}/, 'an sk- API key literal'],
  [/DEEPSEEK_API_KEY/, 'DEEPSEEK_API_KEY'],
  [/DASHSCOPE_API_KEY/, 'DASHSCOPE_API_KEY'],
  [/ALIBABA_API_KEY/, 'ALIBABA_API_KEY'],
  [/VITE_[A-Z0-9_]*(KEY|SECRET|TOKEN|PASSWORD)/, 'a VITE_*KEY/SECRET/TOKEN variable'],
];
for (const [pattern, label] of KEY_PATTERNS) {
  check(`renderer contains no ${label}`, !pattern.test(rendererText));
  check(`preload contains no ${label}`, !pattern.test(preloadText));
}

// 4. safeStorage is a main-process API. Its presence anywhere else would mean
//    key handling had leaked across the boundary.
check('renderer never references safeStorage', !rendererText.includes('safeStorage'));
check('preload never references safeStorage', !preloadText.includes('safeStorage'));
check('main DOES use safeStorage', mainText.includes('safeStorage'));

// 5. The preload must expose a named API, not ipcRenderer itself.
check('preload does not expose ipcRenderer through the bridge',
  !/exposeInMainWorld\([^)]*ipcRenderer\s*\)/.test(preloadText));
check('preload exposes exactly one bridge object',
  (preloadText.match(/exposeInMainWorld/g) ?? []).length === 1);

// 6. The CSP that the audit's Part B will verify at runtime must be the one in
//    the shipped main bundle.
check("main ships connect-src 'self'", mainText.includes("connect-src 'self'"));
check("main ships default-src 'none'", mainText.includes("default-src 'none'"));
// 'wasm-unsafe-eval' is permitted (HarfBuzz); plain 'unsafe-eval' is not, and
// neither is 'unsafe-inline' anywhere.
check("script-src does NOT allow plain 'unsafe-eval'",
  !/script-src[^;"]*'unsafe-eval'/.test(mainText));
check("no directive allows 'unsafe-inline'", !mainText.includes("'unsafe-inline'"));

// 7. Nothing in the package should be able to turn the guards off.
// Phase 8. The updater is the one piece of this app with a standing reason to
// call home, so the "nothing leaves unless the user asked" claim has to be
// re-proved against it — statically here, and live below.
check('nothing calls checkForUpdatesAndNotify at startup',
  !mainText.includes('checkForUpdatesAndNotify'));
check('the updater never downloads on its own', /autoDownload\s*=\s*false/.test(mainText));
check('the updater never installs on quit', /autoInstallOnAppQuit\s*=\s*false/.test(mainText));
check('electron-updater is only ever imported lazily',
  /\bimport\(\s*['"]electron-updater['"]\s*\)/.test(mainText)
    && !/^import[^;\n]*['"]electron-updater['"]/m.test(mainText));
// No publish config means electron-builder writes no app-update.yml, so the
// shipped app holds no feed URL at all — not a disabled one, none.
check('no update feed URL is present in this build',
  !existsSync(path.join(APP, 'Contents/Resources/app-update.yml')),
  'app-update.yml absent');

check('webSecurity is never disabled', !/webSecurity\s*:\s*false/.test(mainText));
check('nodeIntegration is never enabled', !/nodeIntegration\s*:\s*true/.test(mainText));
check('contextIsolation is never disabled', !/contextIsolation\s*:\s*false/.test(mainText));
check('sandbox is never disabled', !/sandbox\s*:\s*false/.test(mainText));

// ---------------------------------------------------------------- Part B ----

console.log('\nPart B — live audit of the running packaged app');

const PORT = 9333;
const packaged = launchPackaged(APP, PORT);

let exitCode = 0;
try {
  const page = await waitForPage(PORT);
  const cdp = await connect(page.webSocketDebuggerUrl);

  await cdp.send('Runtime.enable');
  await cdp.send('Network.enable');
  await cdp.send('Log.enable');

  const requests = [];
  const cspViolations = [];
  cdp.on((message) => {
    if (message.method === 'Network.requestWillBeSent') {
      requests.push(message.params.request.url);
    }
    if (message.method === 'Log.entryAdded') {
      const entry = message.params.entry;
      if (entry.source === 'security' || /Content Security Policy/i.test(entry.text ?? '')) {
        cspViolations.push(entry.text);
      }
    }
  });

  const origin = await evaluate(cdp, 'location.origin');
  check('the packaged renderer runs on the app:// scheme, not file://',
    origin.startsWith('app://'), origin);

  // Node must be unreachable from the page.
  const nodeGlobals = await evaluate(
    cdp,
    `JSON.stringify(['require','process','module','global','Buffer','__dirname']
       .filter(k => typeof window[k] !== 'undefined'))`,
  );
  check('no Node globals are reachable from the page', nodeGlobals === '[]', nodeGlobals);

  const hasIpc = await evaluate(
    cdp,
    `typeof window.ipcRenderer !== 'undefined'
       || typeof window.electron !== 'undefined'
       || typeof (window.mutatis && window.mutatis.ipcRenderer) !== 'undefined'`,
  );
  check('ipcRenderer is not exposed to the page', hasIpc === false);

  // The bridge is a closed list. A new key here is a decision, not an accident.
  const EXPECTED_BRIDGE = ['openKeyPage', 'platform', 'providers', 'secrets', 'startChat', 'threads', 'updates'];
  const bridge = await evaluate(cdp, 'JSON.stringify(Object.keys(window.mutatis).sort())');
  check('the preload bridge exposes exactly the audited surface',
    bridge === JSON.stringify(EXPECTED_BRIDGE), bridge);

  const secretsKeys = await evaluate(cdp, 'JSON.stringify(Object.keys(window.mutatis.secrets).sort())');
  check('the secrets bridge can set and report, but never read a key back',
    secretsKeys === JSON.stringify(['set', 'status']), secretsKeys);

  // The CSP must arrive as a real response header on the document itself.
  const csp = await evaluate(
    cdp,
    `(async () => {
       const response = await fetch(location.href);
       return response.headers.get('content-security-policy') || '';
     })()`,
  );
  check("the document carries a Content-Security-Policy header", csp.length > 0);
  check("CSP restricts connect-src to 'self'", csp.includes("connect-src 'self'"), csp);
  check("CSP default-src is 'none'", csp.includes("default-src 'none'"));
  check("the live CSP allows wasm but not eval",
    csp.includes("'wasm-unsafe-eval'") && !/script-src[^;]*\s'unsafe-eval'/.test(csp), csp);
  check('the live CSP allows no inline anything', !csp.includes("'unsafe-inline'"));

  // Proof, not inference: compile a minimal module and try to eval.
  const sandbox = await evaluate(
    cdp,
    `(async () => {
       let wasm = 'blocked';
       try {
         await WebAssembly.compile(new Uint8Array([0,97,115,109,1,0,0,0]));
         wasm = 'ok';
       } catch { /* blocked */ }
       let evalled = 'blocked';
       try { (0, eval)('1+1'); evalled = 'ALLOWED'; } catch { /* blocked */ }
       return JSON.stringify({ wasm, evalled });
     })()`,
  );
  const sandboxState = JSON.parse(sandbox);
  check('WebAssembly compiles (the PDF redline needs HarfBuzz)',
    sandboxState.wasm === 'ok', sandbox);
  check('eval() is still refused', sandboxState.evalled === 'blocked', sandbox);

  // Phase 8, live. Pressing Check on a build that cannot update must produce a
  // reason and NO request. The `no external network request` assertion below
  // runs after this, so a beacon fired here would be caught there.
  const updates = await evaluate(
    cdp,
    `(async () => {
       const before = await window.mutatis.updates.state();
       const after = await window.mutatis.updates.check();
       return JSON.stringify({ before: before.phase, after: after.phase, why: after.blockers });
     })()`,
  );
  const updateState = JSON.parse(updates);
  check('an unsigned, feedless build reports itself as unable to update',
    updateState.before === 'unsupported' && updateState.after === 'unsupported', updates);
  check('and it says why, in words rather than by going quiet',
    Array.isArray(updateState.why) && updateState.why.length > 0,
    (updateState.why ?? []).join(' | '));

  // The decisive test: the renderer tries to reach a provider directly, and the
  // browser must refuse before a packet leaves the machine.
  const blocked = await evaluate(
    cdp,
    `(async () => {
       try {
         await fetch('https://api.deepseek.com/v1/models', { method: 'GET' });
         return 'REQUEST WAS ALLOWED';
       } catch (error) {
         return 'blocked: ' + String(error && error.message);
       }
     })()`,
  );
  check('a direct provider fetch from the renderer is refused',
    blocked.startsWith('blocked:'), blocked);

  // Reading a PDF is the one path that loads a third-party worker, and it must do
  // it from inside the app. Read two real PDFs in the PACKAGED renderer, compare
  // them and save the redline; the no-external-request assertion below then covers
  // everything that happened.
  const pdfDir = path.join(ROOT, 'tests/fixtures/pdf');
  const pdfFile = (index, name) => `(() => {
       const bin = atob(${JSON.stringify(readFileSync(path.join(pdfDir, name)).toString('base64'))});
       const bytes = new Uint8Array(bin.length);
       for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
       const dt = new DataTransfer();
       dt.items.add(new File([bytes], ${JSON.stringify(name)}, { type: 'application/pdf' }));
       const el = document.querySelectorAll('input[type="file"]')[${index}];
       el.files = dt.files;
       el.dispatchEvent(new Event('change', { bubbles: true }));
     })()`;
  const until = async (expression, ms = 60000) => {
    const end = Date.now() + ms;
    while (Date.now() < end) {
      const value = await evaluate(cdp, expression);
      if (value) return value;
      await new Promise((resolve) => setTimeout(resolve, 300));
    }
    return null;
  };
  await evaluate(cdp, pdfFile(0, 'contract-a.pdf'));
  await evaluate(cdp, pdfFile(1, 'contract-b.pdf'));
  const pdfsAdded = await until(`document.body.innerText.includes('contract-a.pdf') && document.body.innerText.includes('contract-b.pdf')`);
  check('two PDFs are read in the packaged app', pdfsAdded === true);
  await evaluate(cdp, `[...document.querySelectorAll('button')].find((b) => /^Compare$/.test(b.textContent.trim()) && !b.disabled)?.click()`);
  const pdfSummary = await until(`(() => { const m = document.body.innerText.match(/(\\d+) changes?/); return m ? m[0] : null; })()`);
  check('and compared, with the reconstruction notice shown', pdfSummary !== null && (await evaluate(cdp, `/Read from a PDF/.test(document.querySelector('[role=note]')?.innerText ?? '')`)), String(pdfSummary));
  const pdfAssets = requests.filter((url) => url.includes('pdf-assets'));
  check('the PDF worker was loaded from the app itself', pdfAssets.length > 0 && pdfAssets.every((url) => /^(app|file):/.test(url)), pdfAssets[0] ?? 'no worker request seen');

  await new Promise((resolve) => setTimeout(resolve, 500));

  const external = requests.filter((url) => !/^(app|devtools|data|blob):/.test(url));
  check('the renderer issued no external network request',
    external.length === 0, external.join(', ') || 'none');

  check('the CSP reported the blocked provider request',
    cspViolations.some((text) => /deepseek/i.test(text)),
    cspViolations.join(' | ') || 'no violation logged');

  cdp.close();
} catch (error) {
  check('live audit completed', false, String(error));
  exitCode = 1;
} finally {
  packaged.dispose();
  rmSync(work, { recursive: true, force: true });
}

// ------------------------------------------------------------------ done ----

console.log(`\n${results.length - failures}/${results.length} checks passed.`);
if (failures > 0) {
  console.log('\nFAILED:');
  for (const r of results.filter((entry) => !entry.ok)) {
    console.log(`  - ${r.name}${r.detail ? ` — ${r.detail}` : ''}`);
  }
}
process.exit(failures > 0 ? 1 : exitCode);
