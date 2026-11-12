/**
 * MAINTAINER SMOKE CHECK. Uses a local folder of sample documents
 * (`redline-test-docs`, a sibling of this repository) that is not shipped with it;
 * the script skips itself when the folder is absent.
 *
 * Phase 3 UI check. Drives a many-to-many bundle through the pairing grid:
 * four originals and four revisions from four DIFFERENT agreements, so the
 * correct pairing is unambiguous and a mis-pair would be obvious.
 */
import { app, BrowserWindow } from 'electron';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { readFileSync, writeFileSync, appendFileSync, existsSync } from 'node:fs';
import { registerAppScheme, serveRendererOverAppScheme, APP_ENTRY_URL } from '../out/main/app-protocol.js';

const root = path.dirname(fileURLToPath(import.meta.url));
const OUT = '/tmp/dp-pairing-result.txt';
const say = (m) => appendFileSync(OUT, m + '\n');
writeFileSync(OUT, '');

const SETS = ['1-share-purchase-agreement', '2-facility-loan-agreement', '3-services-agreement', '4-lease-england-and-singapore'];
const docsRoot = path.resolve(root, '../../Mutatis/redline-test-docs');
if (!existsSync(docsRoot)) { say('SKIP: no sample documents'); process.exit(0); }

const files = [];
for (const s of SETS) {
  files.push({ name: `${s}-original.docx`, side: 0, b64: readFileSync(path.join(docsRoot, s, 'original.docx')).toString('base64') });
  files.push({ name: `${s}-revised.docx`, side: 1, b64: readFileSync(path.join(docsRoot, s, 'revised.docx')).toString('base64') });
}

const problems = [];
process.on('uncaughtException', (e) => { say('UNCAUGHT: ' + e.stack); app.exit(1); });
registerAppScheme();

app.whenReady().then(async () => {
  serveRendererOverAppScheme(path.join(root, '../out/renderer'));
  const win = new BrowserWindow({
    width: 1500, height: 950, show: false,
    webPreferences: { preload: path.join(root, '../out/preload/index.cjs'), contextIsolation: true, sandbox: true, nodeIntegration: false },
  });
  win.webContents.on('console-message', (e) => {
    const m = typeof e === 'object' && e?.message ? e.message : String(e);
    say('  [renderer] ' + m);
    if (/Refused to|Content-Security-Policy/i.test(m)) problems.push('CSP: ' + m);
  });

  await win.loadURL(APP_ENTRY_URL);
  await new Promise((r) => setTimeout(r, 1200));

  const injected = await win.webContents.executeJavaScript(`(async () => {
    const FILES = ${JSON.stringify(files)};
    const toFile = (b64, name) => {
      const bin = atob(b64);
      const bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      return new File([bytes], name, { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    };
    const inputs = document.querySelectorAll('input[type="file"]');
    if (inputs.length < 2) return { ok: false, reason: 'inputs: ' + inputs.length };
    for (const side of [0, 1]) {
      const dt = new DataTransfer();
      for (const f of FILES.filter((f) => f.side === side)) dt.items.add(toFile(f.b64, f.name));
      inputs[side].files = dt.files;
      inputs[side].dispatchEvent(new Event('change', { bubbles: true }));
    }
    return { ok: true };
  })()`);
  say('inject: ' + JSON.stringify(injected));

  await new Promise((r) => setTimeout(r, 600));
  const clicked = await win.webContents.executeJavaScript(`(() => {
    const btn = Array.from(document.querySelectorAll('button')).find((b) => /Pair and compare|^Compare/.test(b.textContent.trim()));
    if (!btn) return { ok: false, reason: 'no button' };
    if (btn.disabled) return { ok: false, reason: 'disabled' };
    say = null; btn.click(); return { ok: true, label: btn.textContent.trim() };
  })()`);
  say('click: ' + JSON.stringify(clicked));
  if (!clicked.ok) problems.push('could not start: ' + clicked.reason);

  let grid = null;
  for (let i = 0; i < 60; i++) {
    await new Promise((r) => setTimeout(r, 500));
    grid = await win.webContents.executeJavaScript(`(() => {
      const heading = Array.from(document.querySelectorAll('h2')).find((h) => /Confirm how these documents pair/.test(h.textContent));
      if (!heading) return { ready: false, failed: document.body.innerText.includes('Pairing failed'), text: document.body.innerText.slice(0, 300) };
      const rows = Array.from(document.querySelectorAll('li')).filter((li) => li.querySelector('select'));
      return {
        ready: true,
        rowCount: rows.length,
        selects: document.querySelectorAll('select').length,
        statsLine: heading.parentElement.innerText.split('\\n').slice(0, 6),
        pairedLabels: rows.map((r) => r.innerText.split('\\n')[0]),
        injectedStyleTags: document.querySelectorAll('style').length,
      };
    })()`);
    if (grid.ready || grid.failed) break;
  }

  say('=== PAIRING GRID ===');
  say(JSON.stringify(grid, null, 2));
  if (!grid?.ready) problems.push('grid never appeared: ' + (grid?.text ?? 'timeout'));
  if (grid?.ready) {
    if (grid.rowCount !== 4) problems.push('expected 4 rows, got ' + grid.rowCount);
    if (grid.injectedStyleTags !== 0) problems.push('style tags: ' + grid.injectedStyleTags);
    writeFileSync('/tmp/dp-pairing.png', (await win.webContents.capturePage()).toPNG());
    say('screenshot: /tmp/dp-pairing.png');
  }

  say('=== PROBLEMS ===');
  say(problems.length ? problems.map((p) => ' x ' + p).join('\n') : ' none');
  app.exit(problems.length ? 1 : 0);
});
