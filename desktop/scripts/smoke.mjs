/**
 * The fast smoke check. `npm run smoke`, seconds rather than minutes.
 *
 * Boots the built renderer over app:// and asserts that the page painted, that
 * app:// is a secure context, that the preload bridge arrived under
 * sandbox:true, that the CSP landed as a real response header, and that the
 * three-panel layout is structurally intact.
 *
 * SCOPE, stated because it was previously overclaimed. This script builds its
 * OWN minimal main process — it deliberately does not run `registerIpc()`,
 * because that lives inside the bundled `out/main/index.js`, which self-starts
 * the whole application on import. So THE IPC LAYER IS NOT EXERCISED HERE. It
 * is covered by `npm run audit:egress`, which drives the real packaged binary
 * with its real main process.
 *
 * The consequence used to be a silent lie: the renderer's start-up calls failed
 * with "No handler registered", those errors were logged, and the pass/fail
 * filter ignored them — so the check reported "none" while booting into a
 * broken state, and would not have caught a genuinely missing handler either.
 * Now the expected set is pinned below, which turns that noise into a signal:
 * a NEW channel added without a handler changes the set and fails here.
 *
 * NOTE: no top-level await anywhere. In an ESM Electron main script the `ready`
 * event is not dispatched until the module graph finishes evaluating, so a
 * top-level `await app.whenReady()` deadlocks. Use the .then() form — which is
 * what src/main/index.ts does.
 */
import { app, BrowserWindow } from 'electron';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { writeFileSync, appendFileSync } from 'node:fs';
import {
  registerAppScheme,
  serveRendererOverAppScheme,
  APP_ENTRY_URL,
} from '../out/main/app-protocol.js';

const root = path.dirname(fileURLToPath(import.meta.url));
const OUT = '/tmp/dp-smoke-result.txt';
// Both places on purpose. The findings used to go only to the file while the
// exit code went to the shell, so a failure could be read as a pass — which is
// exactly how a background run of this script reported success over an error.
const say = (m) => {
  appendFileSync(OUT, m + '\n');
  process.stdout.write(m + '\n');
};
writeFileSync(OUT, '');

/**
 * The IPC channels the renderer may call on start-up, which this script cannot
 * answer (see the scope note above).
 *
 * An UNEXPECTED name here fails: it means a channel was added to the renderer's
 * start-up path with no handler registered in main.
 *
 * A MISSING name does NOT fail, and that is a deliberate retreat from the first
 * version of this check. `secrets:status` is only reached if `providers:list`
 * resolves first, so with no main IPC at all the observed set depends on how the
 * settings provider short-circuits — it varies between runs. An assertion that
 * flakes is worse than no assertion, so this is a subset check, not equality.
 * Whether the handlers actually ANSWER is `npm run audit:egress`'s job.
 */
const ALLOWED_UNHANDLED = ['providers:list', 'secrets:status', 'threads:list'];
const unhandled = new Set();

const problems = [];
process.on('uncaughtException', (e) => { say('UNCAUGHT: ' + e.stack); app.exit(1); });
process.on('unhandledRejection', (e) => { say('UNHANDLED: ' + (e?.stack ?? e)); app.exit(1); });

registerAppScheme();

app.whenReady().then(async () => {
  serveRendererOverAppScheme(path.join(root, '../out/renderer'));

  const win = new BrowserWindow({
    show: false,
    webPreferences: {
      preload: path.join(root, '../out/preload/index.cjs'),
      contextIsolation: true,
      sandbox: true,
      nodeIntegration: false,
    },
  });

  win.webContents.on('console-message', (e) => {
    const message = typeof e === 'object' && e?.message ? e.message : String(e);
    say('  [renderer] ' + message);
    const noHandler = /No handler registered for '([^']+)'/.exec(message);
    if (noHandler?.[1]) {
      unhandled.add(noHandler[1]);
      return;
    }
    if (/Content-Security-Policy|Refused to|Insecure/i.test(message)) {
      problems.push('CSP/console: ' + message);
    }
  });
  win.webContents.on('did-fail-load', (_e, code, desc, url) => {
    problems.push(`did-fail-load ${code} ${desc} ${url}`);
  });

  await win.loadURL(APP_ENTRY_URL);
  await new Promise((r) => setTimeout(r, 1500));

  const probe = await win.webContents.executeJavaScript(`(() => ({
    title: document.title,
    rootText: document.getElementById('root')?.innerText ?? '(no #root)',
    childCount: document.getElementById('root')?.childElementCount ?? 0,
    origin: location.origin,
    isSecureContext,
    hasBridge: typeof window.mutatis === 'object',
    platform: window.mutatis?.platform ?? null,
    injectedStyleTags: document.querySelectorAll('style').length,
    stylesheets: document.styleSheets.length,
    panels: document.querySelectorAll('[data-panel]').length,
    separators: document.querySelectorAll('[data-separator]').length,
    // v4 marks the Group root with data-group. NOT data-panel-group — that is
    // the v3 attribute, renamed in the same pass that turned
    // data-panel-group-direction into aria-orientation.
    separatorsAreDirectGroupChildren: Array.from(
      document.querySelectorAll('[data-separator]'),
    ).every((sep) => sep.parentElement?.hasAttribute('data-group')),
    dropZones: document.querySelectorAll('section').length,
    themeRadios: document.querySelectorAll('[role="radio"]').length,
    bodyBg: getComputedStyle(document.body).backgroundColor,
  }))()`);

  const cspSeen = await win.webContents.executeJavaScript(
    `fetch('./index.html').then(r => r.headers.get('content-security-policy')).catch(e => 'FETCH_ERR: ' + e.message)`,
  );

  say('=== SMOKE RESULT ===');
  say(JSON.stringify(probe, null, 2));
  say('CSP header on app:// response:');
  say('  ' + cspSeen);

  if (probe.childCount === 0) problems.push('renderer painted nothing (#root empty)');
  if (probe.injectedStyleTags !== 0) {
    problems.push(
      `${probe.injectedStyleTags} <style> element(s) injected — style-src 'self' forbids these`,
    );
  }
  if (probe.panels !== 3) problems.push(`expected 3 panels, found ${probe.panels}`);
  if (probe.separators !== 2) problems.push(`expected 2 separators, found ${probe.separators}`);
  if (!probe.separatorsAreDirectGroupChildren) {
    problems.push('a Separator is not a direct child of its Group — layout will break silently');
  }
  if (probe.themeRadios !== 3) problems.push('theme toggle missing');
  if (!probe.isSecureContext) problems.push('app:// is NOT a secure context');
  if (!probe.hasBridge) problems.push('preload bridge missing under sandbox:true');
  if (!cspSeen || String(cspSeen).startsWith('FETCH_ERR')) problems.push('no CSP header on app://');

  const seen = [...unhandled].sort();
  say('IPC channels called but unanswered (expected — this script runs no main IPC):');
  say('  ' + (seen.join(', ') || 'none'));
  const unexpected = seen.filter((channel) => !ALLOWED_UNHANDLED.includes(channel));
  if (unexpected.length > 0) {
    problems.push(
      `start-up IPC call with no handler in main: ${unexpected.join(', ')}\n` +
        `      Either register it in src/main/ipc.ts or add it to ALLOWED_UNHANDLED\n` +
        `      here with a reason.`,
    );
  }

  say('=== PROBLEMS ===');
  say(problems.length ? problems.map((p) => ' x ' + p).join('\n') : ' none');
  app.exit(problems.length ? 1 : 0);
});
