/**
 * MAINTAINER SMOKE CHECK. Uses a local folder of sample documents
 * (`redline-test-docs`, a sibling of this repository) that is not shipped with it;
 * the script skips itself when the folder is absent.
 *
 * Phase 2 UI check. Drives the packaged app the way a user would: inject two
 * real .docx files into the drop zones, click Compare, and assert a redline
 * actually rendered. Throwaway verification, not part of the build.
 */
import { app, BrowserWindow } from 'electron';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { readFileSync, writeFileSync, appendFileSync, existsSync } from 'node:fs';
import {
  registerAppScheme,
  serveRendererOverAppScheme,
  APP_ENTRY_URL,
} from '../out/main/app-protocol.js';

const root = path.dirname(fileURLToPath(import.meta.url));
const OUT = '/tmp/dp-ui-result.txt';
const say = (m) => appendFileSync(OUT, m + '\n');
writeFileSync(OUT, '');

const SET = process.env.DP_SET || '1-share-purchase-agreement';
const dir = path.resolve(root, '../../Mutatis/redline-test-docs', SET);
if (!existsSync(dir)) {
  say('SKIP: sample documents not found at ' + dir);
  process.exit(0);
}
const b64 = (p) => readFileSync(p).toString('base64');
const originalB64 = b64(path.join(dir, 'original.docx'));
const revisedB64 = b64(path.join(dir, 'revised.docx'));

const problems = [];
process.on('uncaughtException', (e) => { say('UNCAUGHT: ' + e.stack); app.exit(1); });

registerAppScheme();

app.whenReady().then(async () => {
  serveRendererOverAppScheme(path.join(root, '../out/renderer'));

  const win = new BrowserWindow({
    width: 1500, height: 950, show: false,
    webPreferences: {
      preload: path.join(root, '../out/preload/index.cjs'),
      contextIsolation: true, sandbox: true, nodeIntegration: false,
    },
  });

  win.webContents.on('console-message', (e) => {
    const m = typeof e === 'object' && e?.message ? e.message : String(e);
    say('  [renderer] ' + m);
    if (/Refused to|Content-Security-Policy|Insecure/i.test(m)) problems.push('CSP: ' + m);
  });

  await win.loadURL(APP_ENTRY_URL);
  await new Promise((r) => setTimeout(r, 1200));

  // Inject both documents through the real file inputs, exactly as the Browse
  // button does. React sees a genuine change event.
  const injected = await win.webContents.executeJavaScript(`(async () => {
    const toFile = (b64, name) => {
      const bin = atob(b64);
      const bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      return new File([bytes], name, {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });
    };
    const inputs = document.querySelectorAll('input[type="file"]');
    if (inputs.length < 2) return { ok: false, reason: 'found ' + inputs.length + ' file inputs' };

    const put = (input, file) => {
      const dt = new DataTransfer();
      dt.items.add(file);
      input.files = dt.files;
      input.dispatchEvent(new Event('change', { bubbles: true }));
    };
    put(inputs[0], toFile(${JSON.stringify(originalB64)}, 'original.docx'));
    put(inputs[1], toFile(${JSON.stringify(revisedB64)}, 'revised.docx'));
    return { ok: true };
  })()`);
  say('inject: ' + JSON.stringify(injected));
  if (!injected.ok) problems.push('could not inject files: ' + injected.reason);

  await new Promise((r) => setTimeout(r, 500));

  const clicked = await win.webContents.executeJavaScript(`(() => {
    const btn = Array.from(document.querySelectorAll('button'))
      .find((b) => b.textContent.trim().startsWith('Compare'));
    if (!btn) return { ok: false, reason: 'no Compare button' };
    if (btn.disabled) return { ok: false, reason: 'Compare is disabled' };
    btn.click();
    return { ok: true };
  })()`);
  say('click: ' + JSON.stringify(clicked));
  if (!clicked.ok) problems.push('could not start comparison: ' + clicked.reason);

  // Poll for the redline rather than guessing at a fixed delay.
  let state = null;
  for (let i = 0; i < 40; i++) {
    await new Promise((r) => setTimeout(r, 500));
    state = await win.webContents.executeJavaScript(`(() => {
      const article = document.querySelector('article.redline');
      const failed = document.body.innerText.includes('This comparison failed');
      return {
        rendered: !!article,
        failed,
        failureText: failed ? document.body.innerText.slice(0, 400) : null,
        tabs: document.querySelectorAll('[role="tab"]').length,
        insCount: document.querySelectorAll('ins.redline-ins, ins.redline-move-to').length,
        delCount: document.querySelectorAll('del.redline-del, del.redline-move-from').length,
        paragraphs: article ? article.querySelectorAll('p, h2').length : 0,
        summaryText: document.body.innerText.split('\\n').find((l) => / change/.test(l)) || null,
        injectedStyleTags: document.querySelectorAll('style').length,
      };
    })()`);
    if (state.rendered || state.failed) break;
  }

  say('=== UI RESULT (' + SET + ') ===');
  say(JSON.stringify(state, null, 2));

  if (!state?.rendered) problems.push('no redline rendered: ' + (state?.failureText ?? 'timed out'));
  if (state?.rendered) {
    if (state.insCount === 0) problems.push('redline shows no insertions');
    if (state.delCount === 0) problems.push('redline shows no deletions');
    if (state.paragraphs < 10) problems.push('suspiciously few paragraphs: ' + state.paragraphs);
    if (state.injectedStyleTags !== 0) {
      problems.push(state.injectedStyleTags + ' <style> element(s) injected');
    }
  }

  if (state?.rendered) {
    writeFileSync('/tmp/dp-redline.png', (await win.webContents.capturePage()).toPNG());
    say('screenshot: /tmp/dp-redline.png');

    // Scroll to the densest run of changes so the capture shows actual marks,
    // not just the title page.
    const scrolled = await win.webContents.executeJavaScript(`(() => {
      const marks = Array.from(document.querySelectorAll('ins.redline-ins, del.redline-del, ins.redline-move-to, del.redline-move-from'));
      if (marks.length === 0) return null;
      // Pick the mark with the most neighbours within 800px — the busiest region.
      let best = marks[0], bestCount = 0;
      for (const m of marks) {
        const y = m.getBoundingClientRect().top + window.scrollY;
        const count = marks.filter((o) => Math.abs((o.getBoundingClientRect().top + window.scrollY) - y) < 800).length;
        if (count > bestCount) { bestCount = count; best = m; }
      }
      const scroller = best.closest('.overflow-y-auto');
      best.scrollIntoView({ block: 'center' });
      return { neighbours: bestCount, text: best.textContent.slice(0, 60), scroller: !!scroller };
    })()`);
    say('scrolled to: ' + JSON.stringify(scrolled));
    await new Promise((r) => setTimeout(r, 600));
    writeFileSync('/tmp/dp-redline-marks.png', (await win.webContents.capturePage()).toPNG());
    say('screenshot: /tmp/dp-redline-marks.png');
  }

  say('=== PROBLEMS ===');
  say(problems.length ? problems.map((p) => ' x ' + p).join('\n') : ' none');
  app.exit(problems.length ? 1 : 0);
});
