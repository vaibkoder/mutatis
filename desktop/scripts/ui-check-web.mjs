/**
 * MAINTAINER SMOKE CHECK. Uses a local folder of sample documents
 * (`redline-test-docs`, a sibling of this repository) that is not shipped with it;
 * the script skips itself when the folder is absent.
 *
 * Drives the single-file web build (out/web/Mutatis.html) in headless Chrome,
 * from file://, with the network switched off:
 *
 *   single compare -> redline + summary bar + context rail
 *   Save PDF redline -> a real PDF is produced
 *   batch compare -> pairing grid, repoint a row, confirm -> redlines open as tabs
 *   jurisdiction, theme and panel size persist across a reload
 *   zero requests off the page, zero console errors / CSP violations
 *
 * Chrome only. Safari and Firefox are NOT covered by this script.
 *
 *   npm run build:web && node scripts/ui-check-web.mjs
 */
import { spawn } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { connect, evaluate, reporter, waitForPage } from './cdp.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const HTML = process.env['WEB_HTML'] ?? path.join(root, 'out/web/Mutatis.html');
const CHROME =
  process.env['CHROME'] ?? '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
const DOCS = path.resolve(root, '../Mutatis/redline-test-docs');
const SETS = [
  '1-share-purchase-agreement',
  '2-facility-loan-agreement',
  '3-services-agreement',
  '4-lease-england-and-singapore',
];
const PORT = 9341;

if (!existsSync(HTML)) throw new Error('out/web/Mutatis.html is missing — run npm run build:web');
if (!existsSync(CHROME)) throw new Error(`Chrome not found at ${CHROME} (set CHROME=...)`);
if (!existsSync(DOCS)) {
  console.log('SKIP: no sample documents at ' + DOCS);
  process.exit(0);
}

const b64 = (set, which) => readFileSync(path.join(DOCS, set, `${which}.docx`)).toString('base64');
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const profile = mkdtempSync(path.join(tmpdir(), 'mutatis-web-'));
const chrome = spawn(
  CHROME,
  [
    '--headless=new',
    `--remote-debugging-port=${PORT}`,
    `--user-data-dir=${profile}`,
    '--no-first-run',
    '--window-size=1500,950',
    'about:blank',
  ],
  { stdio: 'ignore' },
);

const r = reporter();
const offPage = [];
const consoleProblems = [];

try {
  const page = await waitForPage(PORT);
  const cdp = await connect(page.webSocketDebuggerUrl);
  await cdp.send('Runtime.enable');
  await cdp.send('Log.enable');
  await cdp.send('Network.enable');
  await cdp.send('Page.enable');
  // No network at all, so anything that depended on it fails instead of passing quietly.
  await cdp.send('Network.emulateNetworkConditions', {
    offline: true,
    latency: 0,
    downloadThroughput: 0,
    uploadThroughput: 0,
  });
  // Capture every Blob that is turned into an object URL (that is how savePdf saves).
  await cdp.send('Page.addScriptToEvaluateOnNewDocument', {
    source: `(() => {
      window.__blobs = [];
      const orig = URL.createObjectURL.bind(URL);
      URL.createObjectURL = (b) => { window.__blobs.push(b); return orig(b); };
    })();`,
  });

  cdp.on((m) => {
    if (m.method === 'Network.requestWillBeSent') {
      const url = m.params.request.url;
      if (!/^(file|data|blob):/.test(url)) offPage.push(url);
    }
    if (m.method === 'Runtime.exceptionThrown') {
      consoleProblems.push('exception: ' + (m.params.exceptionDetails.exception?.description ?? m.params.exceptionDetails.text));
    }
    if (m.method === 'Runtime.consoleAPICalled' && m.params.type === 'error') {
      consoleProblems.push('console.error: ' + m.params.args.map((a) => a.value ?? a.description).join(' '));
    }
    if (m.method === 'Log.entryAdded' && ['error', 'warning'].includes(m.params.entry.level)) {
      consoleProblems.push(`${m.params.entry.level}: ${m.params.entry.text} ${m.params.entry.url ?? ''}`);
    }
  });

  const load = async () => {
    await cdp.send('Page.navigate', { url: pathToFileURL(HTML).href });
    for (let i = 0; i < 60; i++) {
      await sleep(250);
      if (await evaluate(cdp, `document.querySelectorAll('input[type="file"]').length >= 2`)) return;
    }
    throw new Error('app did not mount');
  };

  const addDocs = (sets) =>
    evaluate(
      cdp,
      `(() => {
        const toFile = (b64, name) => {
          const bin = atob(b64);
          const bytes = new Uint8Array(bin.length);
          for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
          return new File([bytes], name, { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
        };
        const FILES = ${JSON.stringify(
          sets.flatMap((s) => [
            { side: 0, name: `${s}-original.docx`, data: b64(s, 'original') },
            { side: 1, name: `${s}-revised.docx`, data: b64(s, 'revised') },
          ]),
        )};
        const inputs = document.querySelectorAll('input[type="file"]');
        for (const side of [0, 1]) {
          const dt = new DataTransfer();
          for (const f of FILES.filter((f) => f.side === side)) dt.items.add(toFile(f.data, f.name));
          inputs[side].files = dt.files;
          inputs[side].dispatchEvent(new Event('change', { bubbles: true }));
        }
        return true;
      })()`,
    );

  const clickButton = (pattern) =>
    evaluate(
      cdp,
      `(() => {
        const b = [...document.querySelectorAll('button')].find((b) => ${pattern}.test(b.textContent.trim()) && !b.disabled);
        if (!b) return false;
        b.click();
        return true;
      })()`,
    );

  const waitFor = async (expression, ms = 40000) => {
    const end = Date.now() + ms;
    while (Date.now() < end) {
      const v = await evaluate(cdp, expression);
      if (v) return v;
      await sleep(300);
    }
    return null;
  };

  // ---- 1. single compare --------------------------------------------------
  console.log('single compare');
  await load();
  r.check('app mounts from file://', true);
  r.check(
    'no chat or AI UI',
    await evaluate(
      cdp,
      `!/Assistant|chat|API key|deviation/i.test(document.body.innerText) && !('mutatis' in window)`,
    ),
  );
  await addDocs([SETS[0]]);
  await sleep(300);
  r.check('Compare starts', await clickButton('/^Compare$/'));
  const summary = await waitFor(
    `(() => { const m = document.body.innerText.match(/(\\d+) changes?/); return m ? m[0] : null; })()`,
  );
  r.check('redline shows a change summary', summary !== null, String(summary));
  r.check(
    'redline document renders',
    await evaluate(cdp, `document.body.innerText.length > 2000`),
  );
  r.check(
    'Tailwind styling reached the shared components (not just the web shell)',
    await evaluate(
      cdp,
      `(() => {
        const del = document.querySelector('del');
        const bar = [...document.querySelectorAll('div')].find((d) => /^\\d+ changes?/.test(d.innerText) && d.className.includes('gap-x-3'));
        return !!del && getComputedStyle(del).textDecorationLine.includes('line-through')
          && !!bar && getComputedStyle(bar).columnGap !== 'normal' && getComputedStyle(bar).display === 'flex';
      })()`,
    ),
  );
  r.check(
    'context rail lists the documents',
    await evaluate(cdp, `document.body.innerText.includes('${SETS[0]}-original.docx')`),
  );

  // ---- 2. PDF -------------------------------------------------------------
  console.log('save PDF');
  r.check('Save PDF redline clicked', await clickButton('/Save PDF redline/'));
  const pdf = await waitFor(
    `(async () => {
      const b = (window.__blobs || []).find((b) => b.type === 'application/pdf');
      if (!b) return null;
      const head = new TextDecoder().decode(new Uint8Array(await b.slice(0, 5).arrayBuffer()));
      const buf = new Uint8Array(await b.arrayBuffer());
      let bin = '';
      for (let i = 0; i < buf.length; i += 0x8000) bin += String.fromCharCode(...buf.subarray(i, i + 0x8000));
      return { head, size: b.size, b64: btoa(bin) };
    })()`,
    90000,
  );
  if (pdf && process.env['PDF_OUT']) writeFileSync(process.env['PDF_OUT'], Buffer.from(pdf.b64, 'base64'));
  r.check('a PDF blob is produced', pdf !== null && pdf.head === '%PDF-', pdf ? `${pdf.size} bytes` : 'none');
  r.check(
    'PDF button reports no error',
    await evaluate(cdp, `!document.body.innerText.includes('PDF failed')`),
  );

  // ---- 3. persistence -----------------------------------------------------
  console.log('persistence');
  await evaluate(
    cdp,
    `document.querySelector('button[aria-label="Settings"]').click()`,
  );
  await sleep(400);
  r.check(
    'settings has Jurisdiction + Appearance only',
    await evaluate(
      cdp,
      `(() => { const t = document.querySelector('[role=dialog]').innerText; return /jurisdiction/i.test(t) && /appearance/i.test(t) && !/api keys|updates|model/i.test(t); })()`,
    ),
  );
  await evaluate(
    cdp,
    `[...document.querySelectorAll('[role=radio]')].find((b) => b.textContent.trim() === 'Singapore').click()`,
  );
  await evaluate(cdp, `document.querySelector('button[aria-label="Dark"]').click()`);
  await sleep(300);
  // The dialog is modal: close it before touching the panels.
  await cdp.send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Escape', code: 'Escape', windowsVirtualKeyCode: 27 });
  await cdp.send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Escape', code: 'Escape', windowsVirtualKeyCode: 27 });
  await sleep(500);
  // Drag the separator to change the panel layout.
  const sep = await evaluate(
    cdp,
    `(() => { const s = document.querySelector('[role=separator]'); const b = s.getBoundingClientRect(); return { x: b.x + b.width / 2, y: b.y + b.height / 2 }; })()`,
  );
  await cdp.send('Input.dispatchMouseEvent', { type: 'mouseMoved', x: sep.x, y: sep.y });
  await cdp.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: sep.x, y: sep.y, button: 'left', clickCount: 1 });
  for (let i = 1; i <= 8; i++) {
    await cdp.send('Input.dispatchMouseEvent', { type: 'mouseMoved', x: sep.x - i * 20, y: sep.y, buttons: 1 });
  }
  await cdp.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: sep.x - 160, y: sep.y, button: 'left', clickCount: 1 });
  await sleep(500);
  const layoutBefore = await evaluate(cdp, `localStorage.getItem('mutatis.web.shell.v1') ?? localStorage.getItem(Object.keys(localStorage).find((k) => k.includes('mutatis.web.shell')) ?? '')`);
    r.check('panel layout is saved', typeof layoutBefore === 'string', String(layoutBefore).slice(0, 80));
  await load();
  r.check('theme persists (dark)', await evaluate(cdp, `document.documentElement.getAttribute('data-theme') === 'dark'`));
  await evaluate(cdp, `document.querySelector('button[aria-label="Settings"]').click()`);
  await sleep(400);
  r.check(
    'jurisdiction persists (Singapore)',
    await evaluate(
      cdp,
      `[...document.querySelectorAll('[role=radio]')].find((b) => b.textContent.trim() === 'Singapore').getAttribute('aria-checked') === 'true'`,
    ),
  );
  r.check(
    'panel layout persists',
    (await evaluate(cdp, `localStorage.getItem(Object.keys(localStorage).find((k) => k.includes('mutatis.web.shell')) ?? '')`)) === layoutBefore,
  );
  await cdp.send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Escape', code: 'Escape' });
  await cdp.send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Escape', code: 'Escape' });
  await sleep(300);

  // ---- 4. batch compare ---------------------------------------------------
  console.log('batch compare');
  await load();
  await addDocs(SETS);
  await sleep(300);
  r.check('Pair and compare starts', await clickButton('/Pair and compare/'));
  const grid = await waitFor(
    `(() => document.body.innerText.includes('Confirm how these documents pair up') ? document.querySelectorAll('select').length : 0)()`,
  );
  r.check('pairing grid appears', grid === 4, `${grid} rows`);
  const repointed = await evaluate(
    cdp,
    `(() => {
      const sel = document.querySelectorAll('select')[0];
      const before = sel.value;
      const other = [...sel.options].find((o) => o.value && o.value !== before);
      const setter = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set;
      setter.call(sel, '');
      sel.dispatchEvent(new Event('change', { bubbles: true }));
      return { before, other: other?.value ?? null };
    })()`,
  );
  await sleep(500);
  r.check(
    'repointing a row to "Leave unmatched" updates the grid',
    await evaluate(cdp, `document.body.innerText.includes('no counterpart')`),
    JSON.stringify(repointed).slice(0, 80),
  );
  await evaluate(
    cdp,
    `(() => { const sel = document.querySelectorAll('select')[0]; const setter = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set; setter.call(sel, ${JSON.stringify(repointed.before)}); sel.dispatchEvent(new Event('change', { bubbles: true })); })()`,
  );
  await sleep(500);
  r.check('Confirm and compare', await clickButton('/Confirm and compare/'));
  const tabs = await waitFor(
    `(() => { const n = document.querySelectorAll('[role=tab]').length; return n >= 5 ? n : 0; })()`,
  );
  r.check('confirming opens the redlines as tabs', tabs >= 5, `${tabs} tabs`);

  // ---- 5. pasted text -----------------------------------------------------
  console.log('paste compare');
  await load();
  const paste = (html, text) =>
    evaluate(
      cdp,
      `(() => {
        const ta = document.querySelectorAll('textarea')[0];
        if (!ta) return false;
        ta.focus();
        const dt = new DataTransfer();
        if (${JSON.stringify(html)}) dt.setData('text/html', ${JSON.stringify(html)});
        dt.setData('text/plain', ${JSON.stringify(text)});
        ta.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
        return true;
      })()`,
    );
  const switchWidths = () =>
    evaluate(
      cdp,
      `(() => { const [a, b] = [...document.querySelectorAll('[aria-label="What to compare"] button')].map((x) => x.getBoundingClientRect()); return { files: a.width, paste: b.width, height: Math.min(a.height, b.height) }; })()`,
    );
  const before = await switchWidths();
  r.check('the files / pasted-text switch is large and labelled', before.height >= 48 && before.paste >= 250, JSON.stringify(before));
  r.check('files mode: Compare files takes about 3/4 of the row', before.files / before.paste > 2.5 && before.files / before.paste < 3.5, (before.files / before.paste).toFixed(2));
  r.check('switch to pasted text', await clickButton('/Compare pasted text/'));
  await sleep(600);
  const after = await switchWidths();
  r.check('pasted mode: Compare pasted text takes about 3/4 of the row', after.paste / after.files > 2.5 && after.paste / after.files < 3.5, (after.paste / after.files).toFixed(2));
  r.check('Compare is disabled until both sides have text', !(await evaluate(cdp, `[...document.querySelectorAll('button')].filter((b) => /Compare pasted text/.test(b.textContent)).pop().disabled === false`)));
  await paste(
    '<html><body><!--StartFragment--><h1>Fees</h1><p>The <b>Buyer</b> pays <i>5</i> days.</p><p style="mso-list:l0"><span style="mso-list:Ignore">1.&nbsp;</span>One</p><!--EndFragment--></body></html>',
    'Fees\nThe Buyer pays 5 days.\n1. One',
  );
  await sleep(300);
  r.check(
    'Word-style HTML is kept as formatting (bold in the preview)',
    await evaluate(cdp, `document.body.innerText.includes('Formatting kept') && !!document.querySelector('.font-bold') && !!document.querySelector('.italic')`),
  );
  await paste('', '# Fees\n\nThe **Buyer** pays *7* days.\n\n1. One');
  await sleep(300);
  r.check('Markdown text is detected and read as Markdown', await evaluate(cdp, `document.body.innerText.includes('Read as Markdown')`));
  r.check('Compare pasted text runs', await evaluate(cdp, `(() => { const b = [...document.querySelectorAll('button')].filter((b) => /Compare pasted text/.test(b.textContent) && !b.disabled).pop(); if (!b) return false; b.click(); return true; })()`));
  const pastedSummary = await waitFor(`(() => { const m = document.body.innerText.match(/(\\d+) changes?/); return m ? m[0] : null; })()`);
  r.check('a redline of the pasted texts appears', pastedSummary !== null, String(pastedSummary));
  await clickButton('/New comparison/');
  await sleep(300);
  r.check('pasted text survives a look at the result and back', await evaluate(cdp, `document.querySelectorAll('textarea').length + document.body.innerText.split('Formatting kept').length - 1 >= 1`));

  // ---- 5b. uploaded .md / .txt ---------------------------------------------
  console.log('uploaded text files');
  await load();
  await evaluate(
    cdp,
    `(() => {
      const inputs = document.querySelectorAll('input[type="file"]');
      const put = (i, name, body) => {
        const dt = new DataTransfer();
        dt.items.add(new File([body], name, { type: 'text/plain' }));
        inputs[i].files = dt.files;
        inputs[i].dispatchEvent(new Event('change', { bubbles: true }));
      };
      put(0, 'old.md', '# Fees\\n\\nThe **Buyer** pays 5 days.');
      put(1, 'new.txt', 'Fees\\n\\nThe Buyer pays 7 days.');
    })()`,
  );
  await sleep(500);
  r.check('Compare starts for a .md against a .txt', await clickButton('/^Compare$/'));
  const textSummary = await waitFor(`(() => { const m = document.body.innerText.match(/(\\d+) changes?/); return m ? m[0] : null; })()`);
  r.check('uploaded .md and .txt compare (they used to fail as "not a zip file")', textSummary !== null, String(textSummary));

  // ---- 5c. uploaded PDFs ---------------------------------------------------
  console.log('uploaded PDFs');
  const PDFS = path.join(root, 'tests/fixtures/pdf');
  const pdfB64 = (name) => readFileSync(path.join(PDFS, name)).toString('base64');
  const putFile = (input, name, b64, type) =>
    evaluate(
      cdp,
      `(() => {
        const bin = atob(${JSON.stringify(b64)});
        const bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
        const dt = new DataTransfer();
        dt.items.add(new File([bytes], ${JSON.stringify(name)}, { type: ${JSON.stringify(type)} }));
        const el = document.querySelectorAll('input[type="file"]')[${input}];
        el.files = dt.files;
        el.dispatchEvent(new Event('change', { bubbles: true }));
      })()`,
    );
  await load();
  await putFile(0, 'scanned.pdf', pdfB64('scanned.pdf'), 'application/pdf');
  const refused = await waitFor(`(() => document.body.innerText.includes('was not added') || document.body.innerText.includes('were not added') ? document.body.innerText : null)()`, 30000);
  r.check('a scanned PDF is refused at upload, with the reason', refused !== null && /scanned or image-only/.test(refused), refused ? refused.slice(refused.indexOf('scanned.pdf'), refused.indexOf('scanned.pdf') + 140) : 'no message');
  await putFile(0, 'contract-a.pdf', pdfB64('contract-a.pdf'), 'application/pdf');
  await putFile(1, 'contract-b.pdf', pdfB64('contract-b.pdf'), 'application/pdf');
  const added = await waitFor(`(() => document.body.innerText.includes('contract-a.pdf') && document.body.innerText.includes('contract-b.pdf') ? true : null)()`, 60000);
  r.check('two PDFs are read and added', added === true);
  r.check('Compare starts for PDF against PDF', await clickButton('/^Compare$/'));
  const pdfSummary = await waitFor(`(() => { const m = document.body.innerText.match(/(\\d+) changes?/); return m ? m[0] : null; })()`, 60000);
  r.check('a redline of the two PDFs appears', pdfSummary !== null, String(pdfSummary));
  r.check(
    'the redline says the text was rebuilt from PDF layout',
    await evaluate(cdp, `(() => { const n = document.querySelector('[role=note]'); return !!n && /Read from a PDF/.test(n.innerText) && /Original \\(contract-a\\.pdf\\)/.test(n.innerText); })()`),
  );
  r.check(
    'the PDF edits are in the redline (S$5,000 -> S$6,000, 30 -> 60 days)',
    await evaluate(cdp, `(() => { const t = document.body.innerText; return t.includes('S$6,000') && t.includes('S$5,000') && /60/.test(t); })()`),
  );
  r.check('Save PDF redline works on a PDF-sourced comparison', await clickButton('/Save PDF redline/'));
  const pdfOut = await waitFor(`(async () => { const bs = (window.__blobs || []).filter((b) => b.type === 'application/pdf'); return bs.length ? bs.length : null; })()`, 90000);
  r.check('and produces a PDF', pdfOut !== null);
  r.check('Copy for AI button is shown', await evaluate(cdp, `!!document.querySelector('[data-ai-export]') && /Copy for AI/.test(document.querySelector('[data-ai-export]').innerText)`));
  r.check('the .md export button is clicked', await clickButton('/^\\.md$/'));
  const mdOut = await waitFor(`(async () => { const bs = (window.__blobs || []).filter((b) => b.type === 'text/markdown'); if (!bs.length) return null; return await bs[bs.length - 1].text(); })()`, 10000);
  r.check('the .md holds the change set (S$6,000 and 60)', mdOut !== null && /# Mutatis change set/.test(mdOut) && mdOut.includes('S$6,000') && /60/.test(mdOut), String(mdOut).slice(0, 80));
  r.check('the Structure button opens the structure panel beside the redline', await clickButton('/^Structure$/'));
  const panel = await waitFor(`(() => { const p = document.querySelector('[data-structure-panel]'); return p && p.querySelectorAll('li[data-kind]').length > 0 ? p.querySelectorAll('li[data-kind]').length : null; })()`, 10000);
  r.check('the panel lists the document structure', panel !== null, String(panel));
  r.check('the redline stays on screen with the panel open', await evaluate(cdp, `document.body.innerText.includes('S$6,000')`));
  await evaluate(cdp, `(() => { const b = [...document.querySelectorAll('[data-structure-panel] button')].find((x) => x.innerText.trim() === '.md'); b && b.click(); return !!b; })()`);
  const structOut = await waitFor(`(async () => { const bs = (window.__blobs || []).filter((b) => b.type === 'text/markdown'); if (bs.length < 2) return null; return await bs[bs.length - 1].text(); })()`, 10000);
  r.check('the structure saves as a Markdown outline', structOut !== null && /# Contract structure/.test(structOut), String(structOut).slice(0, 80));

  // ---- 5d. one document with tracked changes ------------------------------
  console.log('single document with tracked changes');
  await load();
  const trackedB64 = readFileSync(path.join(root, 'tests/fixtures/tracked.docx')).toString('base64');
  await putFile(0, 'tracked.docx', trackedB64, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  r.check(
    'a lone document says its tracked changes will be compared',
    (await waitFor(`document.body.innerText.includes('its tracked changes will be compared') ? true : null`, 20000)) === true,
  );
  r.check('Compare starts with only one document', await clickButton('/^Compare$/'));
  const trackedSummary = await waitFor(`(() => { const m = document.body.innerText.match(/(\\d+) changes?/); return m ? m[0] : null; })()`, 60000);
  r.check('the tracked changes appear as a redline', trackedSummary !== null, String(trackedSummary));
  r.check(
    'the redline shows the tracked edits and explains where they came from',
    await evaluate(cdp, `(() => { const t = document.body.innerText; const n = document.querySelector('[role=note]'); return t.includes('£55,000') && t.includes('£40,000') && !!n && /tracked changes/.test(n.innerText); })()`),
  );

  // ---- 6. hygiene ---------------------------------------------------------
  console.log('hygiene');
  r.check('zero network requests off the page', offPage.length === 0, offPage.slice(0, 3).join(', '));
  r.check('no console errors or CSP violations', consoleProblems.length === 0, consoleProblems.slice(0, 3).join(' | '));

  cdp.close();
} catch (error) {
  r.check('script ran to completion', false, String(error?.stack ?? error));
} finally {
  chrome.kill('SIGTERM');
  await new Promise((resolve) => {
    chrome.once('exit', resolve);
    setTimeout(resolve, 3000);
  });
  rmSync(profile, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 });
}

process.exit(r.finish() === 0 ? 0 : 1);
