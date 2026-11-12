/**
 * MAINTAINER SMOKE CHECK. Uses a local folder of sample documents
 * (`redline-test-docs`, a sibling of this repository) that is not shipped with it;
 * the script skips itself when the folder is absent.
 *
 * Phase 6 drive-through. Runs TWO real comparisons in the packaged app, then
 * exercises the assistant's view of them: the context payload the model is
 * handed, the suggested prompts, and the on-rails deviation matrix.
 *
 * The on-rails path is the one that matters most here. The design requires the
 * matrix to be computed deterministically and only narrated, so this asserts
 * that pressing the button produces a matrix tab with real counted numbers
 * BEFORE any model is involved — and it runs with no API key configured, so no
 * model is involved at all and the matrix still appears.
 *
 * Throwaway verification, not part of the build.
 */
import { existsSync, mkdtempSync, readFileSync, readdirSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { connect, evaluate, findPackagedApp, launchPackaged, reporter, waitForPage } from './cdp.mjs';

const ROOT = path.resolve(import.meta.dirname, '..');
const PORT = 9335;
const { check, finish } = reporter();

const DOCS = path.resolve(ROOT, '../Mutatis/redline-test-docs');
const SETS = ['1-share-purchase-agreement', '3-services-agreement'];
for (const set of SETS) {
  if (!existsSync(path.join(DOCS, set, 'original.docx'))) {
    console.error(`Sample documents not found at ${DOCS}. Skipping.`);
    process.exit(0);
  }
}

const APP = findPackagedApp(ROOT);
if (!APP) {
  console.error('No packaged app found. Run `npm run dist` first.');
  process.exit(2);
}
console.log(`Driving packaged build: ${path.relative(ROOT, APP)}\n`);

const b64 = (p) => readFileSync(p).toString('base64');
const pairs = SETS.map((set) => ({
  set,
  original: b64(path.join(DOCS, set, 'original.docx')),
  revised: b64(path.join(DOCS, set, 'revised.docx')),
}));

const packaged = launchPackaged(APP, PORT);
// Downloads go to a throwaway directory. A check must never write into the
// user's Downloads folder just to prove a button works.
const downloads = mkdtempSync(path.join(tmpdir(), 'mutatis-downloads-'));
let failed = 0;

try {
  const page = await waitForPage(PORT);
  const cdp = await connect(page.webSocketDebuggerUrl);
  await cdp.send('Runtime.enable');
  await cdp.send('Page.enable');
  await cdp.send('Log.enable');
  await cdp.send('Network.enable');
  await cdp.send('Browser.setDownloadBehavior', {
    behavior: 'allow',
    downloadPath: downloads,
    eventsEnabled: true,
  });

  const cspViolations = [];
  const requested = [];
  cdp.on((message) => {
    if (message.method === 'Log.entryAdded') {
      const text = message.params.entry.text ?? '';
      if (/Content Security Policy|Refused to/i.test(text)) cspViolations.push(text);
    }
    if (message.method === 'Network.requestWillBeSent') {
      requested.push(message.params.request.url);
    }
  });

  await new Promise((resolve) => setTimeout(resolve, 1200));

  // A key must EXIST for the chat surface to render past its gate. It is a fake
  // one and it is never used: the matrix is computed without a model, and the
  // narration turn that follows fails for want of a real key — which is exactly
  // what proves the matrix does not depend on one.
  await evaluate(
    cdp,
    `(async () => window.mutatis.secrets.set('deepseek', 'sk-audit-not-a-real-key-0000'))()`,
  );
  await new Promise((resolve) => setTimeout(resolve, 400));
  await cdp.send('Page.enable');
  await cdp.send('Page.reload', {});
  await new Promise((resolve) => setTimeout(resolve, 2000));

  // All four documents at once: two per side. That is a many-to-many bundle, so
  // it routes through the pairing grid — which is what a user with two
  // agreements in front of them actually does, and it yields the two
  // comparisons the matrix needs in one pass.
  const injected = await evaluate(
    cdp,
    `(async () => {
       const toFile = (b64, name) => {
         const bin = atob(b64);
         const bytes = new Uint8Array(bin.length);
         for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
         return new File([bytes], name,
           { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
       };
       const inputs = document.querySelectorAll('input[type="file"]');
       if (inputs.length < 2) return 'found ' + inputs.length + ' file inputs';
       const put = (input, files) => {
         const dt = new DataTransfer();
         for (const file of files) dt.items.add(file);
         input.files = dt.files;
         input.dispatchEvent(new Event('change', { bubbles: true }));
       };
       put(inputs[0], [
         toFile(${JSON.stringify(pairs[0].original)}, '${pairs[0].set}-original.docx'),
         toFile(${JSON.stringify(pairs[1].original)}, '${pairs[1].set}-original.docx'),
       ]);
       put(inputs[1], [
         toFile(${JSON.stringify(pairs[0].revised)}, '${pairs[0].set}-revised.docx'),
         toFile(${JSON.stringify(pairs[1].revised)}, '${pairs[1].set}-revised.docx'),
       ]);
       await new Promise((r) => setTimeout(r, 600));
       const button = [...document.querySelectorAll('button')]
         .find((b) => b.textContent.trim().startsWith('Pair and compare'));
       if (!button) return 'no pair-and-compare button';
       if (button.disabled) return 'the button is disabled';
       button.click();
       return 'ok';
     })()`,
  );
  check('four documents injected and pairing started', injected === 'ok', String(injected));

  // Wait for the pairing grid, then confirm it.
  let confirmed = 'never appeared';
  for (let i = 0; i < 60; i += 1) {
    confirmed = await evaluate(
      cdp,
      `(() => {
         const button = [...document.querySelectorAll('button')]
           .find((b) => b.textContent.trim().startsWith('Confirm and compare'));
         if (!button || button.disabled) return 'waiting';
         button.click();
         return 'ok';
       })()`,
    );
    if (confirmed === 'ok') break;
    await new Promise((resolve) => setTimeout(resolve, 400));
  }
  check('the pairing grid appeared and was confirmed', confirmed === 'ok', confirmed);

  // Both comparisons now run.
  for (let i = 0; i < 80; i += 1) {
    const running = await evaluate(
      cdp,
      `[...document.querySelectorAll('[role="tab"]')].filter(t => t.textContent.includes('→')).length`,
    );
    const busy = await evaluate(cdp, `document.body.innerText.includes('Reading documents')`);
    if (running >= 2 && !busy) break;
    await new Promise((resolve) => setTimeout(resolve, 400));
  }
  await new Promise((resolve) => setTimeout(resolve, 1500));

  const tabTitles = await evaluate(
    cdp,
    `JSON.stringify([...document.querySelectorAll('[role="tab"]')].map(t => t.textContent.trim()))`,
  );
  const titles = JSON.parse(tabTitles);
  check('two comparisons ran in the packaged app',
    titles.filter((t) => t.includes('→')).length === 2, tabTitles);

  // ------------------------------------------------- the suggested prompts --
  const suggestions = await evaluate(
    cdp,
    `(async () => {
       // Open a redline tab so the chat has an active comparison.
       const tabs = [...document.querySelectorAll('[role="tab"]')]
         .filter(e => e.textContent.includes('→'));
       tabs[0]?.click();
       await new Promise((r) => setTimeout(r, 600));
       return JSON.stringify(
         [...document.querySelectorAll('button')]
           .map(b => ({ text: b.textContent.trim(), disabled: b.disabled }))
           .filter(b => /push back|matters|one-sided|Deviation matrix/i.test(b.text))
       );
     })()`,
  );
  const buttons = JSON.parse(suggestions);
  check('the chat offers suggested prompts once a comparison is open',
    buttons.length >= 3, suggestions.slice(0, 200));
  check('the deviation-matrix button is present and enabled with two comparisons',
    buttons.some((b) => /Deviation matrix/.test(b.text) && !b.disabled), suggestions.slice(0, 200));
  check('the on-rails button is labelled as computed',
    buttons.some((b) => /computed/.test(b.text)), suggestions.slice(0, 200));

  // ------------------------------------------------------------- the rails --
  // No API key is configured, so the narration turn will fail. The MATRIX must
  // still be computed and shown: it is engine output, not model output.
  const rails = await evaluate(
    cdp,
    `(async () => {
       [...document.querySelectorAll('button')]
         .find(b => /Deviation matrix/.test(b.textContent))?.click();
       await new Promise((r) => setTimeout(r, 1500));
       const table = document.querySelector('table');
       return JSON.stringify({
         heading: document.body.innerText.includes('Deviation matrix'),
         headers: table ? [...table.querySelectorAll('thead th')].map(th => th.textContent.trim()) : [],
         rows: table
           ? [...table.querySelectorAll('tbody tr')].map(tr =>
               [...tr.children].map(td => td.textContent.trim()))
           : [],
         notes: [...document.querySelectorAll('li')].map(li => li.textContent.trim())
                  .filter(t => t.length > 20),
       });
     })()`,
  );
  const matrix = JSON.parse(rails);
  check('pressing the on-rails button opens the matrix tab', matrix.heading === true);
  check('the matrix has a column per comparison plus Total and In',
    matrix.headers.length === 5, matrix.headers.join(' | '));
  check('the matrix has counted rows', matrix.rows.length > 0, `${matrix.rows.length} rows`);
  check('every cell is a number or an em dash — nothing is blank',
    matrix.rows.every((row) => row.slice(1).every((cell) => /^(\d+|—|\d+\/\d+)$/.test(cell))),
    JSON.stringify(matrix.rows.slice(0, 3)));
  check('the matrix states that its numbers are not from a language model',
    matrix.notes.some((note) => /does not come from a language model/.test(note)),
    matrix.notes.join(' | ').slice(0, 200));

  // The decisive one: no key exists, so no model ran. The matrix is there anyway.
  const keys = await evaluate(
    cdp,
    `(async () => JSON.stringify(await window.mutatis.secrets.status()))()`,
  );
  check('the matrix was produced without a usable model — the key stored is a fake',
    JSON.parse(keys).deepseek.mask === 'sk-…0000' && matrix.rows.length > 0, keys);

  // ------------------------------------------------- the redline is intact --
  const beside = await evaluate(
    cdp,
    `(async () => {
       const tabs = [...document.querySelectorAll('[role="tab"]')]
         .filter(e => e.textContent.includes('→'));
       tabs[0]?.click();
       await new Promise((r) => setTimeout(r, 600));
       return JSON.stringify({
         ins: document.querySelectorAll('ins').length,
         del: document.querySelectorAll('del').length,
         matrixTabStillOpen: [...document.querySelectorAll('[role="tab"]')]
           .some(t => t.textContent.includes('Deviation matrix')),
       });
     })()`,
  );
  const redline = JSON.parse(beside);
  check('the full redline is still there beside the assistant',
    redline.ins > 0 && redline.del > 0, beside);
  check('the matrix did not replace anything — it is a tab of its own',
    redline.matrixTabStillOpen === true, beside);

  // ------------------------------------------------------- the PDF redline --
  // The highest-risk thing in this phase, and only a PACKAGED run proves it: the
  // font bytes, the HarfBuzz wasm and the PDF writer all have to load over
  // app:// under the strict CSP. In dev they load over http from Vite, which
  // proves nothing.
  const pdfRun = await evaluate(
    cdp,
    `(async () => {
       const button = document.querySelector('[data-pdf-button]');
       if (!button) return 'no PDF button on the redline';
       button.click();
       for (let i = 0; i < 120; i++) {
         await new Promise(r => setTimeout(r, 250));
         const label = document.querySelector('[data-pdf-button]')?.textContent ?? '';
         const failure = [...document.querySelectorAll('span')]
           .map(s => s.textContent).find(t => t && t.startsWith('PDF failed:'));
         if (failure) return failure;
         if (!label.includes('Typesetting')) return 'done';
       }
       return 'timed out';
     })()`,
  );
  check('the PDF redline renders in the packaged app', pdfRun === 'done', String(pdfRun));

  // Give the download a moment to land, then look at the bytes.
  await new Promise((resolve) => setTimeout(resolve, 2500));
  const files = existsSync(downloads) ? readdirSync(downloads).filter((f) => f.endsWith('.pdf')) : [];
  check('a .pdf was written', files.length > 0, files.join(', ') || 'none');
  if (files.length > 0) {
    const bytes = readFileSync(path.join(downloads, files[0]));
    check('it is a real PDF, not an empty or truncated file',
      bytes.length > 1000 &&
        bytes.subarray(0, 5).toString('latin1') === '%PDF-' &&
        bytes.subarray(-1024).toString('latin1').includes('%%EOF'),
      `${bytes.length} bytes`);
  }
  // Resource Timing reports nothing for a non-HTTP scheme, so the DevTools
  // network log is the only place app:// fetches show up.
  const lazy = requested.filter((url) => /\.wasm$|faces-|\/pdf-/.test(url));
  check('HarfBuzz fetched its wasm over app://',
    lazy.some((url) => url.endsWith('.wasm')), lazy.join(', ') || 'none');
  check('the font faces and the PDF writer arrived as LAZY chunks',
    lazy.some((url) => url.includes('faces-')) && lazy.some((url) => /\/pdf-/.test(url)),
    lazy.join(', ') || 'none');
  check('every asset came from app://, none from the network',
    requested.every((url) => url.startsWith('app://') || url.startsWith('devtools://')),
    requested.filter((url) => !url.startsWith('app://')).join(', ') || 'all app://');

  const styles = await evaluate(cdp, 'document.querySelectorAll("style").length');
  check('no <style> element was injected', styles === 0, String(styles));
  check('no CSP violation during the AI path',
    cspViolations.length === 0, cspViolations.join(' | ') || 'none');

  cdp.close();
} catch (error) {
  check('the AI check completed', false, String(error));
} finally {
  failed = finish();
  packaged.dispose();
  rmSync(downloads, { recursive: true, force: true });
}

process.exit(failed > 0 ? 1 : 0);
