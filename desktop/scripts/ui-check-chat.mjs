/**
 * Phase 4 UI check. Drives the PACKAGED app over the DevTools protocol with a
 * throwaway profile, and exercises the whole chat path end to end: the key gate,
 * safeStorage round-trip, SQLite thread persistence, and a full turn across the
 * MessagePort.
 *
 * The one thing it deliberately does NOT do is call a real provider. A turn is
 * started with no key configured, which drives renderer → preload → IPC → main →
 * stream → port → transport → useChat and comes back as the MissingKey error.
 * That proves every link in the chain without sending a legal test string to a
 * third party. The live provider round-trip is a maintainer test.
 *
 * Throwaway verification, not part of the build.
 */
import path from 'node:path';
import { connect, evaluate, findPackagedApp, launchPackaged, reporter, waitForPage } from './cdp.mjs';

const ROOT = path.resolve(import.meta.dirname, '..');
const PORT = 9334;
const { check, finish } = reporter();

const APP = findPackagedApp(ROOT);
if (!APP) {
  console.error('No packaged app found. Run `npm run dist` first.');
  process.exit(2);
}
console.log(`Driving packaged build: ${path.relative(ROOT, APP)}\n`);

const packaged = launchPackaged(APP, PORT);
let failed = 0;

try {
  const page = await waitForPage(PORT);
  const cdp = await connect(page.webSocketDebuggerUrl);
  await cdp.send('Runtime.enable');
  await cdp.send('Page.enable');
  await cdp.send('Log.enable');

  const cspViolations = [];
  cdp.on((message) => {
    if (message.method === 'Log.entryAdded') {
      const text = message.params.entry.text ?? '';
      if (/Content Security Policy|Refused to/i.test(text)) cspViolations.push(text);
    }
  });

  await new Promise((resolve) => setTimeout(resolve, 1200));

  // ------------------------------------------------------------- the gate --
  const gate = await evaluate(
    cdp,
    `JSON.stringify({
       heading: document.body.innerText.includes('Connect DeepSeek'),
       keyField: !!document.querySelector('input[type="password"]'),
       composer: !!document.querySelector('textarea[aria-label="Message the assistant"]'),
     })`,
  );
  const gateState = JSON.parse(gate);
  check('with no key stored, the chat panel shows the key gate', gateState.heading);
  check('the gate offers a password field, not a text field', gateState.keyField);
  check('the composer is not reachable before a key exists', gateState.composer === false);

  const encryption = await evaluate(
    cdp,
    `(async () => (await window.mutatis.providers()).encryptionAvailable)()`,
  );
  check('the OS keychain is available on this machine', encryption === true);

  // ------------------------------------------------------ safeStorage path --
  const FAKE = 'sk-mutatis-audit-0123456789abcdef';
  const saved = await evaluate(
    cdp,
    `(async () => JSON.stringify(await window.mutatis.secrets.set('deepseek', ${JSON.stringify(FAKE)})))()`,
  );
  const savedResult = JSON.parse(saved);
  check('a key can be stored through the bridge', savedResult.ok === true, saved);
  check('the confirmation is a mask, not the key',
    typeof savedResult.mask === 'string' && !savedResult.mask.includes('0123456789abcdef'),
    savedResult.mask);

  const status = await evaluate(
    cdp,
    `(async () => JSON.stringify(await window.mutatis.secrets.status()))()`,
  );
  check('status reports the key as configured', JSON.parse(status).deepseek.configured === true);
  check('status never returns the key itself', !status.includes('0123456789abcdef'), status);

  // The renderer must not be able to read the key back by any route.
  const leak = await evaluate(
    cdp,
    `JSON.stringify(Object.keys(window.mutatis.secrets).sort())`,
  );
  check('the secrets bridge has no getter at all', leak === '["set","status"]', leak);

  // ------------------------------------------------------------- the chat --
  // Page.reload, not `location.reload()` in an evaluated script: the latter is a
  // no-op when it comes from a CDP evaluation with no user gesture, which makes
  // for a test that silently checks the pre-reload page.
  await cdp.send('Page.reload', {});
  await new Promise((resolve) => setTimeout(resolve, 2500));

  const afterKey = await evaluate(
    cdp,
    `JSON.stringify({
       composer: !!document.querySelector('textarea[aria-label="Message the assistant"]'),
       threadPicker: !!document.querySelector('button[aria-label^="Conversations:"]'),
       threadListCollapsed:
         document.querySelector('button[aria-label^="Conversations:"]')?.getAttribute('aria-expanded'),
       modelPicker: !!document.querySelector('select[aria-label="Model"]'),
       models: [...document.querySelectorAll('select[aria-label="Model"] option')].map(o => o.value),
       groups: [...document.querySelectorAll('select[aria-label="Model"] optgroup')].map(o => o.label),
       gate: document.body.innerText.includes('Connect DeepSeek'),
     })`,
  );
  const chat = JSON.parse(afterKey);
  check('once a key exists the composer replaces the gate',
    chat.composer === true && chat.gate === false);
  // Phase 9 replaced the <select> picker with a disclosure in the rail header
  // that opens a list of buttons — an <option> may not hold the rename and
  // delete controls each row carries. So this asserts the disclosure and then
  // opens it, rather than a dropdown that no longer exists.
  check('the thread picker is present', chat.threadPicker === true);
  check('the thread list starts collapsed', chat.threadListCollapsed === 'false',
    String(chat.threadListCollapsed));

  const listOpened = await evaluate(
    cdp,
    `(async () => {
       document.querySelector('button[aria-label^="Conversations:"]').click();
       await new Promise((r) => setTimeout(r, 300));
       return JSON.stringify({
         expanded: document.querySelector('button[aria-label^="Conversations:"]')
           .getAttribute('aria-expanded'),
         newButton: !!document.querySelector('button[aria-label="New conversation"]'),
         // This profile is thrown away per run, so nothing is saved yet. The list
         // must SAY so — a blank rail and a broken rail look identical.
         emptyState: document.body.innerText.includes('No saved conversations yet'),
       });
     })()`,
  );
  const list = JSON.parse(listOpened);
  check('the picker opens the thread list', list.expanded === 'true' && list.newButton === true,
    listOpened);
  check('an empty list says so rather than rendering blank', list.emptyState === true, listOpened);
  check('the model picker spans every provider, keyed provider/model',
    chat.models.length === 5 && chat.models[0] === 'deepseek/deepseek-v4-flash',
    chat.models.join(', '));
  check('a provider with no key is listed and labelled as such',
    chat.groups.some((label) => /no key/.test(label)), chat.groups.join(' | '));

  // ---------------------------------------------- a real turn over the port --
  // No key for the model we ask for, so main refuses before any HTTP. Every
  // other link in the chain still runs for real.
  const turn = await evaluate(
    cdp,
    `(async () => {
       const chunks = [];
       await new Promise((resolve) => {
         const stop = window.mutatis.startChat(
           {
             messages: [{ id: 'audit-1', role: 'user', parts: [{ type: 'text', text: 'ping' }] }],
             model: { provider: 'alibaba', model: 'qwen-max' },
           },
           (chunk) => {
             chunks.push(chunk);
             if (chunk.type === 'error' || chunk.type === 'finish') resolve();
           },
         );
         setTimeout(() => { stop(); resolve(); }, 8000);
       });
       return JSON.stringify(chunks);
     })()`,
  );
  const chunks = JSON.parse(turn);
  check('a turn crosses the MessagePort and comes back', chunks.length > 0, turn);
  check('an unconfigured provider is refused in main, before any HTTP',
    chunks.some((chunk) => chunk.type === 'error' && /not wired up|No API key/i.test(chunk.errorText)),
    turn);
  check('the refusal carries no key material', !turn.includes('0123456789abcdef'));

  // A rejected payload must be refused, not forwarded.
  const bad = await evaluate(
    cdp,
    `(async () => {
       const chunks = [];
       await new Promise((resolve) => {
         window.mutatis.startChat(
           { messages: 'not an array', model: { provider: 'evil', model: 'x' } },
           (chunk) => { chunks.push(chunk); resolve(); },
         );
         setTimeout(resolve, 3000);
       });
       return JSON.stringify(chunks);
     })()`,
  );
  check('a malformed turn is rejected by main', /Rejected request/.test(bad), bad);

  // ------------------------------------------------------ SQLite threads ----
  const threads = await evaluate(
    cdp,
    `(async () => {
       const api = window.mutatis.threads;
       const id = crypto.randomUUID();
       await api.ensure(id, 'Audit thread');
       await api.save({ id: 'm1', threadId: id, role: 'user',
                        parts: [{ type: 'text', text: 'first' }], createdAt: 1 });
       await api.save({ id: 'm2', threadId: id, role: 'assistant',
                        parts: [{ type: 'text', text: 'second' }], createdAt: 2 });
       // Same id again: an upsert, so a streamed message that lands twice does
       // not duplicate the row.
       await api.save({ id: 'm2', threadId: id, role: 'assistant',
                        parts: [{ type: 'text', text: 'second (final)' }], createdAt: 2 });
       await api.rename(id, 'Renamed thread');
       const listed = await api.list();
       const loaded = await api.load(id);
       const rejected = await api.save({ id: 'm3', threadId: id, role: 'root', parts: [] });
       await api.remove(id);
       const after = await api.list();
       return JSON.stringify({
         titled: listed.find((t) => t.id === id)?.title,
         count: loaded.messages.length,
         skipped: loaded.skipped,
         lastText: loaded.messages[1]?.parts[0]?.text,
         rejected,
         goneAfterDelete: !after.some((t) => t.id === id),
       });
     })()`,
  );
  const store = JSON.parse(threads);
  check('a thread persists and can be renamed', store.titled === 'Renamed thread', threads);
  check('messages round-trip through SQLite in order',
    store.count === 2 && store.lastText === 'second (final)', threads);
  check('re-saving a message id upserts rather than duplicating', store.count === 2);
  check('no rows were unreadable', store.skipped === 0);
  check('a message with an invalid role is refused', store.rejected === false);
  check('deleting a thread removes it', store.goneAfterDelete === true);

  // ---------------------------------------------------------- the settings ----
  const opened = await evaluate(
    cdp,
    `(async () => {
       document.querySelector('button[aria-label="Settings"]').click();
       await new Promise((r) => setTimeout(r, 400));
       return JSON.stringify({
         title: [...document.querySelectorAll('h2,h3')].some(h => h.textContent === 'Settings')
                || document.body.innerText.includes('Jurisdiction'),
         jurisdiction: [...document.querySelectorAll('[role="radio"]')].map(b => b.textContent),
         keyRows: document.querySelectorAll('input[aria-label$="API key"]').length,
         modelOptions: [...document.querySelectorAll('select[aria-label="Assistant model"] option')]
                         .map(o => ({ value: o.value, disabled: o.disabled })),
         masks: [...document.querySelectorAll('span')].map(s => s.textContent)
                  .filter(t => /^sk-…/.test(t)),
       });
     })()`,
  );
  const settings = JSON.parse(opened);
  check('the settings dialog opens from the title bar', settings.title === true, opened);
  check('jurisdiction is a real choice, UK and Singapore',
    settings.jurisdiction.includes('United Kingdom') && settings.jurisdiction.includes('Singapore'),
    settings.jurisdiction.join(', '));
  check('every provider gets a key row', settings.keyRows === 2, String(settings.keyRows));
  check('the stored key is shown only as its mask',
    settings.masks.length === 1 && !settings.masks[0].includes('0123456789abcdef'),
    settings.masks.join(', '));
  check('models for a provider with no key are listed but not selectable',
    settings.modelOptions.filter((o) => o.disabled).length === 3 &&
      settings.modelOptions.filter((o) => !o.disabled).length === 2,
    JSON.stringify(settings.modelOptions));

  // Switching jurisdiction must survive a reload — it changes what the engine
  // reports, so a setting that silently reverts is worse than none.
  await evaluate(
    cdp,
    `(() => {
       [...document.querySelectorAll('[role="radio"]')].find(b => b.textContent === 'Singapore').click();
       document.querySelector('button[aria-label="Close settings"]').click();
     })()`,
  );
  await new Promise((resolve) => setTimeout(resolve, 300));
  await cdp.send('Page.reload', {});
  await new Promise((resolve) => setTimeout(resolve, 2500));
  const persisted = await evaluate(
    cdp,
    `(async () => {
       document.querySelector('button[aria-label="Settings"]').click();
       await new Promise((r) => setTimeout(r, 400));
       const on = [...document.querySelectorAll('[role="radio"]')]
         .filter(b => b.getAttribute('aria-checked') === 'true').map(b => b.textContent);
       document.querySelector('button[aria-label="Close settings"]').click();
       return JSON.stringify(on);
     })()`,
  );
  check('a jurisdiction change survives a reload',
    JSON.parse(persisted).includes('Singapore'), persisted);

  // ------------------------------------------------------------ the shell ----
  const styles = await evaluate(cdp, 'document.querySelectorAll("style").length');
  check('no <style> element was injected (style-src \'self\' holds)', styles === 0, String(styles));
  check('no CSP violation was logged during the chat path',
    cspViolations.length === 0, cspViolations.join(' | ') || 'none');

  cdp.close();
} catch (error) {
  check('the chat UI check completed', false, String(error));
} finally {
  failed = finish();
  packaged.dispose();
}

process.exit(failed > 0 ? 1 : 0);
