/**
 * A minimal Chrome DevTools Protocol client, used to drive the PACKAGED app.
 *
 * Deliberately dependency-free and tiny: the protocol is a JSON envelope with
 * three fields, and adding puppeteer to interrogate our own renderer would put a
 * browser automation stack in the dependency tree of a legal tool.
 *
 * Shared by scripts/egress-audit.mjs and the phase UI checks so that they
 * interrogate the real shipped binary rather than a re-implementation of it.
 */
import { spawn } from 'node:child_process';
import { existsSync, mkdtempSync, readdirSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';

export function findPackagedApp(root) {
  const release = path.join(root, 'release');
  if (!existsSync(release)) return null;
  for (const dir of readdirSync(release)) {
    const app = path.join(release, dir, 'Mutatis Desktop.app');
    if (existsSync(app)) return app;
  }
  return null;
}

/**
 * Launch the packaged binary with a THROWAWAY profile. The audits write API keys
 * and chat threads; they must never touch the user's real key store.
 */
export function launchPackaged(appPath, port) {
  const profile = mkdtempSync(path.join(tmpdir(), 'mutatis-profile-'));
  const binary = path.join(appPath, 'Contents/MacOS/Mutatis Desktop');
  const child = spawn(
    binary,
    [`--remote-debugging-port=${port}`, `--user-data-dir=${profile}`],
    { stdio: ['ignore', 'pipe', 'pipe'] },
  );
  let stderr = '';
  child.stderr.on('data', (chunk) => {
    stderr += String(chunk);
  });
  return {
    child,
    profile,
    stderr: () => stderr,
    dispose() {
      child.kill('SIGTERM');
      rmSync(profile, { recursive: true, force: true });
    },
  };
}

export async function waitForPage(port, timeoutMs = 30000) {
  const deadline = Date.now() + timeoutMs;
  let last = '';
  while (Date.now() < deadline) {
    try {
      const response = await fetch(`http://127.0.0.1:${port}/json/list`);
      const list = await response.json();
      const page = list.find((target) => target.type === 'page' && target.webSocketDebuggerUrl);
      if (page) return page;
      last = 'no page target yet';
    } catch (error) {
      last = String(error);
    }
    await new Promise((resolve) => setTimeout(resolve, 250));
  }
  throw new Error(`No debuggable page after ${timeoutMs}ms (${last})`);
}

export function connect(url) {
  return new Promise((resolve, reject) => {
    const socket = new WebSocket(url);
    const pending = new Map();
    const listeners = [];
    let nextId = 1;

    socket.addEventListener('open', () =>
      resolve({
        send(method, params = {}) {
          const id = nextId++;
          socket.send(JSON.stringify({ id, method, params }));
          return new Promise((ok, bad) => pending.set(id, { ok, bad }));
        },
        on(handler) {
          listeners.push(handler);
        },
        close: () => socket.close(),
      }),
    );
    socket.addEventListener('error', reject);
    socket.addEventListener('message', (event) => {
      const message = JSON.parse(event.data);
      if (message.id && pending.has(message.id)) {
        const { ok, bad } = pending.get(message.id);
        pending.delete(message.id);
        if (message.error) bad(new Error(message.error.message));
        else ok(message.result);
        return;
      }
      for (const handler of listeners) handler(message);
    });
  });
}

export async function evaluate(cdp, expression) {
  const result = await cdp.send('Runtime.evaluate', {
    expression,
    awaitPromise: true,
    returnByValue: true,
  });
  if (result.exceptionDetails) {
    throw new Error(result.exceptionDetails.exception?.description ?? 'evaluation failed');
  }
  return result.result.value;
}

export function reporter() {
  const results = [];
  return {
    results,
    check(name, ok, detail = '') {
      results.push({ name, ok, detail });
      console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? ` — ${detail}` : ''}`);
      return ok;
    },
    finish() {
      const failed = results.filter((entry) => !entry.ok);
      console.log(`\n${results.length - failed.length}/${results.length} checks passed.`);
      if (failed.length > 0) {
        console.log('\nFAILED:');
        for (const entry of failed) {
          console.log(`  - ${entry.name}${entry.detail ? ` — ${entry.detail}` : ''}`);
        }
      }
      return failed.length;
    },
  };
}
