import { ipcMain, shell, type IpcMainInvokeEvent, type IpcMainEvent } from 'electron';
import { APP_ENTRY_URL } from './app-protocol.js';
import { runChatTurn } from './chat/stream.js';
import { isChatTurnRequest, isProviderId, toStoredMessage } from './ipc-validate.js';
import { PROVIDERS, DEFAULT_MODEL } from './chat/providers.js';
import { encryptionAvailable, keyStatus, setKey } from './secrets.js';
import { checkForUpdates, downloadUpdate, installUpdate, updaterState } from './updater.js';
import {
  deleteThread,
  ensureThread,
  listThreads,
  loadThread,
  renameThread,
  upsertMessage,
} from './threads.js';

/**
 * The IPC surface. Everything the renderer can ask main to do.
 *
 * Renderer data is treated as untrusted even though it is local: every payload
 * is shape-checked here, and every sender is verified against our own origin.
 * The renderer is the process that parses hostile documents; main is the one
 * holding the keys.
 *
 * Note what is absent: there is no handler that RETURNS an API key. The renderer
 * can set one and ask whether one exists. It can never read one back.
 */

/**
 * The one origin allowed to talk to main.
 *
 * In a packaged build that is the `app://` scheme and nothing else. The dev
 * server is added ONLY when electron-vite has actually set its variable, so a
 * packaged app cannot be talked to by a frame claiming a localhost origin.
 */
function allowedOrigin(): string {
  const dev = process.env['ELECTRON_RENDERER_URL'];
  if (dev !== undefined && dev !== '') {
    try {
      return new URL(dev).origin;
    } catch {
      /* fall through to the packaged origin */
    }
  }
  return new URL(APP_ENTRY_URL).origin;
}

function senderIsOurs(event: IpcMainInvokeEvent | IpcMainEvent): boolean {
  const url = event.senderFrame?.url ?? '';
  if (url === '') return false;
  try {
    return new URL(url).origin === allowedOrigin();
  } catch {
    return false;
  }
}

export function registerIpc(): void {
  ipcMain.handle('providers:list', (event) => {
    if (!senderIsOurs(event)) return null;
    return {
      providers: PROVIDERS,
      defaultModel: DEFAULT_MODEL,
      encryptionAvailable: encryptionAvailable(),
    };
  });

  ipcMain.handle('secrets:status', (event) => {
    if (!senderIsOurs(event)) return null;
    return keyStatus();
  });

  ipcMain.handle('secrets:set', (event, provider: unknown, key: unknown) => {
    if (!senderIsOurs(event)) return { ok: false, error: 'Rejected sender' };
    if (!isProviderId(provider) || typeof key !== 'string') {
      return { ok: false, error: 'Bad request' };
    }
    try {
      const { mask } = setKey(provider, key);
      return { ok: true, mask };
    } catch (error) {
      return { ok: false, error: error instanceof Error ? error.message : String(error) };
    }
  });

  ipcMain.handle('threads:list', (event) => (senderIsOurs(event) ? listThreads() : []));

  ipcMain.handle('threads:load', (event, id: unknown) => {
    if (!senderIsOurs(event) || typeof id !== 'string') return { messages: [], skipped: 0 };
    return loadThread(id);
  });

  ipcMain.handle('threads:ensure', (event, id: unknown, title: unknown) => {
    if (!senderIsOurs(event) || typeof id !== 'string' || typeof title !== 'string') return false;
    ensureThread(id, title);
    return true;
  });

  ipcMain.handle('threads:rename', (event, id: unknown, title: unknown) => {
    if (!senderIsOurs(event) || typeof id !== 'string' || typeof title !== 'string') return false;
    renameThread(id, title);
    return true;
  });

  ipcMain.handle('threads:save', (event, message: unknown) => {
    if (!senderIsOurs(event)) return false;
    const stored = toStoredMessage(message);
    if (!stored) return false;
    upsertMessage(stored);
    return true;
  });

  ipcMain.handle('threads:delete', (event, id: unknown) => {
    if (!senderIsOurs(event) || typeof id !== 'string') return false;
    deleteThread(id);
    return true;
  });

  // Updates. Every one of these is a button press — see src/main/updater.ts for
  // why nothing here is wired to startup.
  ipcMain.handle('updates:state', (event) => (senderIsOurs(event) ? updaterState() : null));
  ipcMain.handle('updates:check', (event) => (senderIsOurs(event) ? checkForUpdates() : null));
  ipcMain.handle('updates:download', (event) => (senderIsOurs(event) ? downloadUpdate() : null));
  ipcMain.handle('updates:install', (event) => (senderIsOurs(event) ? installUpdate() : false));

  ipcMain.handle('shell:openExternal', (event, url: unknown) => {
    if (!senderIsOurs(event) || typeof url !== 'string') return false;
    // Only ever https, and only ever a provider's own key page — a renderer
    // must not be able to make main open an arbitrary URL.
    const allowed = PROVIDERS.map((p) => p.keyUrl);
    if (!allowed.includes(url)) return false;
    void shell.openExternal(url);
    return true;
  });

  /**
   * Start a chat turn. The renderer transfers one MessagePort per turn; closing
   * it is Stop.
   */
  ipcMain.on('chat:start', (event, payload: unknown) => {
    const port = event.ports[0];
    if (!port) return;
    if (!senderIsOurs(event) || !isChatTurnRequest(payload)) {
      port.start();
      port.postMessage({ type: 'error', errorText: 'Rejected request' });
      port.close();
      return;
    }
    runChatTurn(port, payload);
  });
}
