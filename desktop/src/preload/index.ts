import { contextBridge, ipcRenderer } from 'electron';
import { PORT_CLOSED } from '../common/ipc.js';
import type {
  ChatChunk,
  ChatTurnRequest,
  MutatisApi,
  ProviderId,
  StoredMessage,
  UpdaterState,
} from '../common/ipc.js';

/**
 * The preload bridge.
 *
 * Expose the smallest possible named API — never `ipcRenderer` itself, and never
 * a generic `invoke(channel, ...args)` passthrough, which is the same thing
 * wearing a hat.
 *
 * NOTHING SECRET PASSES THROUGH HERE. API keys live in the main process in
 * safeStorage. A key can be written inwards; there is no operation that reads
 * one back out, so a compromised renderer has nothing to exfiltrate.
 *
 * The MessagePort for a chat stream terminates HERE, in the isolated world,
 * rather than being handed to the renderer. The renderer gets a callback and a
 * stop function, which is all it needs; a real port in page context would be a
 * transferable object the page could hold onto past the turn that owns it.
 */

const api: MutatisApi = {
  platform: process.platform,

  providers: () => ipcRenderer.invoke('providers:list'),

  secrets: {
    status: () => ipcRenderer.invoke('secrets:status'),
    set: (provider: ProviderId, key: string) => ipcRenderer.invoke('secrets:set', provider, key),
  },

  threads: {
    list: () => ipcRenderer.invoke('threads:list'),
    load: (id: string) => ipcRenderer.invoke('threads:load', id),
    ensure: (id: string, title: string) => ipcRenderer.invoke('threads:ensure', id, title),
    rename: (id: string, title: string) => ipcRenderer.invoke('threads:rename', id, title),
    save: (message: StoredMessage) => ipcRenderer.invoke('threads:save', message),
    remove: (id: string) => ipcRenderer.invoke('threads:delete', id),
  },

  updates: {
    state: () => ipcRenderer.invoke('updates:state'),
    check: () => ipcRenderer.invoke('updates:check'),
    download: () => ipcRenderer.invoke('updates:download'),
    install: () => ipcRenderer.invoke('updates:install'),
    onState(listener: (state: UpdaterState) => void): () => void {
      // One named channel, and the Electron `event` argument is dropped rather
      // than forwarded — it carries `sender`, which is a live handle the page
      // has no business holding.
      const handler = (_event: unknown, state: UpdaterState): void => listener(state);
      ipcRenderer.on('updates:state', handler);
      return () => ipcRenderer.off('updates:state', handler);
    },
  },

  openKeyPage: (url: string) => ipcRenderer.invoke('shell:openExternal', url),

  startChat(request: ChatTurnRequest, onChunk: (chunk: ChatChunk) => void): () => void {
    const channel = new MessageChannel();

    channel.port1.onmessage = (event: MessageEvent): void => {
      onChunk(event.data as ChatChunk);
    };
    channel.port1.onmessageerror = (): void => {
      onChunk({ type: 'error', errorText: 'The response could not be read.' });
    };
    // If main goes away mid-stream the port closes with no terminal chunk. This
    // sentinel — filtered out by the transport — keeps the renderer's stream from
    // hanging open forever. It does NOT fire when we close the port ourselves.
    channel.port1.addEventListener('close', () => {
      onChunk({ type: PORT_CLOSED });
    });

    // `postMessage`, not `invoke`: this is the one call that transfers a port.
    ipcRenderer.postMessage('chat:start', request, [channel.port2]);

    return (): void => {
      // Closing the port IS Stop — main sees `close` and aborts the request.
      try {
        channel.port1.close();
      } catch {
        /* already closed */
      }
    };
  },
};

contextBridge.exposeInMainWorld('mutatis', api);
