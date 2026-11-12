import { BrowserWindow, shell } from 'electron';
import path from 'node:path';
import { APP_ENTRY_URL } from './app-protocol.js';

/**
 * Create the main window.
 *
 * Security flags are asserted EXPLICITLY rather than left to Electron's
 * defaults. They are the same values Electron defaults to today, but writing
 * them down means a future Electron default change cannot silently weaken the
 * app, and a reviewer can see the posture without knowing Electron's version
 * history.
 *
 * Note what is deliberately absent: there is no `open-win`-style IPC handler
 * that spawns child windows with `nodeIntegration: true`, and preload exposes a
 * narrow, named API rather than `ipcRenderer` itself.
 */
export function createMainWindow(preloadPath: string, devServerUrl?: string): BrowserWindow {
  const win = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 940,
    minHeight: 600,
    show: false, // avoid the white flash; revealed on ready-to-show
    backgroundColor: '#ffffff',
    titleBarStyle: process.platform === 'darwin' ? 'hiddenInset' : 'default',
    webPreferences: {
      preload: preloadPath,
      contextIsolation: true,
      sandbox: true,
      nodeIntegration: false,
      webSecurity: true,
      spellcheck: false,
    },
  });

  win.once('ready-to-show', () => win.show());

  // Deny every popup; route genuine external links to the system browser.
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (url.startsWith('https://')) void shell.openExternal(url);
    return { action: 'deny' };
  });

  // The renderer may never navigate away from its own origin.
  win.webContents.on('will-navigate', (event, url) => {
    const target = new URL(url);
    const allowed = devServerUrl ? new URL(devServerUrl).origin : 'app://.';
    if (target.origin !== allowed) event.preventDefault();
  });

  if (devServerUrl) {
    void win.loadURL(devServerUrl);
  } else {
    void win.loadURL(APP_ENTRY_URL);
  }

  return win;
}

export function preloadPathFor(mainDir: string): string {
  // .cjs, not .mjs — a sandboxed renderer cannot load an ESM preload.
  return path.join(mainDir, '../preload/index.cjs');
}
