import { app, BrowserWindow } from 'electron';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { registerAppScheme, serveRendererOverAppScheme } from './app-protocol.js';
import { createMainWindow, preloadPathFor } from './window.js';
import { registerIpc } from './ipc.js';
import { closeThreads } from './threads.js';

const mainDir = path.dirname(fileURLToPath(import.meta.url));

/** electron-vite sets this in dev only. Its absence means we are packaged. */
const devServerUrl = process.env['ELECTRON_RENDERER_URL'];

// MUST happen before app-ready.
registerAppScheme();

// One instance only — two windows over one safeStorage-backed config is a
// corruption hazard, not just a UX wart.
if (!app.requestSingleInstanceLock()) {
  app.quit();
} else {
  app.on('second-instance', () => {
    const [win] = BrowserWindow.getAllWindows();
    if (win) {
      if (win.isMinimized()) win.restore();
      win.focus();
    }
  });

  void app.whenReady().then(() => {
    registerIpc();

    if (!devServerUrl) {
      serveRendererOverAppScheme(path.join(mainDir, '../renderer'));
    }

    createMainWindow(preloadPathFor(mainDir), devServerUrl);

    app.on('activate', () => {
      if (BrowserWindow.getAllWindows().length === 0) {
        createMainWindow(preloadPathFor(mainDir), devServerUrl);
      }
    });
  });

  app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit();
  });

  // Flush the WAL and release the file handle. Without this a quit mid-write can
  // leave a -wal alongside the database that the next launch has to recover.
  app.on('will-quit', () => {
    closeThreads();
  });
}
