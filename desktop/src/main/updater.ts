import { app, BrowserWindow } from 'electron';
import { existsSync } from 'node:fs';
import path from 'node:path';
import type { UpdaterPhase, UpdaterState } from '../common/ipc.js';

/**
 * Application updates — MANUAL, and off unless the build can prove it is
 * updatable.
 *
 * Three deliberate refusals, all of them privacy or safety decisions rather
 * than taste:
 *
 * 1. NOTHING RUNS AT STARTUP. `checkForUpdatesAndNotify()` — the one-liner every
 *    Electron tutorial reaches for — fires an HTTPS request the moment the app
 *    opens. This app is used on confidential documents inside law firms, and the
 *    egress audit's central claim is that the only thing that ever leaves the
 *    machine is an AI chat turn the user sent (which carries the
 *    changed wording of the open comparison). A silent startup beacon would make
 *    that claim false, and "it's only a version number" is exactly the argument
 *    that erodes it. The user presses a button.
 * 2. `autoDownload = false`. Checking tells you a version exists; it must not
 *    also pull tens of megabytes over someone's tethered connection.
 * 3. `autoInstallOnAppQuit = false`. Swapping the binary during a quit the user
 *    did not connect to updating is a surprise, and a half-written app bundle
 *    after a forced shutdown is a support call.
 *
 * The module also does not import electron-updater at the top level. The import
 * is dynamic, inside the check, so a session that never asks for an update never
 * loads the library at all — which is what lets the audit assert that no update
 * machinery is even resident at startup.
 */

let state: UpdaterState = { phase: 'unsupported', version: '0.0.0', blockers: [] };
let wired = false;

/**
 * Is this .app code-signed?
 *
 * Squirrel.Mac refuses to install an update into an unsigned bundle — it cannot
 * verify that the replacement came from the same publisher, which is the whole
 * point. Rather than let the user press Check and then meet
 * "Could not get code signature for running application", say so up front.
 *
 * The presence of `Contents/_CodeSignature/CodeResources` is the cheap test. It
 * cannot tell a Developer ID signature from an ad-hoc one — `codesign` could,
 * but shelling out on every status read to learn something this build already
 * knows is not a trade worth making.
 */
function macBundleIsSigned(): boolean {
  // .../Mutatis Desktop.app/Contents/Resources/app.asar → the bundle root
  const contents = path.resolve(app.getAppPath(), '..', '..');
  return existsSync(path.join(contents, '_CodeSignature', 'CodeResources'));
}

/**
 * `app-update.yml` is written by electron-builder ONLY when a publish config
 * exists. No file, no feed — and therefore no URL in the shipped app for anyone
 * to point at anything.
 */
function feedConfigured(): boolean {
  return existsSync(path.join(process.resourcesPath ?? '', 'app-update.yml'));
}

function blockers(): string[] {
  const reasons: string[] = [];
  if (!app.isPackaged) {
    reasons.push('Updates only apply to an installed build, not to a development run.');
    return reasons;
  }
  if (!feedConfigured()) {
    reasons.push('This build was packaged without an update feed, so there is nowhere to check.');
  }
  if (process.platform === 'darwin' && !macBundleIsSigned()) {
    reasons.push('macOS will not replace an unsigned application, so this copy cannot update itself.');
  }
  return reasons;
}

function publish(next: Partial<UpdaterState> & { phase: UpdaterPhase }): UpdaterState {
  state = {
    version: app.getVersion(),
    blockers: state.blockers,
    ...next,
  };
  for (const win of BrowserWindow.getAllWindows()) {
    if (!win.isDestroyed()) win.webContents.send('updates:state', state);
  }
  return state;
}

export function updaterState(): UpdaterState {
  const reasons = blockers();
  state = {
    ...state,
    version: app.getVersion(),
    blockers: reasons,
    phase: reasons.length > 0 ? 'unsupported' : state.phase === 'unsupported' ? 'idle' : state.phase,
  };
  return state;
}

/**
 * electron-updater is CommonJS and publishes `autoUpdater` as a getter on
 * `module.exports`, which named-import interop does not always see. Take the
 * default export and destructure — and do it here rather than at module scope so
 * the library is only ever loaded on a deliberate user action.
 */
async function loadUpdater() {
  const mod = await import('electron-updater');
  const updater = (mod as unknown as { default?: { autoUpdater: unknown } }).default ?? mod;
  const autoUpdater = (updater as { autoUpdater: import('electron-updater').AppUpdater }).autoUpdater;

  if (!wired) {
    wired = true;
    autoUpdater.autoDownload = false;
    autoUpdater.autoInstallOnAppQuit = false;
    // electron-updater logs the full feed response at info level. Left off: this
    // is a legal-documents app and its log file is not a place to accumulate
    // anything nobody asked for.
    autoUpdater.logger = null;

    autoUpdater.on('download-progress', (progress) => {
      publish({ phase: 'downloading', percent: Math.round(progress.percent) });
    });
    autoUpdater.on('update-downloaded', (info) => {
      publish({ phase: 'downloaded', newVersion: info.version });
    });
    autoUpdater.on('error', (error: Error) => {
      publish({ phase: 'error', message: error.message });
    });
  }
  return autoUpdater;
}

export async function checkForUpdates(): Promise<UpdaterState> {
  const current = updaterState();
  if (current.blockers.length > 0) return current;

  publish({ phase: 'checking' });
  try {
    const autoUpdater = await loadUpdater();
    const result = await autoUpdater.checkForUpdates();
    const version = result?.updateInfo.version;
    if (!version || version === app.getVersion()) {
      return publish({ phase: 'up-to-date' });
    }
    return publish({ phase: 'available', newVersion: version });
  } catch (error) {
    // Only the message. A feed error can carry the URL and, on some servers,
    // a signed query string — the same reason the chat stream forwards
    // `error.message` and never the request.
    return publish({ phase: 'error', message: error instanceof Error ? error.message : 'Check failed.' });
  }
}

export async function downloadUpdate(): Promise<UpdaterState> {
  if (state.phase !== 'available') return updaterState();
  publish({
    phase: 'downloading',
    percent: 0,
    // exactOptionalPropertyTypes: spreading a possibly-undefined optional is an
    // error, so only carry the key when there is a value to carry.
    ...(state.newVersion === undefined ? {} : { newVersion: state.newVersion }),
  });
  try {
    const autoUpdater = await loadUpdater();
    await autoUpdater.downloadUpdate();
    return state;
  } catch (error) {
    return publish({ phase: 'error', message: error instanceof Error ? error.message : 'Download failed.' });
  }
}

export async function installUpdate(): Promise<boolean> {
  if (state.phase !== 'downloaded') return false;
  const autoUpdater = await loadUpdater();
  // Restarts the app. Only reachable from a downloaded update the user asked
  // for twice — once to check, once to download.
  autoUpdater.quitAndInstall(false, true);
  return true;
}
