import { detectCapabilities } from './capabilities.js';

let _readyPromise: Promise<void> | null = null;

/**
 * Resolves once Office.onReady fires. Safe to call multiple times — the same
 * promise is returned. Detects capabilities immediately on resolution.
 */
export function waitForOffice(): Promise<void> {
  if (_readyPromise) return _readyPromise;

  _readyPromise = new Promise<void>((resolve, reject) => {
    if (typeof Office === 'undefined') {
      // Not running inside Office — fail cleanly (e.g. plain browser tab).
      reject(new Error('Office.js not loaded'));
      return;
    }
    Office.onReady(info => {
      // `info.host` is sometimes reported as null/undefined during a Word-on-the-web
      // load race even though the add-in is genuinely running inside Word. The
      // manifest scopes this add-in to Word, so treat an absent host as Word and
      // proceed; only reject for a genuinely different, known host.
      if (info.host == null || info.host === Office.HostType.Word) {
        detectCapabilities();
        resolve();
      } else {
        reject(new Error(`Unexpected host: ${String(info.host)}`));
      }
    });
  });

  return _readyPromise;
}

/** Reset for testing. */
export function _resetReady(): void {
  _readyPromise = null;
}
