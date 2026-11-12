export type { Store, RemoteAdapter, StorageProvider } from './types.js';
export { createLocalStore } from './local.js';
export { createDocSettingsStore } from './doc-settings.js';
export { createGitHubAdapter, createSharePointAdapter } from './adapters.js';

import { createLocalStore } from './local.js';
import { createDocSettingsStore } from './doc-settings.js';
import type { StorageProvider } from './types.js';

let _provider: StorageProvider | null = null;

/** Get (or lazily initialise) the storage provider singleton. */
export function getStorageProvider(): StorageProvider {
  if (_provider) return _provider;
  _provider = {
    local: createLocalStore(),
    docSettings: createDocSettingsStore(),
  };
  return _provider;
}
