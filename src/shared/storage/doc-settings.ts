import type { Store } from './types.js';

/**
 * Office.context.document.settings-backed store.
 * Per-document, persisted inside the .docx (saveAsync, ~2 MB shared budget).
 * Used for document-pinned jurisdiction profile and per-document numbering.
 */
export function createDocSettingsStore<T = unknown>(): Store<T> {
  function getSettings(): Office.Settings {
    return Office.context.document.settings;
  }

  return {
    async get(key) {
      try {
        const val = getSettings().get(key);
        return val as T | undefined;
      } catch {
        return undefined;
      }
    },
    async set(key, value) {
      getSettings().set(key, value);
      await new Promise<void>((resolve, reject) => {
        getSettings().saveAsync(result => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(new Error(result.error?.message ?? 'saveAsync failed'));
        });
      });
    },
    async delete(key) {
      getSettings().remove(key);
      await new Promise<void>((resolve, reject) => {
        getSettings().saveAsync(result => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(new Error(result.error?.message ?? 'saveAsync failed'));
        });
      });
    },
    async list(prefix) {
      // Office.Settings doesn't expose enumeration; return empty unless a key index is maintained.
      // Features that need listing store an index key.
      void prefix;
      return [];
    },
    async export() {
      return new Blob(['{}'], { type: 'application/json' });
    },
    async import(_blob) {
      // Import into doc settings is not supported (document must be explicitly saved by user).
    },
  };
}
