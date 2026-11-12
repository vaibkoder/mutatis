import type { Store } from './types.js';

const PREFIX = 'mutatis.';

function fullKey(key: string): string {
  return `${PREFIX}${key}`;
}

/** Read a Blob's text, falling back to FileReader where Blob.text() is unavailable. */
async function readBlobText(blob: Blob): Promise<string> {
  if (typeof blob.text === 'function') return blob.text();
  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error ?? new Error('FileReader failed'));
    reader.readAsText(blob);
  });
}

/**
 * localStorage-backed store.
 * Used for machine-local user preferences and numbering scheme library.
 */
export function createLocalStore<T = unknown>(): Store<T> {
  return {
    async get(key) {
      const raw = localStorage.getItem(fullKey(key));
      if (raw === null) return undefined;
      try { return JSON.parse(raw) as T; }
      catch { return undefined; }
    },
    async set(key, value) {
      localStorage.setItem(fullKey(key), JSON.stringify(value));
    },
    async delete(key) {
      localStorage.removeItem(fullKey(key));
    },
    async list(prefix) {
      const result: string[] = [];
      const full = PREFIX + (prefix ?? '');
      for (let i = 0; i < localStorage.length; i++) {
        const k = localStorage.key(i);
        if (k?.startsWith(full)) {
          result.push(k.slice(PREFIX.length));
        }
      }
      return result.sort();
    },
    async export() {
      const obj: Record<string, unknown> = {};
      for (let i = 0; i < localStorage.length; i++) {
        const k = localStorage.key(i);
        if (k?.startsWith(PREFIX)) {
          const raw = localStorage.getItem(k);
          if (raw !== null) {
            try { obj[k.slice(PREFIX.length)] = JSON.parse(raw); }
            catch { /* skip malformed */ }
          }
        }
      }
      return new Blob([JSON.stringify(obj, null, 2)], { type: 'application/json' });
    },
    async import(blob) {
      const text = await readBlobText(blob);
      const obj = JSON.parse(text) as Record<string, unknown>;
      for (const [k, v] of Object.entries(obj)) {
        localStorage.setItem(fullKey(k), JSON.stringify(v));
      }
    },
  };
}
