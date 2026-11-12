import { app, safeStorage } from 'electron';
import { mkdirSync, readFileSync, writeFileSync, existsSync } from 'node:fs';
import path from 'node:path';
import type { ProviderId } from '../common/ipc.js';

/**
 * API key storage. MAIN PROCESS ONLY.
 *
 * Keys are encrypted with Electron's `safeStorage` (macOS Keychain, Windows
 * DPAPI, libsecret on Linux) and written to the app's userData directory. They
 * never cross the preload bridge in either direction: the renderer can ask
 * WHETHER a key is set and can SET one, but can never read one back.
 *
 * We deliberately do NOT fall back to plaintext when encryption is unavailable.
 * open-codesign does exactly that behind a single `logger.warn` the user never
 * sees, with a test asserting the behaviour — which means on a Linux box with no
 * keyring, its users' API keys sit in a config file in the clear. Refusing is
 * the honest answer: the user finds out immediately and can fix their keyring,
 * rather than discovering it after a leak.
 */

// Re-exported, not re-declared: the renderer and main must agree on the list of
// providers, and two copies of a union type drift the moment one gains a member.
export type { ProviderId };

interface SecretsFile {
  version: 1;
  /** provider → base64 ciphertext. Never plaintext. */
  keys: Record<string, string>;
  /** provider → masked form, safe to show in the UI. */
  masks: Record<string, string>;
}

const EMPTY: SecretsFile = { version: 1, keys: {}, masks: {} };

function secretsPath(): string {
  return path.join(app.getPath('userData'), 'secrets.json');
}

function read(): SecretsFile {
  const file = secretsPath();
  if (!existsSync(file)) return { ...EMPTY };
  try {
    const parsed = JSON.parse(readFileSync(file, 'utf8')) as Partial<SecretsFile>;
    return {
      version: 1,
      keys: parsed.keys ?? {},
      masks: parsed.masks ?? {},
    };
  } catch {
    // A corrupt file must not brick the app. Treat it as empty; the user
    // re-enters the key, which is recoverable. Throwing here is not.
    return { ...EMPTY };
  }
}

function write(data: SecretsFile): void {
  const dir = app.getPath('userData');
  mkdirSync(dir, { recursive: true });
  writeFileSync(secretsPath(), JSON.stringify(data, null, 2), { mode: 0o600 });
}

/** `sk-…` plus the last four. Enough to recognise, useless to steal. */
export function maskSecret(secret: string): string {
  const trimmed = secret.trim();
  if (trimmed.length <= 8) return '••••';
  return `${trimmed.slice(0, 3)}…${trimmed.slice(-4)}`;
}

export class EncryptionUnavailableError extends Error {
  constructor() {
    super(
      'This system has no secure credential store available, so the key was NOT saved. ' +
        'On Linux install gnome-keyring or kwallet; on macOS and Windows this usually means ' +
        'the Keychain or DPAPI is unavailable to this user. The key was not written to disk ' +
        'in any form.',
    );
    this.name = 'EncryptionUnavailableError';
  }
}

export function setKey(provider: ProviderId, key: string): { mask: string } {
  if (!safeStorage.isEncryptionAvailable()) throw new EncryptionUnavailableError();

  const data = read();
  const trimmed = key.trim();
  if (trimmed === '') {
    delete data.keys[provider];
    delete data.masks[provider];
    write(data);
    return { mask: '' };
  }

  data.keys[provider] = safeStorage.encryptString(trimmed).toString('base64');
  data.masks[provider] = maskSecret(trimmed);
  write(data);
  return { mask: data.masks[provider] };
}

/**
 * Decrypt a key. MAIN PROCESS ONLY — the return value must never be sent over
 * IPC, logged, or included in an error message.
 */
export function getKey(provider: ProviderId): string | null {
  const data = read();
  const cipher = data.keys[provider];
  if (!cipher) return null;
  try {
    return safeStorage.decryptString(Buffer.from(cipher, 'base64'));
  } catch {
    // Typically the OS credential store changed under us (a restored backup, a
    // new machine). The stored blob is useless; say so by reporting no key.
    return null;
  }
}

/** What the renderer is allowed to know: whether a key exists, and its mask. */
export function keyStatus(): Record<ProviderId, { configured: boolean; mask: string }> {
  const data = read();
  const of = (p: ProviderId) => ({
    configured: Boolean(data.keys[p]),
    mask: data.masks[p] ?? '',
  });
  return { deepseek: of('deepseek'), alibaba: of('alibaba') };
}

export function encryptionAvailable(): boolean {
  return safeStorage.isEncryptionAvailable();
}
