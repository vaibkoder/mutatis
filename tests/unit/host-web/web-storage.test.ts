/**
 * host-web — web storage provider.
 * Both stores are localStorage-backed and round-trip; nothing touches Office.
 */
import { describe, it, expect } from 'vitest';
import { createWebStorageProvider } from '@host-web/web-storage.js';

describe('host-web storage provider', () => {
  it('exposes local + docSettings stores', () => {
    const storage = createWebStorageProvider();
    expect(storage.local).toBeDefined();
    expect(storage.docSettings).toBeDefined();
  });

  it('local store round-trips a value', async () => {
    const { local } = createWebStorageProvider();
    await local.set('k', { a: 1 });
    expect(await local.get('k')).toEqual({ a: 1 });
    expect(await local.list()).toContain('k');
    await local.delete('k');
    expect(await local.get('k')).toBeUndefined();
  });

  it('docSettings store round-trips a value (localStorage-backed on web)', async () => {
    const { docSettings } = createWebStorageProvider();
    await docSettings.set('profile', 'SG');
    expect(await docSettings.get('profile')).toBe('SG');
  });
});
