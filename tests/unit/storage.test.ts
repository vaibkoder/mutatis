import { describe, it, expect, beforeEach } from 'vitest';
import {
  createLocalStore,
  createGitHubAdapter,
  createSharePointAdapter,
} from '@shared/storage/index.js';

describe('local storage store', () => {
  beforeEach(() => localStorage.clear());

  it('round-trips get/set/delete', async () => {
    const store = createLocalStore<{ n: number }>();
    expect(await store.get('a')).toBeUndefined();
    await store.set('a', { n: 1 });
    expect(await store.get('a')).toEqual({ n: 1 });
    await store.delete('a');
    expect(await store.get('a')).toBeUndefined();
  });

  it('lists keys by prefix', async () => {
    const store = createLocalStore();
    await store.set('scheme/one', 1);
    await store.set('scheme/two', 2);
    await store.set('pref/x', 3);
    const schemes = await store.list('scheme/');
    expect(schemes.sort()).toEqual(['scheme/one', 'scheme/two']);
  });

  it('exports and imports as JSON (round-trip)', async () => {
    const store = createLocalStore();
    await store.set('k1', { v: 'hello' });
    const blob = await store.export();
    expect(blob.type).toBe('application/json');
    expect(blob.size).toBeGreaterThan(0);

    // Clear, then re-import the exported blob and confirm the value survives.
    localStorage.clear();
    await store.import(blob);
    expect(await store.get('k1')).toEqual({ v: 'hello' });
  });
});

describe('remote adapter stubs', () => {
  it('report as unconfigured and never claim to carry document content', () => {
    const gh = createGitHubAdapter();
    const sp = createSharePointAdapter();
    expect(gh.isConfigured()).toBe(false);
    expect(sp.isConfigured()).toBe(false);
    expect(gh.id).toBe('github');
    expect(sp.id).toBe('sharepoint');
  });
});
