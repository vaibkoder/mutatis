import { describe, it, expect, beforeEach, vi } from 'vitest';
import { detectCapabilities, getCapabilities, isCapable, _resetCapabilities } from '@shared/office/index.js';
import { waitForOffice, _resetReady } from '@shared/office/index.js';

describe('office capabilities', () => {
  beforeEach(() => {
    _resetCapabilities();
  });

  it('reports all-false before detection', () => {
    const caps = getCapabilities();
    expect(caps.wordApi15).toBe(false);
    expect(caps.wordApi16).toBe(false);
  });

  it('detects supported requirement sets from Office', () => {
    const isSetSupported = vi.fn((name: string, version: string) => {
      if (name !== 'WordApi') return false;
      // Host supports up to 1.5 but not 1.6
      return parseFloat(version) <= 1.5;
    });
    (globalThis as unknown as { Office: unknown }).Office = {
      context: { requirements: { isSetSupported } },
    };

    const caps = detectCapabilities();
    expect(caps.wordApi15).toBe(true);
    expect(caps.wordApi16).toBe(false);
    expect(caps.ooxmlLoop).toBe(true);
    expect(isCapable('wordApi15')).toBe(true);

    delete (globalThis as unknown as { Office?: unknown }).Office;
  });

  it('degrades to all-false when Office throws', () => {
    _resetCapabilities();
    (globalThis as unknown as { Office: unknown }).Office = {
      context: { requirements: { isSetSupported: () => { throw new Error('nope'); } } },
    };
    const caps = detectCapabilities();
    expect(caps.wordApi15).toBe(false);
    delete (globalThis as unknown as { Office?: unknown }).Office;
  });

  it('detects a desktop host (PlatformType.PC) → isDesktop true (04b)', () => {
    _resetCapabilities();
    (globalThis as unknown as { Office: unknown }).Office = {
      context: { requirements: { isSetSupported: () => false }, platform: 'PC' },
      PlatformType: { PC: 'PC', Mac: 'Mac', OfficeOnline: 'OfficeOnline' },
    };
    expect(detectCapabilities().isDesktop).toBe(true);
    delete (globalThis as unknown as { Office?: unknown }).Office;
  });

  it('treats Word Online (PlatformType.OfficeOnline) as not desktop (04b)', () => {
    _resetCapabilities();
    (globalThis as unknown as { Office: unknown }).Office = {
      context: { requirements: { isSetSupported: () => false }, platform: 'OfficeOnline' },
      PlatformType: { PC: 'PC', Mac: 'Mac', OfficeOnline: 'OfficeOnline' },
    };
    expect(detectCapabilities().isDesktop).toBe(false);
    delete (globalThis as unknown as { Office?: unknown }).Office;
  });

  it('isDesktop is false when platform/PlatformType are absent (safe fallback)', () => {
    _resetCapabilities();
    (globalThis as unknown as { Office: unknown }).Office = {
      context: { requirements: { isSetSupported: () => false } },
    };
    expect(detectCapabilities().isDesktop).toBe(false);
    delete (globalThis as unknown as { Office?: unknown }).Office;
  });
});

describe('waitForOffice — host detection', () => {
  beforeEach(() => {
    _resetReady();
    _resetCapabilities();
  });

  function installOffice(host: unknown): void {
    (globalThis as unknown as { Office: unknown }).Office = {
      HostType: { Word: 'Word' },
      onReady: (cb: (info: { host: unknown }) => void) => cb({ host }),
      context: { requirements: { isSetSupported: () => false } },
    };
  }

  it('resolves when the host is Word', async () => {
    installOffice('Word');
    await expect(waitForOffice()).resolves.toBeUndefined();
    delete (globalThis as unknown as { Office?: unknown }).Office;
  });

  it('resolves when the host is null (Word-on-the-web load race)', async () => {
    installOffice(null);
    await expect(waitForOffice()).resolves.toBeUndefined();
    delete (globalThis as unknown as { Office?: unknown }).Office;
  });

  it('rejects for a genuinely different host', async () => {
    installOffice('Excel');
    await expect(waitForOffice()).rejects.toThrow('Unexpected host: Excel');
    delete (globalThis as unknown as { Office?: unknown }).Office;
  });
});
