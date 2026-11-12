import { describe, it, expect, beforeEach } from 'vitest';
import {
  setActiveHost,
  getActiveHost,
  hasActiveHost,
  _resetActiveHost,
} from '@shared/host/index.js';
import type { HostPort } from '@shared/host/index.js';

// A do-nothing host — enough to exercise the registry.
const stubHost = {
  loadDocument: async () => ({ paragraphs: [], footnotes: [], loadedAt: 0 }),
  onDocumentChanged: () => ({ dispose() {} }),
  reveal: async () => {},
  readDocx: async () => ({} as never),
  writeBack: async () => {},
  storage: {} as never,
} as unknown as HostPort;

describe('active-host registry', () => {
  beforeEach(() => _resetActiveHost());

  it('reports no host before one is installed', () => {
    expect(hasActiveHost()).toBe(false);
  });

  it('throws a clear error if getActiveHost is called before install', () => {
    expect(() => getActiveHost()).toThrow(/No active HostPort/);
  });

  it('returns the installed host', () => {
    setActiveHost(stubHost);
    expect(hasActiveHost()).toBe(true);
    expect(getActiveHost()).toBe(stubHost);
  });

  it('the last installed host wins', () => {
    const other = { ...stubHost } as HostPort;
    setActiveHost(stubHost);
    setActiveHost(other);
    expect(getActiveHost()).toBe(other);
  });

  it('_resetActiveHost clears the registry', () => {
    setActiveHost(stubHost);
    _resetActiveHost();
    expect(hasActiveHost()).toBe(false);
  });
});
