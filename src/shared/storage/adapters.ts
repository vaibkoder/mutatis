import type { RemoteAdapter } from './types.js';

/**
 * Stub adapters for GitHub and SharePoint.
 * These carry ONLY user-authored library data (clause library, numbering schemes).
 * They NEVER transmit document content. (ADR-0001/ADR-0003)
 *
 * Actual network calls live only in these files — the no-network-egress ESLint rule
 * allow-lists src/shared/storage/adapters*.
 *
 * Full implementations are a v2 task; these stubs satisfy the StorageProvider interface
 * and return isConfigured() === false until configured by the user.
 */

export function createGitHubAdapter(): RemoteAdapter {
  return {
    id: 'github',
    isConfigured: () => false,
    async get(_key) { throw new Error('GitHub adapter not configured'); },
    async set(_key, _value) { throw new Error('GitHub adapter not configured'); },
    async list(_prefix) { return []; },
  };
}

export function createSharePointAdapter(): RemoteAdapter {
  return {
    id: 'sharepoint',
    isConfigured: () => false,
    async get(_key) { throw new Error('SharePoint adapter not configured'); },
    async set(_key, _value) { throw new Error('SharePoint adapter not configured'); },
    async list(_prefix) { return []; },
  };
}
