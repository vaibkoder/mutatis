import type { ChatTurnRequest, ProviderId, StoredMessage } from '../common/ipc.js';

/**
 * Payload validation for the IPC surface.
 *
 * Separate from `ipc.ts` so it can be tested without booting Electron, and
 * because it is the one place that decides what counts as well-formed. The
 * renderer is the process that opens hostile `.docx` files; main is the one
 * holding the API key. Everything arriving from there is untrusted even though
 * it is local.
 */

export function isProviderId(value: unknown): value is ProviderId {
  return value === 'deepseek' || value === 'alibaba';
}

export function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

/** Shape-check a chat turn before any of it reaches a provider. */
export function isChatTurnRequest(value: unknown): value is ChatTurnRequest {
  if (!isRecord(value)) return false;
  if (!Array.isArray(value['messages'])) return false;

  const model = value['model'];
  if (!isRecord(model)) return false;
  if (!isProviderId(model['provider'])) return false;
  if (typeof model['model'] !== 'string' || model['model'] === '') return false;

  const system = value['system'];
  if (system !== undefined && typeof system !== 'string') return false;

  return value['messages'].every(isUiMessageLike);
}

function isUiMessageLike(value: unknown): boolean {
  if (!isRecord(value)) return false;
  if (typeof value['id'] !== 'string') return false;
  const role = value['role'];
  if (role !== 'user' && role !== 'assistant' && role !== 'system') return false;
  return Array.isArray(value['parts']);
}

/** Shape-check a message on its way into SQLite. */
export function toStoredMessage(value: unknown): StoredMessage | null {
  if (!isRecord(value)) return null;
  const { id, threadId, role, parts, createdAt } = value;
  if (typeof id !== 'string' || id === '') return null;
  if (typeof threadId !== 'string' || threadId === '') return null;
  if (role !== 'user' && role !== 'assistant' && role !== 'system') return null;
  if (!Array.isArray(parts)) return null;
  return {
    id,
    threadId,
    role,
    parts,
    createdAt: typeof createdAt === 'number' && Number.isFinite(createdAt) ? createdAt : Date.now(),
  };
}
