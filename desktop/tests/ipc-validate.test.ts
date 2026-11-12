import { describe, it, expect } from 'vitest';
import { isChatTurnRequest, isProviderId, toStoredMessage } from '../src/main/ipc-validate.js';

/**
 * Main treats everything from the renderer as untrusted. These tests are the
 * proof that the treatment is real — a validator nobody exercises is decoration.
 */

const goodMessage = { id: 'm1', role: 'user', parts: [{ type: 'text', text: 'hi' }] };
const goodTurn = {
  messages: [goodMessage],
  model: { provider: 'deepseek', model: 'deepseek-v4-flash' },
};

describe('isProviderId', () => {
  it('accepts only the providers this build knows', () => {
    expect(isProviderId('deepseek')).toBe(true);
    expect(isProviderId('alibaba')).toBe(true);
    expect(isProviderId('openai')).toBe(false);
    expect(isProviderId('')).toBe(false);
    expect(isProviderId(null)).toBe(false);
    expect(isProviderId({ toString: () => 'deepseek' })).toBe(false);
  });
});

describe('isChatTurnRequest', () => {
  it('accepts a well-formed turn', () => {
    expect(isChatTurnRequest(goodTurn)).toBe(true);
    expect(isChatTurnRequest({ ...goodTurn, system: 'be brief' })).toBe(true);
  });

  it('rejects anything that is not an object', () => {
    for (const value of [null, undefined, 'x', 7, [], true]) {
      expect(isChatTurnRequest(value)).toBe(false);
    }
  });

  it('rejects an unknown provider — the model is resolved from this field', () => {
    expect(
      isChatTurnRequest({ ...goodTurn, model: { provider: 'evil', model: 'x' } }),
    ).toBe(false);
  });

  it('rejects a missing or empty model id', () => {
    expect(isChatTurnRequest({ ...goodTurn, model: { provider: 'deepseek' } })).toBe(false);
    expect(
      isChatTurnRequest({ ...goodTurn, model: { provider: 'deepseek', model: '' } }),
    ).toBe(false);
  });

  it('rejects a non-string system prompt', () => {
    expect(isChatTurnRequest({ ...goodTurn, system: { toString: () => 'x' } })).toBe(false);
    expect(isChatTurnRequest({ ...goodTurn, system: 12 })).toBe(false);
  });

  it('rejects malformed messages rather than passing them to the provider', () => {
    expect(isChatTurnRequest({ ...goodTurn, messages: 'not an array' })).toBe(false);
    expect(isChatTurnRequest({ ...goodTurn, messages: [{ role: 'user', parts: [] }] })).toBe(false);
    expect(isChatTurnRequest({ ...goodTurn, messages: [{ id: 'm', role: 'root', parts: [] }] })).toBe(
      false,
    );
    expect(isChatTurnRequest({ ...goodTurn, messages: [{ id: 'm', role: 'user' }] })).toBe(false);
    expect(isChatTurnRequest({ ...goodTurn, messages: [null] })).toBe(false);
  });

  it('rejects an array dressed up as the request object', () => {
    expect(isChatTurnRequest([goodMessage])).toBe(false);
  });
});

describe('toStoredMessage', () => {
  it('normalises a well-formed message', () => {
    const stored = toStoredMessage({
      id: 'm1',
      threadId: 't1',
      role: 'assistant',
      parts: [{ type: 'text', text: 'ok' }],
      createdAt: 1700000000000,
    });
    expect(stored).toEqual({
      id: 'm1',
      threadId: 't1',
      role: 'assistant',
      parts: [{ type: 'text', text: 'ok' }],
      createdAt: 1700000000000,
    });
  });

  it('supplies a timestamp rather than storing a bad one', () => {
    const before = Date.now();
    const stored = toStoredMessage({
      id: 'm1',
      threadId: 't1',
      role: 'user',
      parts: [],
      createdAt: Number.NaN,
    });
    expect(stored?.createdAt).toBeGreaterThanOrEqual(before);
  });

  it('rejects rows that would poison the thread store', () => {
    expect(toStoredMessage(null)).toBeNull();
    expect(toStoredMessage({ id: '', threadId: 't', role: 'user', parts: [] })).toBeNull();
    expect(toStoredMessage({ id: 'm', threadId: '', role: 'user', parts: [] })).toBeNull();
    expect(toStoredMessage({ id: 'm', threadId: 't', role: 'tool', parts: [] })).toBeNull();
    expect(toStoredMessage({ id: 'm', threadId: 't', role: 'user', parts: 'x' })).toBeNull();
  });
});
