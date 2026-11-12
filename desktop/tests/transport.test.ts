import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import type { UIMessage } from 'ai';
import { PORT_CLOSED, type ChatChunk, type ChatTurnRequest, type ModelChoice } from '@common/ipc.js';
import { IpcChatTransport } from '../src/renderer/src/chat/transport.js';

/**
 * The wrapper file is the single point where a chat turn leaves the renderer.
 * If it leaks a chunk, hangs a stream, or fails to translate Stop into a closed
 * port, the failure shows up as a hung UI and an upstream request that keeps
 * billing. Worth testing directly.
 */

interface Harness {
  request: ChatTurnRequest | null;
  emit: (chunk: ChatChunk) => void;
  stopped: number;
}

let harness: Harness;

beforeEach(() => {
  harness = { request: null, emit: () => {}, stopped: 0 };
  (globalThis as { window?: unknown }).window = {
    mutatis: {
      startChat(request: ChatTurnRequest, onChunk: (chunk: ChatChunk) => void) {
        harness.request = request;
        harness.emit = onChunk;
        return () => {
          harness.stopped += 1;
        };
      },
    },
  };
});

afterEach(() => {
  delete (globalThis as { window?: unknown }).window;
});

const MODEL: ModelChoice = { provider: 'deepseek', model: 'deepseek-v4-flash' };

function messages(): UIMessage[] {
  return [{ id: 'm1', role: 'user', parts: [{ type: 'text', text: 'hello' }] }];
}

async function drain(stream: ReadableStream<unknown>): Promise<unknown[]> {
  const seen: unknown[] = [];
  for await (const chunk of stream as unknown as AsyncIterable<unknown>) seen.push(chunk);
  return seen;
}

describe('IpcChatTransport', () => {
  it('forwards chunks in order and ends the stream on finish', async () => {
    const transport = new IpcChatTransport({ model: () => MODEL });
    const stream = await transport.sendMessages({ messages: messages(), abortSignal: undefined });

    const collected = drain(stream);
    harness.emit({ type: 'start' });
    harness.emit({ type: 'text-start', id: 't' });
    harness.emit({ type: 'text-delta', id: 't', delta: 'Hel' });
    harness.emit({ type: 'text-delta', id: 't', delta: 'lo' });
    harness.emit({ type: 'text-end', id: 't' });
    harness.emit({ type: 'finish' });

    expect((await collected).map((chunk) => (chunk as ChatChunk).type)).toEqual([
      'start',
      'text-start',
      'text-delta',
      'text-delta',
      'text-end',
      'finish',
    ]);
  });

  it('ignores chunks that arrive after a terminal one', async () => {
    const transport = new IpcChatTransport({ model: () => MODEL });
    const stream = await transport.sendMessages({ messages: messages(), abortSignal: undefined });

    const collected = drain(stream);
    harness.emit({ type: 'finish' });
    harness.emit({ type: 'text-delta', id: 't', delta: 'late' });

    expect(await collected).toHaveLength(1);
  });

  it('ends the stream on the port-closed sentinel WITHOUT forwarding it', async () => {
    const transport = new IpcChatTransport({ model: () => MODEL });
    const stream = await transport.sendMessages({ messages: messages(), abortSignal: undefined });

    const collected = drain(stream);
    harness.emit({ type: 'start' });
    harness.emit({ type: PORT_CLOSED });

    const seen = await collected;
    expect(seen.map((chunk) => (chunk as ChatChunk).type)).toEqual(['start']);
  });

  it('forwards an error chunk and then ends — the UI must see why it stopped', async () => {
    const transport = new IpcChatTransport({ model: () => MODEL });
    const stream = await transport.sendMessages({ messages: messages(), abortSignal: undefined });

    const collected = drain(stream);
    harness.emit({ type: 'error', errorText: 'No API key is configured for deepseek.' });

    const seen = (await collected) as ChatChunk[];
    expect(seen).toHaveLength(1);
    expect(seen[0]?.['errorText']).toBe('No API key is configured for deepseek.');
  });

  it('closes the port when the abort signal fires — Stop must reach main', async () => {
    const controller = new AbortController();
    const transport = new IpcChatTransport({ model: () => MODEL });
    const stream = await transport.sendMessages({
      messages: messages(),
      abortSignal: controller.signal,
    });

    const collected = drain(stream);
    harness.emit({ type: 'start' });
    expect(harness.stopped).toBe(0);

    controller.abort();
    expect(harness.stopped).toBe(1);
    expect(await collected).toHaveLength(1);
  });

  it('closes immediately when handed an already-aborted signal', async () => {
    const controller = new AbortController();
    controller.abort();
    const transport = new IpcChatTransport({ model: () => MODEL });
    const stream = await transport.sendMessages({
      messages: messages(),
      abortSignal: controller.signal,
    });

    expect(harness.stopped).toBe(1);
    expect(await drain(stream)).toEqual([]);
  });

  it('reads the model at send time, not at construction time', async () => {
    let current: ModelChoice = MODEL;
    const transport = new IpcChatTransport({ model: () => current });

    await transport.sendMessages({ messages: messages(), abortSignal: undefined });
    expect(harness.request?.model.model).toBe('deepseek-v4-flash');

    current = { provider: 'deepseek', model: 'deepseek-v4-pro' };
    await transport.sendMessages({ messages: messages(), abortSignal: undefined });
    expect(harness.request?.model.model).toBe('deepseek-v4-pro');
  });

  it('passes the system prompt through when one is supplied', async () => {
    const transport = new IpcChatTransport({ model: () => MODEL, system: () => 'be brief' });
    await transport.sendMessages({ messages: messages(), abortSignal: undefined });
    expect(harness.request?.system).toBe('be brief');
  });

  it('has nothing to reconnect to — a turn dies with its port', async () => {
    const transport = new IpcChatTransport({ model: () => MODEL });
    expect(await transport.reconnectToStream()).toBeNull();
  });
});
