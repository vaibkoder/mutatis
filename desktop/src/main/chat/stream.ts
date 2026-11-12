import { streamText, convertToModelMessages, toUIMessageStream, type UIMessage } from 'ai';
import type { MessagePortMain } from 'electron';
import { resolveModel } from './providers.js';
import { EDIT_TOOLS } from './tools.js';
import type { ChatTurnRequest } from '../../common/ipc.js';

/**
 * Run one chat turn and stream it back over a MessagePort.
 *
 * A port per turn rather than a shared channel, because closing the port IS the
 * stop signal — there is no separate cancel message to lose, and a renderer that
 * navigates away or crashes aborts the upstream HTTP request automatically.
 *
 * What crosses the port is the AI SDK's own `UIMessageChunk` shape, so the
 * renderer transport can hand it to `useChat` untouched. Only plain JSON goes
 * over IPC — no SDK objects, no callbacks, no AbortSignal, nothing that
 * structured clone would reject.
 */

export type { ChatTurnRequest };

export function runChatTurn(port: MessagePortMain, request: ChatTurnRequest): void {
  const controller = new AbortController();
  let closed = false;

  const send = (chunk: unknown): void => {
    if (closed) return;
    try {
      port.postMessage(chunk);
    } catch {
      // The renderer went away mid-stream. Stop rather than throwing into an
      // unhandled rejection.
      closed = true;
      controller.abort();
    }
  };

  const finish = (): void => {
    if (closed) return;
    closed = true;
    try {
      port.close();
    } catch {
      /* already gone */
    }
  };

  // Port closure from the renderer side is Stop.
  port.on('close', () => {
    closed = true;
    controller.abort();
  });
  port.start();

  void (async () => {
    try {
      const model = resolveModel(request.model);

      const result = streamText({
        model,
        // Client-side tools: no `execute`, so the call is streamed out and the
        // renderer performs it. Main never opens a document.
        tools: EDIT_TOOLS,
        ...(request.system !== undefined ? { system: request.system } : {}),
        // Async in AI SDK v7 — it resolves file/image parts before the call.
        messages: await convertToModelMessages(request.messages as UIMessage[]),
        abortSignal: controller.signal,
      });

      // The UI message stream, not the text stream: a tool call is not text, and
      // `textStream` drops it silently. That failure mode is invisible — the
      // model asks to edit a document and the renderer simply never hears.
      const stream = toUIMessageStream({
        stream: result.fullStream,
        tools: EDIT_TOOLS,
        // Errors are surfaced by the catch below, which forwards only
        // `error.message`. The default here would hide them behind a generic
        // string; passing them through unchanged would risk a request dump.
        onError: (error) => (error instanceof Error ? error.message : String(error)),
      });

      const reader = stream.getReader();
      for (;;) {
        const { done, value } = await reader.read();
        if (done || closed) break;
        send(value);
      }
    } catch (error) {
      // An abort is the user pressing Stop, not a failure to report.
      const aborted =
        controller.signal.aborted ||
        (error instanceof Error && (error.name === 'AbortError' || error.name === 'TimeoutError'));
      if (!aborted) {
        send({
          type: 'error',
          // The message is shown to the user, so it must never carry the key.
          // Provider SDKs do not put the key in `message`, but they do put it in
          // request dumps — which is why only `.message` is forwarded.
          errorText: error instanceof Error ? error.message : String(error),
        });
      }
    } finally {
      finish();
    }
  })();
}
