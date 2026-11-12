import type { ChatTransport, UIMessage, UIMessageChunk } from 'ai';
import { PORT_CLOSED, type ChatChunk, type ModelChoice } from '@common/ipc.js';

/**
 * THE WRAPPER FILE.
 *
 * This is the only place in the renderer that knows a chat turn leaves the
 * process. Everything above it uses `useChat` exactly as the AI SDK documents
 * it; everything below it is IPC.
 *
 * It exists because `DirectChatTransport` — the SDK's own in-process transport —
 * calls `agent.stream()` on an SDK object it holds a reference to. That object
 * cannot cross an IPC boundary (structured clone would reject it), so using it
 * would mean constructing the provider client HERE, which would put the API key
 * and the provider HTTP call in the renderer. Both are forbidden. The transport
 * interface is two methods; implementing it costs far less than the alternative.
 *
 * `HttpChatTransport` is equally unavailable: there is no HTTP server to point
 * it at, and the CSP is `connect-src 'none'`.
 */

export interface IpcTransportOptions {
  /** Read at send time, not construction time, so a model change takes effect immediately. */
  readonly model: () => ModelChoice;
  readonly system?: () => string | undefined;
}

export class IpcChatTransport implements ChatTransport<UIMessage> {
  constructor(private readonly options: IpcTransportOptions) {}

  sendMessages(request: {
    messages: UIMessage[];
    abortSignal: AbortSignal | undefined;
  }): Promise<ReadableStream<UIMessageChunk>> {
    const { messages, abortSignal } = request;
    const model = this.options.model();
    const system = this.options.system?.();

    const stream = new ReadableStream<UIMessageChunk>({
      start: (controller) => {
        let settled = false;
        let stop: (() => void) | null = null;

        const settle = (): void => {
          if (settled) return;
          settled = true;
          try {
            controller.close();
          } catch {
            /* already closed */
          }
          // Close the port too. Without this, main keeps streaming into a
          // ReadableStream nobody is reading — and on an error chunk the
          // upstream HTTP request would run to completion for no reason.
          stop?.();
        };

        stop = window.mutatis.startChat(
          {
            // Structured clone gets a plain array of plain objects. UIMessages
            // are JSON by construction, but the spread makes that explicit
            // rather than relying on it.
            messages: messages.map((message) => ({ ...message })),
            model,
            ...(system !== undefined ? { system } : {}),
          },
          (chunk: ChatChunk) => {
            if (settled) return;

            // The bridge's own sentinel, not an SDK chunk. Main went away
            // without finishing; end the stream rather than hanging.
            if (chunk.type === PORT_CLOSED) {
              settle();
              return;
            }

            controller.enqueue(chunk as unknown as UIMessageChunk);

            // `finish` and `error` are terminal. Main closes its end right
            // after, but we do not wait for that round trip to release the UI.
            if (chunk.type === 'finish' || chunk.type === 'error') settle();
          },
        );

        // `useChat`'s stop() aborts this signal. Closing the port is what
        // actually cancels the upstream HTTP request, in main.
        abortSignal?.addEventListener('abort', settle, { once: true });
        if (abortSignal?.aborted) settle();
      },
    });

    return Promise.resolve(stream);
  }

  /**
   * Always null. A turn's stream lives and dies with its MessagePort — there is
   * no server holding a resumable response, so there is nothing to reconnect to.
   * Returning null is the documented way to say "that response is already over".
   */
  reconnectToStream(): Promise<ReadableStream<UIMessageChunk> | null> {
    return Promise.resolve(null);
  }
}
