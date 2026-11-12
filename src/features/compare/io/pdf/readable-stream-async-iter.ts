/**
 * WebKit shim — async iteration of `ReadableStream`.
 *
 * PDF.js's `getTextContent()` consumes its result stream with
 * `for await (const value of readableStream)` (pdf.mjs `getTextContent`), which
 * relies on `ReadableStream.prototype[Symbol.asyncIterator]`. Chromium and Gecko
 * ship it; **WebKit/Safari does not** — and Office-on-the-web on macOS runs the
 * task pane in WebKit. There, the `for await` throws
 * `undefined is not a function (near '...value of readableStream...')` and the
 * whole PDF compare fails. Windows Word / Edge (Chromium) are unaffected.
 *
 * This installs the spec-shaped async iterator only when it is genuinely absent,
 * delegating to the always-supported `getReader()`. It is purely client-side
 * (no network, no Office JS) and idempotent. The loader calls it before any
 * PDF.js work; tests inject their own opener and never import this file.
 */

interface AsyncIterableReadableStream {
  getReader(): {
    read(): Promise<{ done: boolean; value: unknown }>;
    cancel(reason?: unknown): Promise<void>;
    releaseLock(): void;
  };
}

/**
 * Ensure `ReadableStream` supports `for await … of` (and `.values()`), adding a
 * `getReader()`-backed async iterator only on engines that lack the native one.
 */
export function ensureReadableStreamAsyncIteration(): void {
  const RS = (globalThis as { ReadableStream?: { prototype: object } }).ReadableStream;
  if (!RS) return; // No ReadableStream at all ⇒ PDF.js cannot run here regardless.

  const proto = RS.prototype as Record<string | symbol, unknown>;
  if (typeof proto[Symbol.asyncIterator] === 'function') return; // Native support — leave it.

  function asyncIterator(this: AsyncIterableReadableStream) {
    const reader = this.getReader();
    return {
      next(): Promise<IteratorResult<unknown>> {
        return reader.read().then(({ done, value }) =>
          done ? { done: true, value: undefined } : { done: false, value },
        );
      },
      return(value?: unknown): Promise<IteratorResult<unknown>> {
        reader.releaseLock();
        return Promise.resolve({ done: true, value });
      },
      [Symbol.asyncIterator](): unknown {
        return this;
      },
    };
  }

  // Non-enumerable, matching native prototype-method semantics.
  Object.defineProperty(proto, Symbol.asyncIterator, {
    value: asyncIterator,
    writable: true,
    configurable: true,
  });
  if (typeof proto.values !== 'function') {
    Object.defineProperty(proto, 'values', {
      value: asyncIterator,
      writable: true,
      configurable: true,
    });
  }
}
