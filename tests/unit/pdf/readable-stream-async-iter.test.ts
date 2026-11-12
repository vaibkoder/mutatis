/**
 * WebKit shim — async iteration of ReadableStream.
 *
 * Reproduces the Office-on-the-web-on-macOS failure: a ReadableStream whose
 * prototype lacks [Symbol.asyncIterator] makes PDF.js's
 * `for await (const value of readableStream)` throw. The shim must add a
 * getReader()-backed iterator that yields the same chunks, be idempotent, and
 * never overwrite a native implementation.
 */
import { describe, it, expect, afterEach } from 'vitest';
import { ensureReadableStreamAsyncIteration } from '../../../src/features/compare/io/pdf/readable-stream-async-iter.js';

type AsyncIterKey = ReadableStream<unknown> & Record<symbol, unknown>;

/** A real ReadableStream emitting the given chunks then closing. */
function streamOf(chunks: unknown[]): ReadableStream {
  return new ReadableStream({
    start(controller) {
      for (const c of chunks) controller.enqueue(c);
      controller.close();
    },
  });
}

describe('ensureReadableStreamAsyncIteration', () => {
  const original = Object.getOwnPropertyDescriptor(ReadableStream.prototype, Symbol.asyncIterator);

  afterEach(() => {
    // Restore the real prototype between cases (we mutate the global below).
    if (original) Object.defineProperty(ReadableStream.prototype, Symbol.asyncIterator, original);
    else delete (ReadableStream.prototype as AsyncIterKey)[Symbol.asyncIterator];
  });

  it('adds a working async iterator when one is missing (the WebKit case)', async () => {
    // Simulate WebKit: strip the native async iterator off the prototype.
    delete (ReadableStream.prototype as AsyncIterKey)[Symbol.asyncIterator];
    expect((ReadableStream.prototype as AsyncIterKey)[Symbol.asyncIterator]).toBeUndefined();

    ensureReadableStreamAsyncIteration();
    expect(typeof (ReadableStream.prototype as AsyncIterKey)[Symbol.asyncIterator]).toBe('function');

    const seen: unknown[] = [];
    for await (const value of streamOf(['a', 'b', 'c']) as unknown as AsyncIterable<unknown>) seen.push(value);
    expect(seen).toEqual(['a', 'b', 'c']);
  });

  it('is idempotent and does not overwrite a native implementation', () => {
    function sentinel() { /* pretend-native */ }
    Object.defineProperty(ReadableStream.prototype, Symbol.asyncIterator, {
      value: sentinel,
      writable: true,
      configurable: true,
    });

    ensureReadableStreamAsyncIteration();
    ensureReadableStreamAsyncIteration();

    expect((ReadableStream.prototype as AsyncIterKey)[Symbol.asyncIterator]).toBe(sentinel);
  });
});
