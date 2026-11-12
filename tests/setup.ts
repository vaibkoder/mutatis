/**
 * Vitest setup. jsdom provides localStorage and DOM; we add minimal Office/Word
 * global stubs for modules that reference them. Tests that exercise Office
 * behaviour install their own richer mocks.
 */
import { beforeEach } from 'vitest';

// jsdom does not implement Blob.prototype.arrayBuffer (real browsers do). The
// host-web file-picker path uses File.arrayBuffer(); back it with FileReader.
if (typeof Blob !== 'undefined' && typeof Blob.prototype.arrayBuffer !== 'function') {
  Blob.prototype.arrayBuffer = function arrayBuffer(this: Blob): Promise<ArrayBuffer> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result as ArrayBuffer);
      reader.onerror = () => reject(reader.error ?? new Error('FileReader failed'));
      reader.readAsArrayBuffer(this);
    });
  };
}

beforeEach(() => {
  localStorage.clear();
});
