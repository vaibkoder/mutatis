import type { Block } from './model.js';

/** One paragraph per line, as the engine's own plain-text synthesiser does. */
export function plainToBlocks(text: string): Block[] {
  const lines = text.replace(/\r\n?/g, '\n').split('\n');
  // A terminal newline terminates the last line; it is not an extra blank one.
  if (lines.length > 1 && lines[lines.length - 1] === '') lines.pop();
  return lines.map((line) => ({
    kind: 'p' as const,
    runs: line === '' ? [] : [{ text: line }],
  }));
}
