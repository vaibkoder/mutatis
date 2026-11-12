/**
 * A2 — the `?bytes` plugin embeds binary assets as in-process Uint8Arrays
 * (no runtime fetch). Importing a known vendored font through `@fonts/...?bytes`
 * must yield the exact on-disk byte length (zero-egress asset delivery, ADR-0024).
 */
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import caladeaRegular from '@fonts/Caladea-Regular.ttf?bytes';

describe('?bytes plugin', () => {
  it('embeds a vendored font as a Uint8Array of the right length', () => {
    const onDisk = readFileSync(resolve(__dirname, '../../../public/fonts/Caladea-Regular.ttf'));
    expect(caladeaRegular).toBeInstanceOf(Uint8Array);
    expect(caladeaRegular.length).toBe(onDisk.length);
    // sfnt magic: 0x00010000 (TrueType) — proves real bytes, decoded correctly.
    expect(Array.from(caladeaRegular.slice(0, 4))).toEqual([0x00, 0x01, 0x00, 0x00]);
  });
});
