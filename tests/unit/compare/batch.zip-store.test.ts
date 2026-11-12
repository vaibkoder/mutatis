/**
 * Compare at scale — the STORE-only ZIP writer (ADR-0037; spec §5.3).
 *
 * The archive is assembled from the artefacts' own `Blob`s so its bytes never
 * enter the JS heap — the same reason the runner materialises artefacts as
 * `Blob`s in the first place. That makes correctness entirely a question of the
 * headers, so these tests read the archive back with PizZip, which is the
 * library the rest of the product uses to open `.docx` packages: if PizZip can
 * open it, Word and Explorer can.
 */
import { describe, it, expect } from 'vitest';
import PizZip from 'pizzip';

import {
  writeStoreZip,
  safeEntryName,
  crc32,
  ZipEntry,
} from '@features/compare/batch/zip-store.js';

const enc = new TextEncoder();

function entry(name: string, text: string): ZipEntry {
  return { name, blob: new Blob([enc.encode(text)]) };
}

async function open(blob: Blob): Promise<PizZip> {
  return new PizZip(new Uint8Array(await blob.arrayBuffer()));
}

describe('crc32', () => {
  it('matches the known IEEE 802.3 vector', () => {
    // "123456789" -> 0xCBF43926 is the standard CRC-32 check value.
    expect(crc32(enc.encode('123456789'))).toBe(0xcbf4_3926);
  });

  it('is 0 for empty input', () => {
    expect(crc32(new Uint8Array(0))).toBe(0);
  });
});

describe('writeStoreZip', () => {
  it('produces an archive PizZip can open, with every entry intact', async () => {
    const zip = await open(
      await writeStoreZip([
        entry('nda-acme.docx', 'first redline'),
        entry('nda-borealis.docx', 'second redline'),
        entry('nda-cygnus.docx', 'third redline'),
      ]),
    );

    expect(Object.keys(zip.files).sort()).toEqual([
      'nda-acme.docx',
      'nda-borealis.docx',
      'nda-cygnus.docx',
    ]);
    expect(zip.file('nda-acme.docx')?.asText()).toBe('first redline');
    expect(zip.file('nda-borealis.docx')?.asText()).toBe('second redline');
    expect(zip.file('nda-cygnus.docx')?.asText()).toBe('third redline');
  });

  it('round-trips bytes exactly, including non-ASCII content', async () => {
    const bytes = new Uint8Array([0x00, 0xff, 0x7f, 0x80, 0x50, 0x4b, 0x03, 0x04]);
    const zip = await open(
      await writeStoreZip([{ name: 'raw.bin', blob: new Blob([bytes]) }]),
    );
    const out = zip.file('raw.bin')?.asUint8Array();
    expect(Array.from(out ?? [])).toEqual(Array.from(bytes));
  });

  it('handles an empty archive', async () => {
    const zip = await open(await writeStoreZip([]));
    expect(Object.keys(zip.files)).toEqual([]);
  });

  it('handles an empty entry', async () => {
    const zip = await open(await writeStoreZip([entry('empty.txt', '')]));
    expect(zip.file('empty.txt')?.asText()).toBe('');
  });

  it('is byte-identical across runs — no wall-clock timestamp leaks in', async () => {
    const build = (): Promise<Blob> =>
      writeStoreZip([entry('a.docx', 'same'), entry('b.docx', 'same')]);
    const [one, two] = await Promise.all([build(), build()]);
    const [a, b] = await Promise.all([one.arrayBuffer(), two.arrayBuffer()]);
    expect(Array.from(new Uint8Array(a))).toEqual(Array.from(new Uint8Array(b)));
  });

  it('de-duplicates colliding names rather than losing a redline', async () => {
    // Two bundles can easily each contain an "SPA.docx". Most extractors keep
    // only one entry of a duplicated name, so a collision silently costs the
    // user a comparison.
    const zip = await open(
      await writeStoreZip([
        entry('SPA.docx', 'first'),
        entry('SPA.docx', 'second'),
        entry('SPA.docx', 'third'),
      ]),
    );

    expect(Object.keys(zip.files).sort()).toEqual([
      'SPA (2).docx',
      'SPA (3).docx',
      'SPA.docx',
    ]);
    expect(zip.file('SPA.docx')?.asText()).toBe('first');
    expect(zip.file('SPA (2).docx')?.asText()).toBe('second');
    expect(zip.file('SPA (3).docx')?.asText()).toBe('third');
  });

  it('consumes an async iterable, so an entry can be produced lazily', async () => {
    const produced: string[] = [];
    async function* lazy(): AsyncGenerator<ZipEntry> {
      for (const name of ['one.docx', 'two.docx']) {
        produced.push(name);
        yield entry(name, name);
      }
    }

    const zip = await open(await writeStoreZip(lazy()));

    expect(produced).toEqual(['one.docx', 'two.docx']);
    expect(zip.file('two.docx')?.asText()).toBe('two.docx');
  });
});

describe('safeEntryName', () => {
  it('replaces path separators and characters filesystems reject', () => {
    const used = new Set<string>();
    expect(safeEntryName('deal/SPA:v7?.docx', used)).toBe('deal_SPA_v7_.docx');
  });

  it('leaves spaces and hyphens alone — they are legitimate, and the name is the only link back to the source document', () => {
    expect(safeEntryName('Project Falcon SPA-v7.docx', new Set())).toBe(
      'Project Falcon SPA-v7.docx',
    );
  });

  it('strips control characters, which no filesystem accepts', () => {
    const name = 'a' + String.fromCharCode(0) + 'b' + String.fromCharCode(31) + 'c.docx';
    expect(safeEntryName(name, new Set())).toBe('a_b_c.docx');
  });

  it('escapes a backslash as well as a forward slash, so no entry can climb out of the archive', () => {
    // Both halves of the traversal are neutralised: the separator becomes `_`,
    // and the leading `..` is collapsed by the dotfile guard.
    expect(safeEntryName('..' + String.fromCharCode(92) + 'evil.docx', new Set()))
      .toBe('__evil.docx');
  });

  it('refuses to produce a dotfile or an empty name', () => {
    expect(safeEntryName('...', new Set())).toBe('_');
    expect(safeEntryName('', new Set())).toBe('document');
  });

  it('suffixes before the extension, not after it', () => {
    const used = new Set(['redline.docx']);
    expect(safeEntryName('redline.docx', used)).toBe('redline (2).docx');
  });

  it('keeps counting past the first collision', () => {
    const used = new Set<string>();
    expect(safeEntryName('a.docx', used)).toBe('a.docx');
    expect(safeEntryName('a.docx', used)).toBe('a (2).docx');
    expect(safeEntryName('a.docx', used)).toBe('a (3).docx');
  });
});
