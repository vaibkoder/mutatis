import { describe, it, expect } from 'vitest';
import {
  acceptFiles,
  comparisonShape,
  emptyDocumentSet,
  MAX_BUNDLE_DOCUMENTS,
  type LoadedDoc,
} from '../src/renderer/src/state/documents.js';
import { MAX_DOCX_BYTES } from '@features/compare/batch/limits.js';

/**
 * These tests double as the first proof that the sibling path alias into
 * Mutatis resolves in the TEST runner, not just in the bundler and tsc.
 * All three have to work for phase 2, and they fail independently.
 */

function fakeFile(name: string, size = 1024): File {
  // Only name and size are read at this stage; bytes stay unread until a
  // comparison actually runs.
  return { name, size } as File;
}

describe('acceptFiles', () => {
  it('accepts the four supported formats', () => {
    const { accepted, rejected } = acceptFiles(
      [fakeFile('a.docx'), fakeFile('b.pdf'), fakeFile('c.txt'), fakeFile('d.md')],
      [],
    );
    expect(accepted.map((d) => d.kind)).toEqual(['docx', 'pdf', 'text', 'text']);
    expect(rejected).toHaveLength(0);
  });

  it('rejects unsupported formats with a reason, never silently', () => {
    const { accepted, rejected } = acceptFiles([fakeFile('notes.pages')], []);
    expect(accepted).toHaveLength(0);
    expect(rejected).toHaveLength(1);
    expect(rejected[0]?.reason).toContain('.pages');
  });

  it('names the problem when a file has no extension at all', () => {
    const { rejected } = acceptFiles([fakeFile('README')], []);
    expect(rejected[0]?.reason).toContain('no file extension');
  });

  it('rejects a duplicate of something already on this side', () => {
    const first = acceptFiles([fakeFile('deed.docx', 2048)], []);
    const second = acceptFiles([fakeFile('deed.docx', 2048)], first.accepted);
    expect(second.accepted).toHaveLength(0);
    expect(second.rejected[0]?.reason).toContain('already added');
  });

  it('FLAGS an oversize .docx rather than dropping it', () => {
    // The engine's own cap helpers make the same choice: a document the user
    // handed us stays visible even when it cannot be compared. Dropping it
    // would be suppression by another name.
    const { accepted, rejected } = acceptFiles([fakeFile('huge.docx', MAX_DOCX_BYTES + 1)], []);
    expect(rejected).toHaveLength(0);
    expect(accepted).toHaveLength(1);
    expect(accepted[0]?.overCap).toBe(true);
  });

  it('does not flag a .docx exactly at the cap', () => {
    const { accepted } = acceptFiles([fakeFile('exact.docx', MAX_DOCX_BYTES)], []);
    expect(accepted[0]?.overCap).toBe(false);
  });

  it('refuses documents past the per-side limit, and says which limit', () => {
    const existing = Array.from({ length: MAX_BUNDLE_DOCUMENTS }, (_, i) => ({
      id: `x${i}`,
      name: `f${i}.docx`,
      byteLength: 10,
      kind: 'docx' as const,
      file: fakeFile(`f${i}.docx`),
      overCap: false,
    })) satisfies LoadedDoc[];

    const { accepted, rejected } = acceptFiles([fakeFile('one-too-many.docx')], existing);
    expect(accepted).toHaveLength(0);
    expect(rejected[0]?.reason).toContain(String(MAX_BUNDLE_DOCUMENTS));
  });
});

describe('comparisonShape', () => {
  const doc = (n: string): LoadedDoc => ({
    id: n,
    name: n,
    byteLength: 1,
    kind: 'docx',
    file: fakeFile(n),
    overCap: false,
  });

  it('needs a document before it is anything; a lone one means its own tracked changes', () => {
    expect(comparisonShape(emptyDocumentSet)).toBe('none');
    expect(comparisonShape({ original: [doc('a')], modified: [] })).toBe('tracked-changes');
    // Several documents on one side is still not a comparison.
    expect(comparisonShape({ original: [doc('a'), doc('b')], modified: [] })).toBe('none');
  });

  it('classifies 1:1, 1:N and N:M', () => {
    expect(comparisonShape({ original: [doc('a')], modified: [doc('b')] })).toBe('one-to-one');
    expect(comparisonShape({ original: [doc('a')], modified: [doc('b'), doc('c')] })).toBe(
      'one-to-many',
    );
    expect(comparisonShape({ original: [doc('a'), doc('b')], modified: [doc('c')] })).toBe(
      'one-to-many',
    );
    expect(
      comparisonShape({ original: [doc('a'), doc('b')], modified: [doc('c'), doc('d')] }),
    ).toBe('many-to-many');
  });
});
