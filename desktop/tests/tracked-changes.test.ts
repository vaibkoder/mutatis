import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { runComparison } from '../src/renderer/src/compare/run-comparison.js';
import {
  acceptAll,
  hasTrackedChanges,
  NoTrackedChangesError,
  rejectAll,
  splitTrackedChanges,
} from '../src/renderer/src/compare/tracked-changes.js';
import { comparisonShape } from '../src/renderer/src/state/documents.js';
import { makeDocxFile } from './helpers/make-docx.js';
import type { LoadedDoc } from '../src/renderer/src/state/documents.js';

const docOf = (file: File, id = 'd'): LoadedDoc => ({
  id, name: file.name, byteLength: file.size, kind: 'docx', file, overCap: false,
});
const tracked = () => {
  const bytes = readFileSync(resolve(import.meta.dirname, 'fixtures/tracked.docx'));
  return docOf(new File([bytes as BlobPart], 'tracked.docx'));
};

describe('accepting and rejecting tracked changes', () => {
  const xml =
    '<w:p><w:r><w:t>A </w:t></w:r><w:del w:id="1"><w:r><w:delText>old</w:delText></w:r></w:del>' +
    '<w:ins w:id="2"><w:r><w:t>new</w:t></w:r></w:ins><w:tblBorders><w:insideH/></w:tblBorders></w:p>';
  it('accept keeps insertions, drops deletions', () => {
    const out = acceptAll(xml);
    expect(out).toContain('new');
    expect(out).not.toContain('old');
    expect(out).toContain('<w:insideH/>');
  });
  it('reject keeps deletions as ordinary text, drops insertions', () => {
    const out = rejectAll(xml);
    expect(out).toContain('<w:t>old</w:t>');
    expect(out).not.toContain('new');
  });
  it('detects tracked changes, ignoring paragraph-mark-only markers and table borders', () => {
    expect(hasTrackedChanges(xml)).toBe(true);
    expect(hasTrackedChanges('<w:p><w:pPr><w:rPr><w:ins w:id="1"/></w:rPr></w:pPr></w:p><w:insideH/>')).toBe(false);
  });
});

describe('one document with tracked changes', () => {
  it('a lone document is the tracked-changes shape, from either side', () => {
    const d = tracked();
    expect(comparisonShape({ original: [d], modified: [] })).toBe('tracked-changes');
    expect(comparisonShape({ original: [], modified: [d] })).toBe('tracked-changes');
    expect(comparisonShape({ original: [], modified: [] })).toBe('none');
  });

  it('compares the before text against the after text', async () => {
    const { original, modified } = await splitTrackedChanges(tracked());
    const out = await runComparison({
      originalBytes: await original.file.arrayBuffer(),
      revisedBytes: await modified.file.arrayBuffer(),
    });
    expect(out.summary.total).toBeGreaterThanOrEqual(3);
    const shown = JSON.stringify(out.model);
    for (const t of ['£40,000', '£55,000', 'wholly new clause']) expect(shown).toContain(t);
  });

  it('refuses a document with no tracked changes, saying why', async () => {
    const plain = docOf(makeDocxFile('plain.docx', ['Nothing tracked here.']));
    await expect(splitTrackedChanges(plain)).rejects.toBeInstanceOf(NoTrackedChangesError);
  });
});
