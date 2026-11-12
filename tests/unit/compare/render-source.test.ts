/**
 * RenderCore feature 16, Phase 2 B1 — the driver surfaces a read-only render
 * source (`authoredBodyXml` + styles + footnotes + numbering + the w:id→class
 * map) for both Word and Text compares, WITHOUT perturbing the editable-redline
 * (materialise / openRedline) paths.
 *
 * Spec: spec/features/16-render-core.md §6 Phase 2.
 */
import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { setActiveHost, _resetActiveHost } from '@shared/host/index.js';
import type { HostPort } from '@shared/host/index.js';
import { PizZipBundle } from '@shared/ooxml/bundle.js';
import {
  compareDriverState,
  textCompareDriverState,
  compareRenderSource,
  textRenderSource,
} from '@features/compare/driver.js';
import { makeAuthorContext, authorBody } from '@features/compare/output/index.js';
import type { DiffOp } from '@features/compare/model.js';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** Build a minimal but real .docx (zip) File from a body's run text. */
function makeDocxFile(name: string, bodyText: string): File {
  const bundle = new PizZipBundle();
  bundle.setPart(
    'word/document.xml',
    `<w:document xmlns:w="${W_NS}"><w:body>` +
      `<w:p><w:r><w:t>${bodyText}</w:t></w:r></w:p>` +
      '</w:body></w:document>',
  );
  bundle.setPart(
    'word/styles.xml',
    `<w:styles xmlns:w="${W_NS}"><w:docDefaults/></w:styles>`,
  );
  return new File([new Uint8Array(bundle.toUint8Array())], name);
}

let writeBack: ReturnType<typeof vi.fn>;

beforeEach(() => {
  writeBack = vi.fn().mockResolvedValue(undefined);
  setActiveHost({ writeBack } as unknown as HostPort);
  compareRenderSource.value = null;
  textRenderSource.value = null;
});

afterEach(() => {
  _resetActiveHost();
  vi.restoreAllMocks();
});

describe('B1 — render source (Word mode)', () => {
  it('carries a non-empty authoredBodyXml + styles after a Word compare', async () => {
    const d = compareDriverState;
    d.uploadDocx('original', makeDocxFile('orig.docx', 'The Seller shall deliver the Goods.'));
    d.uploadDocx('revised', makeDocxFile('rev.docx', 'The Seller may deliver the Goods.'));

    await d.runCompare();

    expect(d.status.value).toBe('done');
    const src = compareRenderSource.value;
    expect(src).not.toBeNull();
    expect(src!.authoredBodyXml.length).toBeGreaterThan(0);
    expect(src!.authoredBodyXml).toContain('<w:body>');
    expect(src!.stylesXml ?? '').toContain('<w:styles');
    // The replace produced ins/del revisions → at least one class entry.
    expect(src!.classByRevId).toBeDefined();
    expect(src!.classByRevId!.size).toBeGreaterThan(0);
    // The source bundle is carried for embedded-font extraction (B2).
    expect(src!.bundle.part('word/document.xml')).toBeDefined();
  });
});

describe('B1 — render source (Text mode)', () => {
  it('carries a non-empty authoredBodyXml + synthesised styles after a Text compare', async () => {
    const d = textCompareDriverState;
    d.originalText.value = 'The Seller shall deliver the Goods.';
    d.revisedText.value = 'The Seller may deliver the Goods.';

    await d.runCompare();

    expect(d.status.value).toBe('done');
    const src = textRenderSource.value;
    expect(src).not.toBeNull();
    expect(src!.authoredBodyXml).toContain('<w:body>');
    // Text mode synthesises a default-styles package (acceptance 11 source): the
    // engine renders the redline in the synthesised Calibri 11 default style.
    const styles = src!.stylesXml ?? '';
    expect(styles).toContain('<w:styles');
    expect(styles).toContain('w:ascii="Calibri"');
    expect(styles).toContain('<w:sz w:val="22"/>'); // 11pt
    // Eager-generate / click-to-open is untouched: no materialise on Compare.
    expect(writeBack).not.toHaveBeenCalled();
  });
});

describe('B1 — class capture is a pure side channel (authorBody byte-identical)', () => {
  it('captures the same XML with and without captureClasses', () => {
    const ops: DiffOp[] = [
      {
        id: 'o1', type: 'insert', origRange: null, revRange: { start: 0, end: 0 },
        partId: 'document', classification: 'substantiveText', classes: ['substantiveText'],
        author: 'A', date: '2026-01-01T00:00:00Z',
      },
    ];
    const revised = [
      {
        index: 0, partId: 'document', text: 'Hello world.',
        rawXml: '<w:p><w:r><w:t>Hello world.</w:t></w:r></w:p>',
        numberingPrefix: '', styleId: null, isHeadingCandidate: false,
      },
    ] as unknown as Parameters<typeof authorBody>[1];

    const plain = authorBody(ops, [], revised, makeAuthorContext('A', '2026-01-01T00:00:00Z'));
    const cap = makeAuthorContext('A', '2026-01-01T00:00:00Z', 1, { captureClasses: true });
    const captured = authorBody(ops, [], revised, cap);

    expect(captured).toBe(plain); // not one byte different
    expect(cap.classByRevId!.size).toBeGreaterThan(0);
    expect([...cap.classByRevId!.values()][0]).toEqual(['substantiveText']);
  });
});
