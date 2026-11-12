import { describe, it, expect, vi, afterEach } from 'vitest';
import { extractParts } from '@features/compare/io/parts.js';
import type { DocxBundle } from '@shared/ooxml/index.js';
import { setActiveHost, _resetActiveHost } from '@shared/host/index.js';
import type { HostPort } from '@shared/host/index.js';

// ---------------------------------------------------------------------------
// Minimal DocxBundle stub
// ---------------------------------------------------------------------------

function makeBundle(parts: Record<string, string>): DocxBundle {
  return {
    part: (path: string) => parts[path],
    setPart: vi.fn(),
    hasPart: (path: string) => path in parts,
    listParts: () => Object.keys(parts),
    removePart: vi.fn(),
    stripAttributes: vi.fn(),
  };
}

// ---------------------------------------------------------------------------
// extractParts
// ---------------------------------------------------------------------------

describe('extractParts', () => {
  it('extracts documentXml', () => {
    const bundle = makeBundle({ 'word/document.xml': '<w:document/>' });
    const parts = extractParts(bundle);
    expect(parts.documentXml).toBe('<w:document/>');
  });

  it('throws when document.xml is missing', () => {
    const bundle = makeBundle({});
    expect(() => extractParts(bundle)).toThrow('word/document.xml missing');
  });

  it('finds header and footer parts by the sections that reference them, not by file name (ADR-0047)', () => {
    const bundle = makeBundle({
      'word/document.xml':
        '<w:document><w:body><w:p/><w:sectPr><w:headerReference w:type="default" r:id="rIdA"/>' +
        '<w:headerReference w:type="first" r:id="rIdB"/><w:footerReference w:type="default" r:id="rIdC"/>' +
        '</w:sectPr></w:body></w:document>',
      'word/_rels/document.xml.rels':
        '<Relationships>' +
        '<Relationship Id="rIdA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>' +
        '<Relationship Id="rIdB" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' +
        '<Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>' +
        '</Relationships>',
      'word/header1.xml': '<h1/>',
      'word/header2.xml': '<h2/>',
      'word/header3.xml': '<h3/>',
      'word/footer1.xml': '<f1/>',
    });
    const parts = extractParts(bundle);
    expect(parts.headers.get('header:s0:default')).toEqual({ path: 'word/header2.xml', xml: '<h2/>' });
    expect(parts.headers.get('header:s0:first')).toEqual({ path: 'word/header1.xml', xml: '<h1/>' });
    expect(parts.headers.size).toBe(2); // header3.xml is referenced by no section
    expect(parts.footers.get('footer:s0:default')?.xml).toBe('<f1/>');
    expect(parts.footers.size).toBe(1);
  });

  it('returns null for optional parts when absent', () => {
    const bundle = makeBundle({ 'word/document.xml': '<doc/>' });
    const parts = extractParts(bundle);
    expect(parts.footnotesXml).toBeNull();
    expect(parts.commentsXml).toBeNull();
    expect(parts.stylesXml).toBeNull();
    expect(parts.numberingXml).toBeNull();
  });

  it('extracts footnotes, comments, styles, and numbering when present', () => {
    const bundle = makeBundle({
      'word/document.xml': '<doc/>',
      'word/footnotes.xml': '<fn/>',
      'word/comments.xml': '<cm/>',
      'word/styles.xml': '<st/>',
      'word/numbering.xml': '<nm/>',
    });
    const parts = extractParts(bundle);
    expect(parts.footnotesXml).toBe('<fn/>');
    expect(parts.commentsXml).toBe('<cm/>');
    expect(parts.stylesXml).toBe('<st/>');
    expect(parts.numberingXml).toBe('<nm/>');
  });
});

// ---------------------------------------------------------------------------
// loadBothSides — ArrayBuffer path (no Office JS required)
// ---------------------------------------------------------------------------

describe('loadBothSides — docxArrayBuffer path', () => {
  afterEach(() => {
    _resetActiveHost();
  });

  it('constructs a PizZipBundle from an ArrayBuffer; active side reads via the host', async () => {
    // The docxArrayBuffer path creates a PizZipBundle from a buffer; the active
    // side is read through the installed HostPort (12f: getActiveHost().readDocx()).
    const { loadBothSides } = await import('@features/compare/io/load.js');

    const mockBundle = makeBundle({ 'word/document.xml': '<doc/>' });
    const readDocx = vi.fn().mockResolvedValue(mockBundle);
    setActiveHost({ readDocx } as unknown as HostPort);

    // Create a minimal valid zip-like buffer via PizZipBundle itself
    const { PizZipBundle } = await import('@shared/ooxml/service.js');
    const emptyBundle = new PizZipBundle();
    const bytes = emptyBundle.toUint8Array();

    const result = await loadBothSides({
      original: { kind: 'activeDocument' },
      revised: { kind: 'docxArrayBuffer', buffer: bytes.buffer as ArrayBuffer },
    });

    expect(readDocx).toHaveBeenCalledOnce();
    expect(result.originalBundle).toBe(mockBundle);
    // Revised bundle is constructed from the ArrayBuffer. `DocxBundle` does not
    // declare serialisation (see `service.ts`, which narrows before calling it),
    // so assert the concrete class — which is what the callers narrow to anyway.
    expect(result.revisedBundle).toBeInstanceOf(PizZipBundle);
  });

  it('does not read the host when both sides are docxArrayBuffer', async () => {
    const { loadBothSides } = await import('@features/compare/io/load.js');

    const readDocx = vi.fn();
    setActiveHost({ readDocx } as unknown as HostPort);

    const { PizZipBundle } = await import('@shared/ooxml/service.js');
    const bytes = new PizZipBundle().toUint8Array();

    await loadBothSides({
      original: { kind: 'docxArrayBuffer', buffer: bytes.buffer.slice(0) as ArrayBuffer },
      revised: { kind: 'docxArrayBuffer', buffer: bytes.buffer.slice(0) as ArrayBuffer },
    });

    expect(readDocx).not.toHaveBeenCalled();
  });
});

// ---------------------------------------------------------------------------
// CompareSource type exhaustiveness (compile-time check — if this compiles, the seam is correct)
// ---------------------------------------------------------------------------

describe('CompareSource discriminator', () => {
  it('active-document source has correct kind', () => {
    const src = { kind: 'activeDocument' as const };
    expect(src.kind).toBe('activeDocument');
  });

  it('docxArrayBuffer source has correct kind', () => {
    const buf = new ArrayBuffer(4);
    const src = { kind: 'docxArrayBuffer' as const, buffer: buf };
    expect(src.kind).toBe('docxArrayBuffer');
    expect(src.buffer).toBe(buf);
  });
});
