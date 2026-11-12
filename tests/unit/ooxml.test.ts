import { describe, it, expect } from 'vitest';
import { PizZipBundle, detectArtifactParagraph, toReadDocxError } from '@shared/ooxml/index.js';

const DOC_XML = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p w:rsidR="00AB12CD" w:rsidRDefault="00AB12CD"><w:r><w:t>Hello</w:t></w:r></w:p>
  </w:body>
</w:document>`;

describe('OOXML DocxBundle (PizZip-backed)', () => {
  it('round-trips part / setPart / hasPart / removePart', () => {
    const bundle = new PizZipBundle();
    expect(bundle.hasPart('word/document.xml')).toBe(false);

    bundle.setPart('word/document.xml', DOC_XML);
    expect(bundle.hasPart('word/document.xml')).toBe(true);
    expect(bundle.part('word/document.xml')).toContain('Hello');
    expect(bundle.listParts()).toContain('word/document.xml');

    bundle.removePart('word/document.xml');
    expect(bundle.hasPart('word/document.xml')).toBe(false);
  });

  it('strips named attributes (e.g. w:rsid*) preserving other content', () => {
    const bundle = new PizZipBundle();
    bundle.setPart('word/document.xml', DOC_XML);
    bundle.stripAttributes('word/document.xml', ['w:rsidR', 'w:rsidRDefault']);
    const out = bundle.part('word/document.xml')!;
    expect(out).not.toContain('w:rsidR');
    expect(out).not.toContain('w:rsidRDefault');
    // Content survives the strip
    expect(out).toContain('<w:t>Hello</w:t>');
  });

  it('serialises to a non-empty base64 package', () => {
    const bundle = new PizZipBundle();
    bundle.setPart('word/document.xml', DOC_XML);
    const b64 = bundle.toBase64();
    expect(typeof b64).toBe('string');
    expect(b64.length).toBeGreaterThan(0);

    // Re-open the serialised bytes and confirm the part survives a round-trip.
    const reopened = new PizZipBundle(bundle.toUint8Array());
    expect(reopened.part('word/document.xml')).toContain('Hello');
  });
});

describe('detectArtifactParagraph', () => {
  it('returns null when no artifact', () => {
    expect(detectArtifactParagraph([{ text: 'Hello' }, { text: 'World' }])).toBeNull();
  });

  it('detects trailing empty paragraph', () => {
    expect(detectArtifactParagraph([{ text: 'Hello' }, { text: '' }])).toBe('trailing');
  });

  it('detects leading empty paragraph (multi-para)', () => {
    expect(detectArtifactParagraph([{ text: '' }, { text: 'Hello' }])).toBe('leading');
  });

  it('prefers trailing over leading detection', () => {
    // Both first and last are empty — trailing takes priority
    expect(detectArtifactParagraph([{ text: '' }, { text: 'Hello' }, { text: '' }])).toBe('trailing');
  });

  it('returns null for single-item empty list', () => {
    expect(detectArtifactParagraph([{ text: '' }])).toBeNull();
  });

  it('returns null for empty list', () => {
    expect(detectArtifactParagraph([])).toBeNull();
  });
});

describe('toReadDocxError — getFileAsync failure mapping', () => {
  // Word on the web cannot read the full compressed .docx; the host surfaces an
  // ApiNotFound code (either in result.error or as a synchronous throw). Map it
  // to a desktop-Word message rather than leaking the raw code.
  it('maps an ApiNotFound code to the desktop-Word message', () => {
    const err = toReadDocxError('ApiNotFound', 'The API is not available.');
    expect(err.message).toContain('desktop');
    expect(err.message).toContain('web');
  });

  it('matches ApiNotFound case-insensitively, in code or message', () => {
    expect(toReadDocxError('apinotfound', undefined).message).toContain('desktop');
    expect(toReadDocxError(undefined, 'Error: apiNotFound thrown').message).toContain('desktop');
  });

  it('passes a non-ApiNotFound message through unchanged', () => {
    expect(toReadDocxError(5001, 'Disk full').message).toBe('Disk full');
  });

  it('falls back to a default message when none is supplied', () => {
    expect(toReadDocxError(undefined, undefined).message).toBe('getFileAsync failed');
  });
});
