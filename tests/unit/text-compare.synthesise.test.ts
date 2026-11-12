/**
 * Feature 12 (Compare — additional sources) T1: text→docx synthesiser.
 * Spec: spec/features/10-compare-additional-sources.md §4.1 · ADR-0018.
 *
 * Verifies the synthesised package is a VALID minimal .docx that:
 *  - round-trips through PizZip (re-zip → re-open),
 *  - carries the required OOXML parts + content types + relationships,
 *  - emits one <w:p> per line (TT4 fixture: 3-line blob → 3 paragraphs),
 *  - escapes XML-special characters without corrupting the package (TT4),
 *  - is extractable by the existing feature-08 `extractParts` / `preprocessSide`.
 */
import { describe, it, expect } from 'vitest';
import PizZip from 'pizzip';

import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import { PizZipBundle } from '@shared/ooxml/service.js';
import { extractParts } from '@features/compare/io/parts.js';
import { preprocessSide } from '@features/compare/preprocess/index.js';

import {
  synthesiseDocx,
  splitLines,
} from '@features/compare/io/text/synthesise-docx.js';

const CONFIG = getCompareConfig(SG_PROFILE);

// Count <w:p …> opening tags (covers both <w:p>…</w:p> and self-closing <w:p/>).
function countParagraphs(documentXml: string): number {
  return (documentXml.match(/<w:p(?:\s[^>]*)?(?:\/)?>/g) ?? []).length;
}

describe('splitLines — deterministic one-line-per-paragraph', () => {
  it('splits a 3-line blob into 3 lines', () => {
    expect(splitLines('a\nb\nc')).toEqual(['a', 'b', 'c']);
  });

  it('normalises CRLF and lone CR before splitting', () => {
    expect(splitLines('a\r\nb\rc')).toEqual(['a', 'b', 'c']);
  });

  it('drops a single terminal newline by default', () => {
    expect(splitLines('a\nb\n')).toEqual(['a', 'b']);
  });

  it('keeps interior blank lines by default', () => {
    expect(splitLines('a\n\nb')).toEqual(['a', '', 'b']);
  });

  it('drops blank lines when blankLines=drop', () => {
    expect(splitLines('a\n\n  \nb', { blankLines: 'drop' })).toEqual(['a', 'b']);
  });
});

describe('synthesiseDocx — valid minimal package', () => {
  it('contains every required OOXML part', () => {
    const bundle = synthesiseDocx('hello');
    for (const part of [
      '[Content_Types].xml',
      '_rels/.rels',
      'word/_rels/document.xml.rels',
      'word/document.xml',
      'word/styles.xml',
      'word/settings.xml',
    ]) {
      expect(bundle.hasPart(part)).toBe(true);
    }
  });

  it('declares content types + the officeDocument relationship', () => {
    const bundle = synthesiseDocx('hello');
    const ct = bundle.part('[Content_Types].xml')!;
    expect(ct).toContain('wordprocessingml.document.main+xml');
    const rels = bundle.part('_rels/.rels')!;
    expect(rels).toContain('relationships/officeDocument');
    expect(rels).toContain('Target="word/document.xml"');
  });

  it('document.xml root declares the w namespace and a single sectPr', () => {
    const doc = synthesiseDocx('hello').part('word/document.xml')!;
    expect(doc).toContain('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"');
    expect((doc.match(/<w:sectPr>/g) ?? []).length).toBe(1);
  });

  it('emits exactly one <w:p> per line (3-line blob → 3 paragraphs)', () => {
    const doc = synthesiseDocx('one\ntwo\nthree').part('word/document.xml')!;
    expect(countParagraphs(doc)).toBe(3);
  });

  it('round-trips through PizZip (re-zip then re-open) — TT4 validity', () => {
    const bytes = synthesiseDocx('alpha\nbeta').toUint8Array();
    // A corrupt package would throw here.
    const reopened = new PizZip(bytes);
    expect(reopened.file('word/document.xml')).not.toBeNull();
    expect(reopened.file('[Content_Types].xml')).not.toBeNull();
    // And re-wrap as a bundle to confirm parts are readable.
    const bundle = new PizZipBundle(bytes);
    expect(bundle.part('word/document.xml')).toContain('alpha');
  });

  it('escapes XML-special characters (&<>") without corrupting the package — TT4', () => {
    const tricky = 'A & B < C > D " E';
    const bundle = synthesiseDocx(tricky);
    const doc = bundle.part('word/document.xml')!;
    expect(doc).toContain('A &amp; B &lt; C &gt; D " E');
    expect(doc).not.toMatch(/A & B/); // raw ampersand would be invalid
    // Round-trip proves the package is still valid bytes.
    const reopened = new PizZip(bundle.toUint8Array());
    expect(reopened.file('word/document.xml')!.asText()).toContain('&amp;');
  });

  it('strips XML-1.0-illegal control characters from text', () => {
    // Bell (0x07) and vertical tab (0x0B) are illegal in XML 1.0 → stripped.
    const withControls = `a${String.fromCharCode(0x07)}b${String.fromCharCode(0x0b)}cd`;
    const doc = synthesiseDocx(withControls).part('word/document.xml')!;
    expect(doc).toContain('abcd');
    // Package stays valid bytes after the strip.
    expect(new PizZip(synthesiseDocx(withControls).toUint8Array()).file('word/document.xml')).not.toBeNull();
  });

  it('preserves leading/trailing whitespace via xml:space=preserve', () => {
    const doc = synthesiseDocx('  indented line  ').part('word/document.xml')!;
    expect(doc).toContain('xml:space="preserve"');
    expect(doc).toContain('  indented line  ');
  });

  it('empty lines become bare <w:p/> when kept', () => {
    const doc = synthesiseDocx('a\n\nb').part('word/document.xml')!;
    expect(doc).toContain('<w:p/>');
    expect(countParagraphs(doc)).toBe(3);
  });
});

describe('synthesiseDocx — feeds the feature-08 pipeline', () => {
  it('extractParts reads the synthesised document + styles parts', () => {
    const parts = extractParts(synthesiseDocx('one\ntwo'));
    expect(parts.documentXml).toContain('<w:body>');
    expect(parts.stylesXml).not.toBeNull();
  });

  it('preprocessSide produces one ParaRecord per non-empty line', () => {
    const parts = extractParts(synthesiseDocx('First clause.\nSecond clause.'));
    const result = preprocessSide(parts.documentXml, 'document', 'acceptAll', CONFIG);
    const texts = result.paras.map(p => p.normText);
    expect(texts).toContain('First clause.');
    expect(texts).toContain('Second clause.');
  });
});
