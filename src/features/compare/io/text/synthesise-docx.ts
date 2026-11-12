/**
 * Feature 12 (Compare — additional sources) T1: plain text → minimal valid .docx.
 * Spec: spec/features/10-compare-additional-sources.md §4.1 · ADR-0018.
 *
 * Synthesises a controlled, minimal-but-VALID OOXML package from a pasted plain-text
 * blob so it can ride the *existing* feature-08 Compare pipeline unchanged
 * (loadBothSides → runPipeline → output/author → materialise). One `<w:p>` per line,
 * text XML-escaped into a single run, a single trailing `<w:sectPr>`, plus the
 * styles/settings/rels parts Word needs to open the package and render revisions.
 *
 * Pure: no Office JS, no Date.now(), no Math.random(). The package is built entirely
 * in memory (PizZip); nothing leaves the machine (ADR-0001/0003).
 *
 * The one real engineering risk this file carries (per ADR-0018) is producing a
 * package Word opens cleanly AND renders tracked changes from — covered by TT4.
 */
import { PizZipBundle } from '@shared/ooxml/bundle.js';

// ---------------------------------------------------------------------------
// Options + named constants (pinned here, not magic literals — §7 config keys)
// ---------------------------------------------------------------------------

/** How blank lines in the pasted blob are represented in the synthesised body. */
export type BlankLinePolicy = 'keep' | 'drop';

export interface SynthesiseOptions {
  /**
   * Blank-line handling (deterministic). 'keep' (default) emits one empty `<w:p/>`
   * per blank line — the faithful representation (categorise-never-suppress spirit:
   * structure is not silently dropped). 'drop' omits blank lines entirely.
   */
  blankLines?: BlankLinePolicy;
  /**
   * Whether a single trailing empty line produced by a terminal newline
   * (e.g. "a\nb\n" → ["a","b",""]) is dropped. Default true — a terminal newline
   * is a line *terminator*, not an extra blank paragraph.
   */
  dropTrailingEmpty?: boolean;
}

const DEFAULT_OPTIONS: Required<SynthesiseOptions> = {
  blankLines: 'keep',
  dropTrailingEmpty: true,
};

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

// ---------------------------------------------------------------------------
// Line splitting (deterministic, one paragraph per line)
// ---------------------------------------------------------------------------

/**
 * Split a plain-text blob into lines, deterministically.
 *
 * - CRLF and lone CR are normalised to LF first, so the split never depends on
 *   the host's newline convention.
 * - A single terminal newline is treated as a line terminator (dropped) unless
 *   `dropTrailingEmpty` is false.
 * - Blank lines are kept or dropped per `blankLines`.
 *
 * Exported for direct unit testing of the split contract.
 */
export function splitLines(text: string, opts: SynthesiseOptions = {}): string[] {
  const o = { ...DEFAULT_OPTIONS, ...opts };
  const normalised = text.replace(/\r\n?/g, '\n');
  let lines = normalised.split('\n');

  // A terminal newline yields a trailing '' — drop exactly one when configured.
  if (o.dropTrailingEmpty && lines.length > 1 && lines[lines.length - 1] === '') {
    lines = lines.slice(0, -1);
  }

  if (o.blankLines === 'drop') {
    lines = lines.filter(line => line.trim() !== '');
  }

  return lines;
}

// ---------------------------------------------------------------------------
// XML escaping / sanitising
// ---------------------------------------------------------------------------

/**
 * Escape text for an XML element body (`<w:t>` content). Escapes the three
 * characters that are significant in element content; quotes are NOT escaped
 * (they are only significant inside attribute values).
 *
 * Also strips characters that are illegal in XML 1.0 (most control characters)
 * so a stray control byte in a paste cannot corrupt the package. Tab (0x09) is
 * allowed and preserved; newline/CR cannot appear here (lines are already split).
 */
export function escapeXmlText(text: string): string {
  const sanitised = text.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  return sanitised
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

// ---------------------------------------------------------------------------
// Part builders
// ---------------------------------------------------------------------------

/** Build one `<w:p>` for a line of text. Empty lines become a bare `<w:p/>`. */
export function paragraphXml(line: string): string {
  if (line === '') return '<w:p/>';
  // xml:space="preserve" so leading/trailing whitespace in the line survives.
  return `<w:p><w:r><w:t xml:space="preserve">${escapeXmlText(line)}</w:t></w:r></w:p>`;
}

function documentXml(lines: string[]): string {
  const paras = lines.map(paragraphXml).join('');
  // Single trailing <w:sectPr> (A4-size default page) inside the body. Only the
  // `w` namespace is referenced anywhere in the authored markup, including the
  // <w:ins>/<w:del>/<w:moveFrom>/<w:moveTo> the redline author later emits.
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:document xmlns:w="${W_NS}">` +
    `<w:body>${paras}` +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" ' +
    'w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>' +
    '</w:body></w:document>'
  );
}

const CONTENT_TYPES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/word/document.xml" ' +
  'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  '<Override PartName="/word/styles.xml" ' +
  'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
  '<Override PartName="/word/settings.xml" ' +
  'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
  '</Types>';

const PACKAGE_RELS_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" ' +
  'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" ' +
  'Target="word/document.xml"/>' +
  '</Relationships>';

const DOCUMENT_RELS_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" ' +
  'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
  '<Relationship Id="rId2" ' +
  'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>' +
  '</Relationships>';

// Minimal valid styles part: docDefaults + a default Normal paragraph style.
// Tracked-change elements (<w:ins>/<w:del>) reference no named style, but Word
// expects a styles part with a default paragraph style to open cleanly.
const STYLES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  `<w:styles xmlns:w="${W_NS}">` +
  '<w:docDefaults>' +
  '<w:rPrDefault><w:rPr>' +
  '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>' +
  '<w:sz w:val="22"/><w:szCs w:val="22"/>' +
  '</w:rPr></w:rPrDefault>' +
  '<w:pPrDefault/>' +
  '</w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">' +
  '<w:name w:val="Normal"/>' +
  '</w:style>' +
  '</w:styles>';

// Minimal valid settings part. Intentionally does NOT set <w:trackChanges/> — the
// redline carries explicit <w:ins>/<w:del> that render regardless; we do not want
// to silently force the *whole* document into track-changes editing mode.
const SETTINGS_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  `<w:settings xmlns:w="${W_NS}"/>`;

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Synthesise a minimal valid `.docx` package from a plain-text blob, returned as
 * an in-memory `PizZipBundle` ready to feed straight into `extractParts` /
 * `preprocessSide` and (as the Original side) `materialise`.
 *
 * The package contains exactly: `[Content_Types].xml`, `_rels/.rels`,
 * `word/_rels/document.xml.rels`, `word/document.xml`, `word/styles.xml`,
 * `word/settings.xml`.
 */
export function synthesiseDocx(text: string, opts: SynthesiseOptions = {}): PizZipBundle {
  const lines = splitLines(text, opts);
  const bundle = new PizZipBundle();
  bundle.setPart('[Content_Types].xml', CONTENT_TYPES_XML);
  bundle.setPart('_rels/.rels', PACKAGE_RELS_XML);
  bundle.setPart('word/_rels/document.xml.rels', DOCUMENT_RELS_XML);
  bundle.setPart('word/document.xml', documentXml(lines));
  bundle.setPart('word/styles.xml', STYLES_XML);
  bundle.setPart('word/settings.xml', SETTINGS_XML);
  return bundle;
}
