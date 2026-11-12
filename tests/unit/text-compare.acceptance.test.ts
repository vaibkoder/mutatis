/**
 * Feature 12 (Compare — additional sources) T5: Text-vs-text acceptance suite.
 * Spec: spec/features/10-compare-additional-sources.md §4.1 · ADR-0018.
 *
 * Runs the FULL text path on synthesised fixtures — synthesiseDocx (both sides)
 * → extractParts → runPipeline (preprocess→structural→word→move→classify→cascade
 * →author) — i.e. the same chain the driver runs, exercised purely. TT5–TT7 use
 * the live driver with a stub HostPort whose writeBack is spied (12f: materialise
 * routes through getActiveHost().writeBack) to prove the click-gating, empty
 * no-op, and oversize cap. The closing guard mirrors feature 08's ADR-0008
 * no-suppression assertion.
 */
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import PizZip from 'pizzip';

import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import type { CompareConfig } from '@shared/config/index.js';
import { setActiveHost, _resetActiveHost } from '@shared/host/index.js';
import type { HostPort } from '@shared/host/index.js';

import { synthesiseDocx } from '@features/compare/io/text/synthesise-docx.js';
import { extractParts } from '@features/compare/io/parts.js';
import { runPipeline } from '@features/compare/pipeline.js';
import { replaceBody } from '@features/compare/output/writeout.js';
import type { DiffOp, ChangeSummary } from '@features/compare/model.js';
import {
  textCompareDriverState as driver,
  MAX_TEXT_PASTE_CHARS,
} from '@features/compare/driver.js';

const CONFIG: CompareConfig = getCompareConfig(SG_PROFILE);
const AUTHOR = 'Mutatis Compare';
const DATE = '2026-06-08T10:00:00Z';

interface TextRun {
  ops: DiffOp[];
  summary: ChangeSummary;
  authoredBodyXml: string;
  /** A valid full document.xml of the redline (authored body in the original root). */
  redlineDocumentXml: string;
}

/** Run the full text path on two plain-text blobs (mirrors the driver, purely). */
async function runText(original: string, revised: string): Promise<TextRun> {
  const originalBundle = synthesiseDocx(original);
  const revisedBundle = synthesiseDocx(revised);
  const result = await runPipeline({
    originalParts: extractParts(originalBundle),
    revisedParts: extractParts(revisedBundle),
    config: CONFIG,
    author: AUTHOR,
    date: DATE,
  });
  const originalDocXml = originalBundle.part('word/document.xml')!;
  return {
    ops: result.ops,
    summary: result.summary,
    authoredBodyXml: result.authoredBodyXml,
    redlineDocumentXml: replaceBody(originalDocXml, result.authoredBodyXml),
  };
}

/**
 * Count PARAGRAPHS carrying an insertion (whole-paragraph insertions).
 *
 * Counts paragraphs, not `<w:ins>` elements: since ADR-0035 a whole-paragraph
 * insertion emits two — one round the runs, one on the paragraph MARK so that
 * rejecting the change removes the paragraph instead of blanking it.
 */
function countInsParagraphs(bodyXml: string): number {
  return bodyXml
    .split('</w:p>')
    .filter((chunk) => chunk.includes('<w:p') && /<w:ins\b/.test(chunk)).length;
}

// ---------------------------------------------------------------------------
// TT1 — inserted paragraph mid-text
// ---------------------------------------------------------------------------

describe('TT1 — inserted paragraph mid-text', () => {
  it('one inserted paragraph → exactly one insert op, others equal', async () => {
    const { ops, authoredBodyXml } = await runText(
      'Clause 1.\nClause 2.\nClause 3.',
      'Clause 1.\nNEW CLAUSE.\nClause 2.\nClause 3.',
    );
    const nonEqual = ops.filter(o => o.type !== 'equal');
    expect(nonEqual).toHaveLength(1);
    expect(nonEqual[0]!.type).toBe('insert');
    // Exactly one <w:ins> wrapper in the authored body (the inserted paragraph).
    expect(countInsParagraphs(authoredBodyXml)).toBe(1);
    expect(authoredBodyXml).not.toContain('<w:del');
  });
});

// ---------------------------------------------------------------------------
// TT2 — word-level shall→may
// ---------------------------------------------------------------------------

describe('TT2 — word-level obligation change', () => {
  it('shall→may → in-paragraph ins/del, classified obligationLanguage, not filtered', async () => {
    const { ops, authoredBodyXml } = await runText(
      'The Seller shall deliver the Goods.',
      'The Seller may deliver the Goods.',
    );
    const nonEqual = ops.filter(o => o.type !== 'equal');
    expect(nonEqual).toHaveLength(1);
    expect(nonEqual[0]!.classification).toBe('obligationLanguage');
    // Run-level revisions inside the paragraph.
    expect(authoredBodyXml).toContain('<w:ins');
    expect(authoredBodyXml).toContain('<w:del');
    expect(authoredBodyXml).toContain('may');
    expect(authoredBodyXml).toContain('shall');
    // obligationLanguage is never a filter-eligible (formatting/whitespace) class.
    expect(nonEqual[0]!.classification).not.toBe('formatting');
    expect(nonEqual[0]!.classification).not.toBe('whitespacePunct');
  });
});

// ---------------------------------------------------------------------------
// TT3 — verbatim move
// ---------------------------------------------------------------------------

describe('TT3 — verbatim moved paragraph', () => {
  it('a verbatim move → moveFrom/moveTo, not a delete+insert pair', async () => {
    const moved = 'This indemnity clause is moved verbatim and unchanged.';
    const { ops, authoredBodyXml } = await runText(
      `Clause 1.\nClause 2.\n${moved}\nClause 3.`,
      `${moved}\nClause 1.\nClause 2.\nClause 3.`,
    );
    const moves = ops.filter(o => o.type === 'move');
    expect(moves.length).toBe(2); // from + to pair
    expect(authoredBodyXml).toContain('w:moveFromRangeStart');
    expect(authoredBodyXml).toContain('w:moveToRangeStart');
    // Shared move name across both endpoints.
    const names = [...new Set((authoredBodyXml.match(/w:name="(move-[^"]+)"/g) ?? []))];
    expect(names.length).toBe(1);
  });
});

// ---------------------------------------------------------------------------
// TT4 — package validity + special-char escaping
// ---------------------------------------------------------------------------

describe('TT4 — valid OOXML package + escaping', () => {
  it('the authored redline assembles into a valid, re-openable package', async () => {
    const { redlineDocumentXml } = await runText('Clause 1.', 'Clause 1.\nClause 2 added.');
    // Re-pack the redline document into a fresh bundle and re-open it.
    const bundle = synthesiseDocx(''); // reuse the controlled package skeleton
    bundle.setPart('word/document.xml', redlineDocumentXml);
    const reopened = new PizZip(bundle.toUint8Array());
    expect(reopened.file('word/document.xml')).not.toBeNull();
    expect(reopened.file('word/document.xml')!.asText()).toContain('<w:ins');
  });

  it('XML-special characters in the paste are escaped and do not corrupt the package', async () => {
    const { authoredBodyXml, redlineDocumentXml } = await runText(
      'Plain line.',
      'A & B < C > D inserted.',
    );
    expect(authoredBodyXml).toContain('&amp;');
    expect(authoredBodyXml).toContain('&lt;');
    expect(authoredBodyXml).toContain('&gt;');
    // No raw, unescaped ampersand survived into the markup.
    expect(authoredBodyXml).not.toMatch(/&(?!amp;|lt;|gt;|quot;|apos;|#)/);
    // The reassembled document is still valid bytes.
    const bundle = synthesiseDocx('');
    bundle.setPart('word/document.xml', redlineDocumentXml);
    expect(new PizZip(bundle.toUint8Array()).file('word/document.xml')).not.toBeNull();
  });
});

// ---------------------------------------------------------------------------
// TT5–TT7 — driver behaviour (click-gating, empty no-op, oversize)
// ---------------------------------------------------------------------------

describe('TT5–TT7 — driver gating', () => {
  let writeBack: ReturnType<typeof vi.fn>;
  beforeEach(() => {
    driver.status.value = 'idle';
    driver.errorMessage.value = null;
    driver.ops.value = null;
    driver.summary.value = null;
    driver.originalText.value = '';
    driver.revisedText.value = '';
    driver.redlineReady.value = false;
    writeBack = vi.fn().mockResolvedValue(undefined);
    setActiveHost({ writeBack } as unknown as HostPort);
  });
  afterEach(() => {
    _resetActiveHost();
    vi.restoreAllMocks();
  });

  it('TT5 — redline generated but not opened until openRedline()', async () => {
    const spy = writeBack;
    driver.originalText.value = 'Clause 1.';
    driver.revisedText.value = 'Clause 1 amended.';
    await driver.runCompare();
    expect(driver.redlineReady.value).toBe(true);
    expect(spy).not.toHaveBeenCalled();
    await driver.openRedline();
    expect(spy).toHaveBeenCalledTimes(1);
  });

  it('TT6 — empty/whitespace both sides → message, no crash, no redline', async () => {
    const spy = writeBack;
    driver.originalText.value = '   ';
    driver.revisedText.value = '\n\n';
    await driver.runCompare();
    expect(driver.status.value).toBe('error');
    expect(driver.errorMessage.value).toBeTruthy();
    expect(driver.redlineReady.value).toBe(false);
    expect(spy).not.toHaveBeenCalled();
  });

  it('TT7 — oversize paste → friendly limit message, nothing generated', async () => {
    driver.originalText.value = 'x'.repeat(MAX_TEXT_PASTE_CHARS + 1);
    driver.revisedText.value = 'y';
    await driver.runCompare();
    expect(driver.status.value).toBe('error');
    expect(driver.errorMessage.value).toMatch(/too large/i);
    expect(driver.redlineReady.value).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// TT8 — ADR-0008 no-suppression
// ---------------------------------------------------------------------------

describe('TT8 — ADR-0008 no-suppression', () => {
  it('every change between the two blobs is present in the redline', async () => {
    const { ops, summary, authoredBodyXml } = await runText(
      'The Seller shall deliver the Goods.\nClause 2 stays.\nDelete this line.',
      'The Seller may deliver the Products.\nClause 2 stays.\nA brand new line.',
    );
    const nonEqual = ops.filter(o => o.type !== 'equal');
    // Multiple changes detected; the summary total reflects all of them.
    expect(nonEqual.length).toBeGreaterThanOrEqual(2);
    expect(summary.total).toBe(nonEqual.length);
    // The redline carries insertions and deletions — nothing is silently dropped.
    expect(authoredBodyXml).toContain('<w:ins');
    expect(authoredBodyXml).toContain('<w:del');
    // The changed words all appear in the authored markup.
    expect(authoredBodyXml).toContain('may');
    expect(authoredBodyXml).toContain('Products');
    // The third line's replacement is present (tokenised across runs by the word diff).
    expect(authoredBodyXml).toContain('brand new');
    expect(authoredBodyXml).toContain('Delete this'); // deleted text still in the redline
  });

  it('classification only categorises — it never reduces the op count', async () => {
    const { ops } = await runText(
      'Alpha.\nBeta.\nGamma.',
      'Alpha changed.\nBeta.\nGamma also changed.',
    );
    // Two replaced paragraphs + one equal = three ops; none dropped.
    expect(ops.length).toBe(3);
    expect(ops.filter(o => o.type === 'equal').length).toBe(1);
    expect(ops.filter(o => o.type === 'replace').length).toBe(2);
  });
});
