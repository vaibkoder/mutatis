/**
 * A blank paragraph written self-closed — `<w:p/>`, `<w:p w:rsidR="…"/>` — is
 * how Word writes an empty paragraph, and very often the paragraph a table cell
 * must end with (ADR-0046). Every helper that slices a paragraph into open tag,
 * properties, runs and close tag used to read it as an OPEN tag:
 *
 *  - `wrapParaRuns` wrapped the whole paragraph in the revision —
 *    `<w:ins><w:p/></w:ins>`, which `CT_RunTrackChange` does not allow;
 *  - `markParagraphMark` returned it unmarked, or with attributes put
 *    `<w:pPr>` after the closed element;
 *  - `splitParagraph` let a replace's runs land outside any paragraph and lose
 *    a footnote reference on reject (found by the 08h table sweep, seed 439).
 */
import { describe, it, expect } from 'vitest';

import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import { runPipeline } from '@features/compare/pipeline.js';
import { extractParts } from '@features/compare/io/parts.js';
import { splitParagraph } from '@features/compare/output/author.js';
import { flattenTrackedChanges } from '@features/compare/preprocess/flatten.js';
import { markParagraphMark, wrapParaRuns } from '@features/compare/output/revision-markup.js';
import { para, tableBundle } from '../../fixtures/word-table-docx.js';

const CONFIG = getCompareConfig(SG_PROFILE);
const DATE = '2026-10-18T00:00:00Z';

/** A revision element whose content opens a paragraph — never valid. */
const REVISION_AROUND_PARAGRAPH = /<w:(ins|del)\s[^>]*>\s*<w:p[\s/>]/;

describe('a self-closed paragraph is opened before it is revised', () => {
  it('wrapParaRuns keeps the revision inside the paragraph', () => {
    const out = wrapParaRuns('<w:p/>', (runs) => `<w:ins w:id="1">${runs}</w:ins>`);
    expect(out).not.toMatch(REVISION_AROUND_PARAGRAPH);
    expect(out.startsWith('<w:p>')).toBe(true);
    expect(out.endsWith('</w:p>')).toBe(true);
  });

  it('markParagraphMark marks the mark of a self-closed paragraph, attributes kept', () => {
    for (const p of ['<w:p/>', '<w:p w:rsidR="00A1B2C3"/>', '<w:p w:rsidR="00A1B2C3" />']) {
      const out = markParagraphMark(p, 'del', 7, 'Mutatis', DATE);
      expect(out, p).toMatch(/^<w:p(\s[^>]*)?><w:pPr><w:rPr><w:del w:id="7"[^>]*\/><\/w:rPr><\/w:pPr><\/w:p>$/);
      if (p.includes('rsidR')) expect(out).toContain('w:rsidR="00A1B2C3"');
    }
  });

  it('splitParagraph gives an open and a close tag, nothing inside', () => {
    expect(splitParagraph('<w:p w:rsidR="00A1B2C3"/>')).toEqual({
      open: '<w:p w:rsidR="00A1B2C3">', pPr: '', inner: '', close: '</w:p>',
    });
  });

  it('an inserted or deleted blank paragraph in the body is revised by its mark, never wrapped whole', async () => {
    const before = para('The Lender shall notify the Borrower.') + para('The Borrower shall repay the Loan.');
    const after = para('The Lender shall notify the Borrower.') + '<w:p/>' + para('The Borrower shall repay the Loan.');
    for (const [o, r] of [[before, after], [after, before]] as const) {
      const result = await runPipeline({
        originalParts: extractParts(tableBundle(o)),
        revisedParts: extractParts(tableBundle(r)),
        config: CONFIG,
        author: 'Mutatis',
        date: DATE,
      });
      const xml = result.authoredBodyXml;
      expect(xml).not.toMatch(REVISION_AROUND_PARAGRAPH);
      // The blank paragraph is revised by its MARK (ADR-0035), which is how Word
      // records a whole paragraph added or removed — and since the paragraph-mark change
      // the flattener joins on a revised mark as Word does, so a paragraph
      // count through it measures that too.
      const kind = o === before ? 'ins' : 'del';
      expect(xml).toMatch(new RegExp(`<w:p><w:pPr><w:rPr><w:${kind}\\s[^>]*/></w:rPr></w:pPr></w:p>`));
      const count = (x: string): number => (x.match(/<w:p[\s/>]/g) ?? []).length;
      expect(count(flattenTrackedChanges(xml, 'acceptAll'))).toBe(count(r));
      expect(count(flattenTrackedChanges(xml, 'rejectAll'))).toBe(count(o));
    }
  });
});
