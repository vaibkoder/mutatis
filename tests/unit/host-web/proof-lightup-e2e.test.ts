/**
 * host-web — Phase 3b feature light-up, end-to-end.
 *
 * Proves the three Section-Index-dependent checkers run identically off-host once
 * phase 3a renders list numbers. A single web-loaded `.docx` exercises:
 *   - Cross-reference: a reference to "Clause 1" RESOLVES (→ static badge under the
 *     ctx=null degraded path) — and it can only resolve because body-para-0's
 *     rendered `listString` "1." (computed by the renderer from numbering.xml) created
 *     the clause:1 index key. A reference to "Clause 9" is not-found. This is the
 *     load-bearing 3a→3b bridge.
 *   - Static-numbering ("numbering detect"): a hand-typed "4.2 …" clause number is
 *     flagged.
 *   - Defined terms: a quoted term defined in a TWO-COLUMN TABLE is harvested — which
 *     exercises the web loader's `tableInfo` grouping (`defined-terms/tables.ts` is the
 *     first ported tableInfo consumer; this is the Session-35 re-validation).
 *
 * The pipeline mirrors WebApp.runAllChecks exactly: load → analyse with ctx=null.
 */
import { describe, it, expect } from 'vitest';
import { PizZipBundle } from '@shared/ooxml/index.js';
import { UK_PROFILE } from '@shared/config/index.js';
import { loadDocSnapshotFromOoxml } from '@host-web/ooxml-loader.js';
import { analyseCrossReference } from '@features/cross-reference/index.js';
import { analyseDefinedTerms } from '@features/defined-terms/index.js';
import { analysePartyRole } from '@features/party-role/index.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function bundleWith(parts: Record<string, string>): PizZipBundle {
  const b = new PizZipBundle();
  for (const [path, xml] of Object.entries(parts)) b.setPart(path, xml);
  return b;
}

// body-para-0 is an auto-numbered clause (numId 1 → "1."); the rest is prose + a
// two-column definition table. No <w:numPr> anywhere except para-0.
const DOCUMENT_XML =
  `<?xml version="1.0"?><w:document ${W}><w:body>` +
  // para-0: Clause 1 (auto-numbered). Its rendered listString "1." → index key clause:1.
  '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Definitions and Interpretation</w:t></w:r></w:p>' +
  // para-1: resolvable cross-reference to Clause 1.
  '<w:p><w:r><w:t>The provisions of Clause 1 apply throughout this agreement.</w:t></w:r></w:p>' +
  // para-2: broken cross-reference to a clause that does not exist.
  '<w:p><w:r><w:t>Subject to Clause 9, the parties shall perform their obligations.</w:t></w:r></w:p>' +
  // para-3: manually-typed clause number (no numPr) → static-numbering flag.
  '<w:p><w:r><w:t>4.2 The parties shall act in good faith at all times.</w:t></w:r></w:p>' +
  // table: two rows, two columns → quoted-term definitions (tableInfo grouping).
  '<w:tbl>' +
  '<w:tr>' +
  '<w:tc><w:p><w:r><w:t>"Widget"</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>means a small mechanical device supplied under this agreement.</w:t></w:r></w:p></w:tc>' +
  '</w:tr>' +
  '<w:tr>' +
  '<w:tc><w:p><w:r><w:t>"Gadget"</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>means any other appliance listed in the schedule.</w:t></w:r></w:p></w:tc>' +
  '</w:tr>' +
  '</w:tbl>' +
  '</w:body></w:document>';

// abstractNum 0: level 0 decimal "%1."; num 1 → abstract 0.
const NUMBERING_XML =
  `<?xml version="1.0"?><w:numbering ${W}>` +
  '<w:abstractNum w:abstractNumId="0">' +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
  '</w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
  '</w:numbering>';

function loadParas() {
  const snap = loadDocSnapshotFromOoxml(
    bundleWith({ 'word/document.xml': DOCUMENT_XML, 'word/numbering.xml': NUMBERING_XML }),
  );
  return { snap, paras: [...snap.paragraphs, ...snap.footnotes] };
}

describe('host-web Phase 3b — Proof checkers light up on web', () => {
  it('renders the clause number so the Section Index can resolve a reference to it', async () => {
    const { snap, paras } = loadParas();

    // 3a bridge: body-para-0 carries the rendered list string.
    expect(snap.paragraphs[0]!.listString).toBe('1.');

    const result = await analyseCrossReference(paras, UK_PROFILE, null);
    const cats = result.xrefFindings.map(f => f.category);

    // Clause 9 does not exist → not-found.
    expect(cats).toContain('cross-reference/not-found');
    const notFound = result.xrefFindings.find(f => f.category === 'cross-reference/not-found');
    expect(notFound?.title).toContain('9');

    // Clause 1 RESOLVED (it is in the index, built from the rendered "1.") → it must
    // NOT appear as not-found, and surfaces as a manually-typed reference badge
    // (ctx=null degraded path). If the renderer had not run, clause:1 would be absent
    // and "Clause 1" would itself be not-found — so this is the 3a→3b proof.
    expect(cats).toContain('cross-reference/static');
    const notFoundTitles = result.xrefFindings
      .filter(f => f.category === 'cross-reference/not-found')
      .map(f => f.title)
      .join(' ');
    expect(notFoundTitles).not.toContain('Clause 1');
  });

  it('flags a manually-typed clause number (static-numbering detect)', async () => {
    const { paras } = loadParas();
    const result = await analyseCrossReference(paras, UK_PROFILE, null);
    const cats = result.staticNumFindings.map(f => f.category);
    expect(cats).toContain('static-numbering/manually-typed');
    // The flag points at the "4.2" paragraph, not the auto-numbered Clause 1.
    const flagged = result.staticNumFindings.find(f => f.category === 'static-numbering/manually-typed');
    expect(flagged?.title).toContain('4.2');
  });

  it('harvests a table-cell definition via the web loader\'s tableInfo grouping', () => {
    const { paras } = loadParas();
    const analysis = analyseDefinedTerms(paras, UK_PROFILE);

    // The quoted terms in the two-column table were reconstructed from tableInfo
    // (the map keys terms normalised to lower case).
    expect(analysis.definitionsMap.has('widget')).toBe(true);
    expect(analysis.definitionsMap.has('gadget')).toBe(true);
    const widget = analysis.definitionsMap.get('widget');
    expect(widget?.definitions.some(d => d.source === 'table')).toBe(true);

    // Defined-but-never-used → a defined-terms finding is emitted.
    expect(analysis.findings.map(f => f.category)).toContain('defined-terms/defined-not-used');
  });

  it('flags a party-role inconsistency over the web loader (feature 09)', () => {
    // Minimal own fixture: one entity introduced under two different roles. The
    // party-role checker is pure analysis (no host), so it runs unchanged on the
    // web-loaded snapshot exactly as in the add-in.
    const PARTY_XML =
      `<?xml version="1.0"?><w:document ${W}><w:body>` +
      '<w:p><w:r><w:t>ACME Pte Ltd (the "Buyer") is a party to this agreement.</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>ACME Pte Ltd (the "Purchaser") shall complete the transaction.</w:t></w:r></w:p>' +
      '</w:body></w:document>';
    const snap = loadDocSnapshotFromOoxml(bundleWith({ 'word/document.xml': PARTY_XML }));
    const paras = [...snap.paragraphs, ...snap.footnotes];

    const analysis = analysePartyRole(paras, UK_PROFILE);
    const incon = analysis.findings.filter(f => f.category === 'party-role/inconsistent-label');
    expect(incon.length).toBeGreaterThan(0);
    expect(incon[0]?.title).toMatch(/Buyer/);
    expect(incon[0]?.title).toMatch(/Purchaser/);
  });
});
