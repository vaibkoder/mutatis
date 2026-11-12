/**
 * host-web — OOXML → DocSnapshot loader.
 *
 * The web loader must produce the SAME model the Office single-pass loader does:
 * full paragraph text (run concatenation, tab/br → space), canonical
 * `body-para-{i}` / `footnote-para-{i}` ids, table-cell flagging, and NO list
 * strings yet (phase 3). These properties keep web findings identical to the
 * add-in's.
 */
import { describe, it, expect } from 'vitest';
import { PizZipBundle } from '@shared/ooxml/index.js';
import { loadDocSnapshotFromOoxml, collectParaNumPrs } from '@host-web/ooxml-loader.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function bundleWith(parts: Record<string, string>): PizZipBundle {
  const b = new PizZipBundle();
  for (const [path, xml] of Object.entries(parts)) b.setPart(path, xml);
  return b;
}

const DOCUMENT_XML =
  `<?xml version="1.0"?><w:document ${W}><w:body>` +
  '<w:p><w:r><w:t>First paragraph.</w:t></w:r></w:p>' +
  '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>A Heading</w:t></w:r></w:p>' +
  '<w:tbl><w:tr>' +
  '<w:tc><w:p><w:r><w:t>Cell one</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>Cell two</w:t></w:r></w:p></w:tc>' +
  '</w:tr></w:tbl>' +
  '<w:p><w:r><w:t>Tabbed</w:t></w:r><w:tab/><w:r><w:t>word</w:t></w:r></w:p>' +
  '</w:body></w:document>';

const FOOTNOTES_XML =
  `<?xml version="1.0"?><w:footnotes ${W}>` +
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:id="1"><w:p><w:r><w:t>A real footnote.</w:t></w:r></w:p></w:footnote>' +
  '</w:footnotes>';

// A document mixing list paragraphs (numId 1, two levels), plain prose, and a
// numId="0" sentinel ("not a list"). Drives both numPr-extraction order and the
// rendering integration case.
const NUMBERED_DOCUMENT_XML =
  `<?xml version="1.0"?><w:document ${W}><w:body>` +
  '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>First item</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t>Intro prose (not a list).</w:t></w:r></w:p>' +
  '<w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Sub item</w:t></w:r></w:p>' +
  '<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>ilvl defaults to 0</w:t></w:r></w:p>' +
  '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr></w:pPr><w:r><w:t>numId 0 = not a list</w:t></w:r></w:p>' +
  '</w:body></w:document>';

// abstractNum 0: level 0 decimal "%1.", level 1 lowerLetter "(%2)"; num 1 → abstract 0.
const NUMBERING_XML =
  `<?xml version="1.0"?><w:numbering ${W}>` +
  '<w:abstractNum w:abstractNumId="0">' +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
  '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%2)"/></w:lvl>' +
  '</w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
  '</w:numbering>';

describe('host-web numPr extraction (collectParaNumPrs)', () => {
  it('emits list paras only, in document order, with default ilvl 0', () => {
    const numPrs = collectParaNumPrs(bundleWith({ 'word/document.xml': NUMBERED_DOCUMENT_XML }));
    expect(numPrs).toEqual([
      { paraId: 'body-para-0', numId: '1', ilvl: 0 },
      { paraId: 'body-para-2', numId: '1', ilvl: 1 }, // prose para-1 skipped
      { paraId: 'body-para-3', numId: '1', ilvl: 0 }, // <w:ilvl> absent → 0
      // para-4 (numId="0") excluded — Word's "not a list" sentinel
    ]);
  });

  it('returns [] for a document with no list paragraphs', () => {
    expect(collectParaNumPrs(bundleWith({ 'word/document.xml': DOCUMENT_XML }))).toEqual([]);
  });

  it('returns [] when document.xml is absent', () => {
    expect(collectParaNumPrs(new PizZipBundle())).toEqual([]);
  });

  // Deferred hook (resolveStyleLinkedNumPr): a paragraph whose numbering is inherited
  // from its <w:pStyle> (→ styles.xml pPr/numPr) is currently treated as non-list.
  // Known gap, not a silent drop — a deferred edge. When the hook lands,
  // this paragraph must emit { paraId: 'body-para-0', numId: '1', ilvl: 0 }.
  it.todo('resolves style-linked numbering via <w:pStyle> → styles.xml pPr/numPr');
});

describe('host-web OOXML loader', () => {
  it('extracts body paragraphs in document order with canonical ids', () => {
    const snap = loadDocSnapshotFromOoxml(bundleWith({ 'word/document.xml': DOCUMENT_XML }));
    expect(snap.paragraphs.map(p => p.text)).toEqual([
      'First paragraph.',
      'A Heading',
      'Cell one',
      'Cell two',
      'Tabbed word', // <w:tab/> collapses to a single space
    ]);
    expect(snap.paragraphs.map(p => p.id)).toEqual([
      'body-para-0', 'body-para-1', 'body-para-2', 'body-para-3', 'body-para-4',
    ]);
    expect(snap.paragraphs.map(p => p.index)).toEqual([0, 1, 2, 3, 4]);
  });

  it('reads the paragraph style from <w:pStyle>, defaulting to Normal', () => {
    const snap = loadDocSnapshotFromOoxml(bundleWith({ 'word/document.xml': DOCUMENT_XML }));
    expect(snap.paragraphs[0]!.styleName).toBe('Normal');
    expect(snap.paragraphs[1]!.styleName).toBe('Heading1');
  });

  it('flags table-cell paragraphs with container + 0-based tableInfo', () => {
    const snap = loadDocSnapshotFromOoxml(bundleWith({ 'word/document.xml': DOCUMENT_XML }));
    const [cellOne, cellTwo] = [snap.paragraphs[2]!, snap.paragraphs[3]!];
    expect(cellOne.container).toBe('tableCell');
    expect(cellOne.tableInfo).toEqual({ tableId: 'table-0', row: 0, col: 0 });
    expect(cellTwo.tableInfo).toEqual({ tableId: 'table-0', row: 0, col: 1 });
    // Body paragraphs are not table cells.
    expect(snap.paragraphs[0]!.container).toBe('body');
    expect(snap.paragraphs[0]!.tableInfo).toBeUndefined();
  });

  it('leaves listString / listLevel undefined for NON-LIST paragraphs', () => {
    // DOCUMENT_XML has no <w:numPr>, so no list paras exist and the (stubbed)
    // renderer is never invoked — these paragraphs stay unnumbered.
    const snap = loadDocSnapshotFromOoxml(bundleWith({ 'word/document.xml': DOCUMENT_XML }));
    for (const p of snap.paragraphs) {
      expect(p.listString).toBeUndefined();
      expect(p.listLevel).toBeUndefined();
    }
  });

  // Renderer landed: numbered paragraphs carry rendered
  // listString/listLevel end-to-end through the loader. Broader edge matrix lives
  // in numbering-render.todo.test.ts.
  it('numbered paragraphs carry rendered listString/listLevel (renderer landed)', () => {
    const snap = loadDocSnapshotFromOoxml(
      bundleWith({
        'word/document.xml': NUMBERED_DOCUMENT_XML,
        'word/numbering.xml': NUMBERING_XML,
      }),
    );
    // body-para-0 = first decimal item; body-para-2 = its lowerLetter sub-item.
    expect(snap.paragraphs[0]!.listString).toBe('1.');
    expect(snap.paragraphs[0]!.listLevel).toBe(0);
    expect(snap.paragraphs[2]!.listString).toBe('(a)');
    expect(snap.paragraphs[2]!.listLevel).toBe(1);
    // body-para-1 is plain prose — never numbered.
    expect(snap.paragraphs[1]!.listString).toBeUndefined();
  });

  it('extracts real footnotes only, skipping separators, with footnote ids', () => {
    const snap = loadDocSnapshotFromOoxml(
      bundleWith({ 'word/document.xml': DOCUMENT_XML, 'word/footnotes.xml': FOOTNOTES_XML }),
    );
    expect(snap.footnotes).toHaveLength(1);
    expect(snap.footnotes[0]!.text).toBe('A real footnote.');
    expect(snap.footnotes[0]!.id).toBe('footnote-para-0');
    expect(snap.footnotes[0]!.container).toBe('footnote');
  });

  it('returns an empty snapshot when document.xml is absent', () => {
    const snap = loadDocSnapshotFromOoxml(new PizZipBundle());
    expect(snap.paragraphs).toEqual([]);
    expect(snap.footnotes).toEqual([]);
  });

  it('getRange throws — navigation on web goes through HostPort.reveal', () => {
    const snap = loadDocSnapshotFromOoxml(bundleWith({ 'word/document.xml': DOCUMENT_XML }));
    expect(() => snap.paragraphs[0]!.getRange({} as never)).toThrow(/reveal/);
  });
});
