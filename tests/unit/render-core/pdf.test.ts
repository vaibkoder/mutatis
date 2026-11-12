import { describe, it, expect, beforeAll } from 'vitest';
import { PDFDocument } from '@pdfme/pdf-lib';
import {
  renderToLayout, renderPdf, createFontResolver, createHarfBuzzShaper,
  type Shaper, type LayoutResult,
} from '@shared/render-core/index.js';
import { testFaces } from './_fixtures.js';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const fonts = createFontResolver(testFaces());
let shaper: Shaper;
beforeAll(async () => { shaper = await createHarfBuzzShaper(); });

/** A two-page-ish body with a footnote ref + an insertion, for the PDF assertions. */
function sampleLayout(): LayoutResult {
  const paras: string[] = [];
  for (let i = 0; i < 60; i++) {
    paras.push(
      `<w:p><w:r><w:t xml:space="preserve">This Agreement clause ${i} is made between the parties and shall be governed by the laws of Singapore. </w:t></w:r>` +
      (i === 0
        ? '<w:r><w:footnoteReference w:id="2"/></w:r><w:ins w:id="9"><w:r><w:t xml:space="preserve">inserted obligation</w:t></w:r></w:ins>'
        : '') +
      '</w:p>',
    );
  }
  const body = `<w:body ${NS}>${paras.join('')}</w:body>`;
  const footnotesXml = `<w:footnotes ${NS}><w:footnote w:id="2"><w:p><w:r><w:t>See the recitals above for defined terms.</w:t></w:r></w:p></w:footnote></w:footnotes>`;
  return renderToLayout({ authoredBodyXml: body, footnotesXml }, fonts, shaper);
}

describe('PDF backend (acceptance criterion 7)', () => {
  it('emits a valid multi-page vector PDF; page count == layout pages (preview↔PDF parity by construction)', async () => {
    const layout = sampleLayout();
    expect(layout.pages.length).toBeGreaterThan(1);

    const bytes = await renderPdf(layout, fonts, { title: 'Test redline' });
    expect(bytes.byteLength).toBeGreaterThan(0);

    const reparsed = await PDFDocument.load(bytes);
    expect(reparsed.getPageCount()).toBe(layout.pages.length);
    expect(reparsed.getTitle()).toBe('Test redline');
  });

  it('subset-embeds the font (FontFile2 ≪ original, Type0/CIDFontType2, ToUnicode)', async () => {
    const layout = sampleLayout();
    // Save uncompressed only to INSPECT the font dicts (engine ships compressed).
    const doc = await PDFDocument.load(await renderPdf(layout, fonts));
    const raw = await doc.save({ useObjectStreams: false });
    const text = Buffer.from(raw).toString('latin1');

    expect(/\/Subtype\s*\/Type0/.test(text)).toBe(true);
    expect(/\/Subtype\s*\/CIDFontType2/.test(text)).toBe(true);
    expect(/\/ToUnicode/.test(text)).toBe(true);
    expect(/\/FontFile2/.test(text)).toBe(true);

    // FontFile2 subset stream must be far smaller than the original font file.
    const origSize = testFaces()[0]!.bytes.byteLength;
    const ffRef = text.match(/\/FontFile2\s+(\d+)\s+0\s+R/);
    expect(ffRef).not.toBeNull();
    const lenRe = new RegExp(`\\b${ffRef![1]}\\s+0\\s+obj[\\s\\S]{0,400}?/Length\\s+(\\d+)`);
    const subsetLen = parseInt(text.match(lenRe)![1]!, 10);
    expect(subsetLen).toBeLessThan(origSize * 0.5);
  });

  it('produces selectable text (pdfjs getTextContent extracts the real words)', async () => {
    const layout = sampleLayout();
    const bytes = await renderPdf(layout, fonts);

    const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
    const task = pdfjs.getDocument({
      data: new Uint8Array(bytes), useWorkerFetch: false, isEvalSupported: false, disableFontFace: true,
      // `isEvalSupported` is a real pdf.js parameter, absent from these typings.
    } as Parameters<typeof pdfjs.getDocument>[0]);
    const pdf = await task.promise;
    const p1 = await pdf.getPage(1);
    const tc = await p1.getTextContent();
    const extracted = tc.items.map((it) => ('str' in it ? it.str : '')).join(' ').replace(/\s+/g, ' ');
    expect(/Agreement/.test(extracted)).toBe(true);
    expect(/Singapore/.test(extracted)).toBe(true);
  });
});
