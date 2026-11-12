import { describe, expect, it } from 'vitest';
import PizZip from 'pizzip';
import { markdownToBlocks, looksLikeMarkdown } from '../src/renderer/src/paste/markdown-to-blocks.js';
import { plainToBlocks } from '../src/renderer/src/paste/plain-to-blocks.js';
import { blocksToDocx } from '../src/renderer/src/paste/blocks-to-docx.js';
import { buildPastedPair, emptyPasteSide, type PasteSide } from '../src/renderer/src/paste/pasted.js';
import { blocksText, hasFormatting, type Para } from '../src/renderer/src/paste/model.js';
import { runComparison } from '../src/renderer/src/compare/run-comparison.js';

const md = (s: string) => markdownToBlocks(s);
const side = (text: string, format: PasteSide['format'] = 'plain'): PasteSide => ({
  ...emptyPasteSide,
  text,
  format,
});

describe('plain text', () => {
  it('is one paragraph per line and keeps blank lines', () => {
    const blocks = plainToBlocks('one\n\ntwo\n');
    expect(blocks.map((b) => (b.kind === 'p' ? b.runs.map((r) => r.text).join('') : '?'))).toEqual(['one', '', 'two']);
  });
});

describe('markdown', () => {
  it('carries bold, italic, strike and headings', () => {
    const [h, p] = md('## Term\n\nThe **Seller** shall *not* ~~sell~~.') as Para[];
    expect(h?.heading).toBe(2);
    expect(p?.runs.find((r) => r.text === 'Seller')?.bold).toBe(true);
    expect(p?.runs.find((r) => r.text === 'not')?.italic).toBe(true);
    expect(p?.runs.find((r) => r.text === 'sell')?.strike).toBe(true);
  });

  it('writes list markers into the text, ordered and bulleted, nested', () => {
    const text = blocksText(md('- a\n  - b\n- c\n\n3. x\n4. y'));
    expect(text).toBe('• a\n• b\n• c\n3. x\n4. y');
  });

  it('reads a GFM table with a bold header row', () => {
    const [t] = md('| A | B |\n|---|---|\n| 1 | 2 |');
    expect(t?.kind).toBe('table');
    if (t?.kind !== 'table') return;
    expect(t.rows).toHaveLength(2);
    expect(t.rows[0]?.[0]?.[0]?.runs[0]?.bold).toBe(true);
    expect(t.rows[1]?.[1]?.[0]?.runs[0]?.text).toBe('2');
  });

  it('keeps raw HTML as inert text rather than interpreting it', () => {
    expect(blocksText(md('hello <script>alert(1)</script>'))).toContain('<script>');
  });

  it('only calls it markdown on a clear signal', () => {
    expect(looksLikeMarkdown('# Heading\n\ntext')).toBe(true);
    expect(looksLikeMarkdown('a **bold** word')).toBe(true);
    expect(looksLikeMarkdown('- one\n- two')).toBe(true);
    expect(looksLikeMarkdown('1. Definitions\n2. Interpretation')).toBe(false);
    expect(looksLikeMarkdown('The price is 5 * 3 = 15')).toBe(false);
    expect(hasFormatting(md('plain'))).toBe(false);
  });
});

describe('docx output', () => {
  const parts = (bytes: Uint8Array) => {
    const zip = new PizZip(bytes);
    return {
      document: zip.file('word/document.xml')!.asText(),
      names: Object.keys(zip.files).sort(),
    };
  };

  it('is a complete package with schema-ordered run properties', () => {
    const { document, names } = parts(blocksToDocx(md('# T\n\n***both*** & <b>')));
    expect(names).toEqual(
      ['[Content_Types].xml', '_rels/.rels', 'word/_rels/document.xml.rels', 'word/document.xml', 'word/settings.xml', 'word/styles.xml'].sort(),
    );
    expect(document).toContain('<w:pStyle w:val="Heading1"/>');
    expect(document).toContain('<w:rPr><w:b/><w:bCs/><w:i/><w:iCs/></w:rPr>');
    expect(document).toContain('&amp;');
    expect(document).toContain('&lt;b&gt;');
  });

  it('strips control characters that would corrupt the XML', () => {
    expect(parts(blocksToDocx(plainToBlocks('a\u0001b'))).document).toContain('>ab<');
  });
});

describe('through the real compare engine', () => {
  const compare = async (a: PasteSide, b: PasteSide) => {
    const pair = buildPastedPair(a, b)!;
    expect(pair).not.toBeNull();
    return runComparison({
      originalBytes: await pair.original.file.arrayBuffer(),
      revisedBytes: await pair.modified.file.arrayBuffer(),
      jurisdiction: 'UK',
    });
  };

  it('finds the edit between two pasted plain texts', async () => {
    const out = await compare(
      side('The Seller shall pay S$5,000 on Completion.'),
      side('The Seller shall pay S$6,000 on Completion.'),
    );
    expect(out.summary.total).toBeGreaterThan(0);
  });

  it('hands pasted formatting to the engine, which classes a bold-only edit as formatting', async () => {
    const out = await compare(
      side('The Seller shall pay on Completion.', 'markdown'),
      side('The **Seller** shall pay on Completion.', 'markdown'),
    );
    // The engine's default profile hides formatting-only changes from the total
    // (summary.filtersActive.hideFormattingOnly) but still reports the op.
    expect(out.ops.some((op) => op.classes.includes('formatting'))).toBe(true);
  });

  it('compares headings, lists and tables end to end', async () => {
    const original = '# Fees\n\n- Base: 5\n- Extra: 2\n\n| Item | Cost |\n|---|---|\n| A | 1 |';
    const modified = '# Fees\n\n- Base: 6\n- Extra: 2\n\n| Item | Cost |\n|---|---|\n| A | 2 |';
    const out = await compare(side(original, 'markdown'), side(modified, 'markdown'));
    expect(out.summary.total).toBeGreaterThan(0);
  });

  it('reports no changes for identical pastes', async () => {
    const out = await compare(side('Same text.\n\nSecond paragraph.'), side('Same text.\n\nSecond paragraph.'));
    expect(out.summary.total).toBe(0);
  });

  it('refuses to build a pair when a side is empty', () => {
    expect(buildPastedPair(side('text'), side('   \n '))).toBeNull();
  });
});

describe('uploaded .txt and .md files', () => {
  it('are converted to docx that the engine can compare', async () => {
    const { convertUploads } = await import('../src/renderer/src/paste/convert-uploads.js');
    const mk = (name: string, body: string, id: string) => ({
      id, name, byteLength: body.length, kind: 'text' as const, overCap: false,
      file: new File([body], name, { type: 'text/plain' }),
    });
    const { docs: [a, b], failed } = await convertUploads([
      mk('old.md', '# Fees\n\nThe **Buyer** pays 5 days.', 'd1'),
      mk('new.txt', '# Fees\n\nThe **Buyer** pays 7 days.', 'd2'),
    ]);
    expect(failed).toEqual([]);
    expect(a?.kind).toBe('docx');
    expect(a?.name).toBe('old.md');
    expect(a?.id).toBe('d1');
    const out = await runComparison({
      originalBytes: await a!.file.arrayBuffer(),
      revisedBytes: await b!.file.arrayBuffer(),
    });
    expect(out.summary.total).toBeGreaterThan(0);
    // .txt is plain: the "#" and "**" survive as characters.
    expect(JSON.stringify(out.model)).toContain('**');
  });
});
