import { describe, it, expect } from 'vitest';
import { readLineMetrics, singleLineHeightPt } from '@shared/render-core/fonts/sfnt-metrics.js';
import { createFontResolver, CLONE_MAP } from '@shared/render-core/fonts/registry.js';
import { deobfuscateOdttf, parseFontKey } from '@shared/render-core/fonts/odttf.js';
import { createHarfBuzzShaper } from '@shared/render-core/layout/shape.js';
import type { EffectiveRPr } from '@shared/render-core/types.js';
import { libSansRegular, testFaces } from './_fixtures.js';

const rpr = (over: Partial<EffectiveRPr>): EffectiveRPr => ({
  fontFamily: 'Times New Roman', sizePt: 11, bold: false, italic: false,
  underline: false, superscript: false, subscript: false, ...over,
});

describe('sfnt-metrics', () => {
  it('reads Liberation Sans OS/2 win metrics + usWin line height (Spike A figures)', () => {
    const m = readLineMetrics(libSansRegular());
    expect(m.unitsPerEm).toBe(2048);
    expect(m.useTypoMetrics).toBe(false);
    expect(m.winAscent).toBe(1854);
    expect(m.winDescent).toBe(434);
    expect(singleLineHeightPt(m, 11)).toBeCloseTo(12.289, 2);
  });
});

describe('font registry (clone mapping)', () => {
  const resolver = createFontResolver(testFaces());

  it('maps Arial → Liberation Sans (metric clone)', () => {
    const f = resolver.resolve(rpr({ fontFamily: 'Arial' }));
    expect(f.family).toBe('Liberation Sans');
    expect(f.isMetricClone).toBe(true);
  });
  it('maps Times New Roman → Liberation Serif (metric clone)', () => {
    const f = resolver.resolve(rpr({ fontFamily: 'Times New Roman' }));
    expect(f.family).toBe('Liberation Serif');
    expect(f.isMetricClone).toBe(true);
  });
  it('flags a no-clone family (Garamond) as non-metric', () => {
    const f = resolver.resolve(rpr({ fontFamily: 'Garamond' }));
    expect(f.isMetricClone).toBe(false);
  });
  it('selects the bold face when bold is set', () => {
    const f = resolver.resolve(rpr({ fontFamily: 'Arial', bold: true }));
    expect(f.bold).toBe(true);
  });
  it('CLONE_MAP covers the four core MS families', () => {
    expect(CLONE_MAP.get('times new roman')).toBe('Liberation Serif');
    expect(CLONE_MAP.get('arial')).toBe('Liberation Sans');
    expect(CLONE_MAP.get('calibri')).toBe('Carlito');
    expect(CLONE_MAP.get('cambria')).toBe('Caladea');
  });
});

describe('odttf de-obfuscation', () => {
  it('parses a GUID into 16 bytes', () => {
    const k = parseFontKey('{01234567-89AB-CDEF-0123-456789ABCDEF}');
    expect(k.length).toBe(16);
    expect(k[0]).toBe(0x01);
    expect(k[15]).toBe(0xef);
  });
  it('is an involution on the first 32 bytes; leaves the rest untouched', () => {
    const key = '{01234567-89AB-CDEF-0123-456789ABCDEF}';
    const original = new Uint8Array(48);
    for (let i = 0; i < 48; i++) original[i] = (i * 7 + 3) & 0xff;
    const once = deobfuscateOdttf(original, key);
    const twice = deobfuscateOdttf(once, key);
    expect([...twice]).toEqual([...original]); // XOR twice = identity
    // bytes 32..47 unchanged after a single pass
    for (let i = 32; i < 48; i++) expect(once[i]).toBe(original[i]);
    // at least one of the first 32 changed
    expect([...once.slice(0, 32)]).not.toEqual([...original.slice(0, 32)]);
  });
});

describe('HarfBuzz shaper (kern/liga off, deterministic)', () => {
  it('matches Spike A width for "Office" @12pt and is deterministic', async () => {
    const shaper = await createHarfBuzzShaper();
    const f = createFontResolver(testFaces()).resolve(rpr({ fontFamily: 'Arial', sizePt: 12 }));
    const w1 = shaper.advanceWidthPt(f, 'Office', 12);
    const w2 = shaper.advanceWidthPt(f, 'Office', 12);
    expect(w1).toBeCloseTo(31.34, 1);
    expect(w1).toBe(w2);
  });
  it('has NO kerning: width("AV") == width("A") + width("V")', async () => {
    const shaper = await createHarfBuzzShaper();
    const f = createFontResolver(testFaces()).resolve(rpr({ fontFamily: 'Arial', sizePt: 12 }));
    const av = shaper.advanceWidthPt(f, 'AV', 12);
    const a = shaper.advanceWidthPt(f, 'A', 12);
    const v = shaper.advanceWidthPt(f, 'V', 12);
    expect(av).toBeCloseTo(a + v, 4);
  });
});
