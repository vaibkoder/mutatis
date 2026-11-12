/**
 * A4 — `loadRenderFaces()` yields the vendored OFL clone set as `FaceInput[]`
 * (real bytes, embedded via `?bytes`), and `createFontResolver` over it maps the
 * Microsoft requested families to their metric clones. `registerPreviewFonts`
 * is a safe no-op where FontFace is unavailable (jsdom).
 */
import { describe, it, expect } from 'vitest';
import { loadRenderFaces, registerPreviewFonts, CLONE_DISPLAY_FAMILIES } from '@shared/render-core/assets/faces.js';
import { createFontResolver } from '@shared/render-core/fonts/registry.js';
import type { EffectiveRPr } from '@shared/render-core/types.js';

const rpr = (fontFamily: string, bold = false, italic = false): EffectiveRPr => ({
  fontFamily, sizePt: 11, bold, italic, underline: false, superscript: false, subscript: false,
});

describe('loadRenderFaces', () => {
  it('returns all 16 clone faces with non-empty sfnt bytes', () => {
    const faces = loadRenderFaces();
    expect(faces).toHaveLength(16);
    for (const f of faces) {
      expect(f.bytes.length).toBeGreaterThan(1000);
      // sfnt magic 0x00010000 (TrueType).
      expect(Array.from(f.bytes.slice(0, 4))).toEqual([0x00, 0x01, 0x00, 0x00]);
    }
    expect(new Set(faces.map((f) => f.family))).toEqual(new Set(CLONE_DISPLAY_FAMILIES));
  });

  it('resolves Microsoft families to their metric clones', () => {
    const r = createFontResolver(loadRenderFaces());
    expect(r.resolve(rpr('Times New Roman')).family).toBe('Liberation Serif');
    expect(r.resolve(rpr('Arial')).family).toBe('Liberation Sans');
    expect(r.resolve(rpr('Calibri')).family).toBe('Carlito');
    expect(r.resolve(rpr('Cambria')).family).toBe('Caladea');
    // metric clone flag + bold/italic style selection.
    const boldTnr = r.resolve(rpr('Times New Roman', true, false));
    expect(boldTnr.isMetricClone).toBe(true);
    expect(boldTnr.bold).toBe(true);
  });

  it('registerPreviewFonts is a safe no-op without FontFace (jsdom)', () => {
    expect(() => registerPreviewFonts()).not.toThrow();
  });
});
