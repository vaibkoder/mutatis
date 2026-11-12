// Shared HarfBuzz shaping helpers for the RenderCore spikes (THROWAWAY).
// Word Latin defaults: kerning OFF, ligatures OFF, hyphenation OFF.
import { readFileSync } from 'node:fs';
import * as hb from 'harfbuzzjs';

// Features that switch off everything Word leaves off by default.
const OFF = ['-kern', '-liga', '-clig', '-rlig', '-calt'].map((s) => hb.Feature.fromString(s));

export function loadFont(path) {
  const bytes = readFileSync(path);
  const blob = new hb.Blob(bytes);
  const face = new hb.Face(blob, 0);
  const font = new hb.Font(face);
  return { bytes, blob, face, font, upem: face.upem };
}

/** Shape one run of text; return {glyphs, widthFu} in font units (kern/liga off). */
export function shapeRun(font, text) {
  const buf = new hb.Buffer();
  buf.addText(text);
  buf.guessSegmentProperties();
  hb.shape(font, buf, OFF);
  const glyphs = buf.getGlyphInfosAndPositions();
  const widthFu = glyphs.reduce((a, g) => a + (g.xAdvance ?? 0), 0);
  return { glyphs, widthFu };
}

/** Advance width of `text` in points at `sizePt`. */
export function widthPt(font, upem, text, sizePt) {
  return (shapeRun(font, text).widthFu / upem) * sizePt;
}

export { hb };
