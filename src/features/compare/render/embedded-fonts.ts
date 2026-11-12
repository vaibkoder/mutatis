/**
 * Embedded-font extraction (feature 16 / ADR-0024, Phase 2 B2).
 *
 * When a `.docx` embeds its own fonts (the documented `w:embedRegular` … bits),
 * those bytes are the *exact* faces the author saw — so RenderCore prefers them
 * over the OFL metric clones. The bytes live obfuscated in `word/fonts/*.odttf`
 * (XOR'd with a per-font GUID); `word/fontTable.xml` carries the family name +
 * `w:fontKey` per style and a relationship id; `word/_rels/fontTable.xml.rels`
 * maps the id → the `.odttf` part path.
 *
 * Host-free-by-discipline: imports only `@shared/*`, never `Office/Word`. Every
 * step degrades gracefully — a missing fontTable, missing rels, missing bytes, or
 * a single malformed key skips just that face and never throws. The clones cover
 * everything an extraction misses, so a degraded extraction is a quality, not a
 * correctness, gap.
 */
import { parsePreserveOrder, tagOf, childrenOf, attrOf } from '@shared/ooxml/po-xml.js';
import type { DocxBundle } from '@shared/ooxml/types.js';
import { deobfuscateOdttf, type FaceInput } from '@shared/render-core/index.js';

/** w:embed* tag → the (bold, italic) style it carries. */
const EMBED_STYLE: Record<string, { bold: boolean; italic: boolean }> = {
  'w:embedRegular': { bold: false, italic: false },
  'w:embedBold': { bold: true, italic: false },
  'w:embedItalic': { bold: false, italic: true },
  'w:embedBoldItalic': { bold: true, italic: true },
};

/** Parse `word/_rels/fontTable.xml.rels` → relationship id → target part path (under `word/`). */
function relTargets(relsXml: string): Map<string, string> {
  const map = new Map<string, string>();
  const roots = parsePreserveOrder(relsXml);
  const rels = roots.find((n) => tagOf(n) === 'Relationships');
  if (!rels) return map;
  for (const rel of childrenOf(rels)) {
    if (tagOf(rel) !== 'Relationship') continue;
    const id = attrOf(rel, '@_Id');
    const target = attrOf(rel, '@_Target');
    if (id === undefined || target === undefined) continue;
    // Targets are relative to the part's base (`word/`). Normalise a leading slash.
    const path = target.replace(/^\/+/, '');
    map.set(id, path.startsWith('word/') ? path : `word/${path}`);
  }
  return map;
}

/**
 * Extract embedded faces (de-obfuscated sfnt bytes) from a `.docx` bundle.
 * Returns `[]` when the document embeds no fonts or binary access is unavailable.
 * Each face is flagged `embedded: true` so the resolver prefers it over a clone
 * of the same family.
 */
export function extractEmbeddedFaces(bundle: DocxBundle): FaceInput[] {
  if (typeof bundle.partBytes !== 'function') return []; // text-only bundle (no binary parts)
  const fontTableXml = bundle.part('word/fontTable.xml');
  const relsXml = bundle.part('word/_rels/fontTable.xml.rels');
  if (fontTableXml === undefined || relsXml === undefined) return [];

  const targets = relTargets(relsXml);
  const out: FaceInput[] = [];

  let fonts;
  try {
    fonts = parsePreserveOrder(fontTableXml).find((n) => tagOf(n) === 'w:fonts');
  } catch {
    return [];
  }
  if (!fonts) return [];

  for (const fontNode of childrenOf(fonts)) {
    if (tagOf(fontNode) !== 'w:font') continue;
    const family = attrOf(fontNode, '@_w:name');
    if (family === undefined || family.trim() === '') continue;

    for (const embed of childrenOf(fontNode)) {
      const style = EMBED_STYLE[tagOf(embed) ?? ''];
      if (!style) continue;
      const rId = attrOf(embed, '@_r:id');
      const fontKey = attrOf(embed, '@_w:fontKey');
      if (rId === undefined || fontKey === undefined) continue;
      const partPath = targets.get(rId);
      if (partPath === undefined) continue;

      try {
        const odttf = bundle.partBytes(partPath);
        if (!odttf || odttf.length === 0) continue;
        const bytes = deobfuscateOdttf(odttf, fontKey);
        out.push({ family, bold: style.bold, italic: style.italic, bytes, embedded: true });
      } catch {
        // Malformed key / unreadable part → skip this one face; clones cover it.
      }
    }
  }
  return out;
}
