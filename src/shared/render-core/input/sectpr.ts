/**
 * Page geometry from `<w:sectPr>` — host-free, pure.
 * Jurisdiction invariant (ADR-0002): the fallback when sectPr is absent is **A4**
 * (SG/UK), never US Letter. sectPr values are twips (20 twips = 1pt).
 */
import { type PoNode, tagOf, attrOf } from '@shared/ooxml/po-xml.js';
import type { PageGeometry } from '../types.js';

const TWIP = 20;

/** A4 portrait, Word default 1-inch margins (the SG/UK default). */
export const A4_DEFAULT: PageGeometry = {
  widthPt: 595.28,
  heightPt: 841.89,
  marginTopPt: 72,
  marginRightPt: 72,
  marginBottomPt: 72,
  marginLeftPt: 72,
};

function num(node: PoNode, attr: string): number | undefined {
  const v = attrOf(node, attr);
  if (v === undefined) return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

/** Read geometry from a sectPr child list; missing fields fall back to A4 defaults. */
export function geometryFromSectPr(sectPr: PoNode[] | undefined): PageGeometry {
  const g: PageGeometry = { ...A4_DEFAULT };
  if (!sectPr) return g;
  for (const node of sectPr) {
    const t = tagOf(node);
    if (t === 'w:pgSz') {
      const w = num(node, '@_w:w');
      const h = num(node, '@_w:h');
      if (w !== undefined) g.widthPt = w / TWIP;
      if (h !== undefined) g.heightPt = h / TWIP;
    } else if (t === 'w:pgMar') {
      const top = num(node, '@_w:top');
      const right = num(node, '@_w:right');
      const bottom = num(node, '@_w:bottom');
      const left = num(node, '@_w:left');
      if (top !== undefined) g.marginTopPt = top / TWIP;
      if (right !== undefined) g.marginRightPt = right / TWIP;
      if (bottom !== undefined) g.marginBottomPt = bottom / TWIP;
      if (left !== undefined) g.marginLeftPt = left / TWIP;
    }
  }
  return g;
}
