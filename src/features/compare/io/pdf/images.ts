/**
 * Per-page image-presence detection.
 *
 * A text diff is BLIND to images: signatures, seals, stamps and inserted figures
 * carry no extractable text, so a "no differences" result on an execution copy
 * whose signature block changed would be dangerously misleading. This pass is a
 * blind-spot DISCLOSURE, not a difference detector and NOT a pixel diff: it lists
 * which pages contain raster/image content that the text comparison never saw.
 *
 * Render-free — it reads the page OPERATOR LIST only (no canvas, no pixels, fully
 * offline). For each image-paint op it reconstructs the current transformation
 * matrix (save/restore/transform) to recover the image's page rectangle, then:
 *   - DEDUPES repeated letterhead/logo images — an identical rectangle (same
 *     position+size) appearing on ≥2 pages is boilerplate, excluded from the
 *     notice so it fires on genuine content (a one-off signature/seal), not the
 *     letterhead on every page (P13);
 *   - flags PAGE-DOMINANT rasters (image area ≈ whole page) as the image-dominance
 *     signal feeding the P7b OCR-layer assessment (extract.ts `flagOcrLayers`).
 *
 * The OPS constants are injected ({@link PdfImageOps}) so this module never imports
 * the heavy pdfjs bundle and unit-tests against a hand-built operator list. Pure
 * except for the awaited PDF.js calls; no Office JS, no Date.now().
 */
import type {
  PdfImageOps,
  PdfOpener,
  PdfPageProxyLike,
} from './pdfjs-types.js';
import type { AbortLike } from './extract.js';

/** Tunables for image-presence detection. Named constants — no magic literals. */
export const IMAGE_PRESENCE_CONFIG = {
  /** Rectangles are rounded to this many points before dedupe comparison. */
  RECT_QUANTISE_POINTS: 2,
  /** A rectangle present on ≥ this many pages is repeated boilerplate (logo). */
  BOILERPLATE_MIN_PAGES: 2,
  /** image-area / page-area ≥ this ⇒ a page-dominant raster (OCR-layer signal). */
  DOMINANT_COVERAGE_RATIO: 0.5,
} as const;

export type ImagePresenceConfig = typeof IMAGE_PRESENCE_CONFIG;

/** The image-presence report for a whole document. */
export interface ImagePresenceReport {
  /** Pages carrying ≥1 content (non-boilerplate) image region not text-compared. */
  pagesWithContentImages: number[];
  /** Count of content image regions across the document (post-dedupe). */
  contentImageCount: number;
  /** Distinct repeated letterhead/logo images deduped out of the notice. */
  boilerplateImageCount: number;
  /** Pages dominated by a full-page raster — feeds the P7b OCR-layer flag. */
  imageDominantPages: number[];
}

/** One image draw recovered from the operator list. */
interface ImageRect {
  page: number;
  x: number;
  y: number;
  w: number;
  h: number;
  /** Fraction of the page area this image covers (0–1+). */
  coverage: number;
}

/**
 * Detect per-page image presence from raw PDF bytes (render-free).
 *
 * @param ops  the injected PDF.js OPS subset (so this module imports no bundle).
 */
export async function detectImagePresence(
  data: Uint8Array,
  opener: PdfOpener,
  ops: PdfImageOps,
  config: ImagePresenceConfig = IMAGE_PRESENCE_CONFIG,
  abortSignal?: AbortLike,
): Promise<ImagePresenceReport> {
  const doc = await opener(data);
  try {
    const rects: ImageRect[] = [];
    for (let p = 1; p <= doc.numPages; p++) {
      if (abortSignal?.aborted) throw new DOMException('PDF image scan cancelled', 'AbortError');
      const page = await doc.getPage(p);
      rects.push(...(await pageImageRects(page, p, ops)));
    }
    return summarise(rects, config);
  } finally {
    await doc.destroy();
  }
}

/**
 * Reduce recovered image rectangles to the disclosure report. Pure — exported so
 * the dedupe/dominance logic is unit-testable without opening a PDF.
 */
export function summarise(
  rects: ImageRect[],
  config: ImagePresenceConfig = IMAGE_PRESENCE_CONFIG,
): ImagePresenceReport {
  // Group by quantised rectangle key, tracking the distinct pages each appears on.
  const byKey = new Map<string, Set<number>>();
  for (const r of rects) {
    const key = rectKey(r, config.RECT_QUANTISE_POINTS);
    const pages = byKey.get(key) ?? new Set<number>();
    pages.add(r.page);
    byKey.set(key, pages);
  }
  // A rectangle on ≥ BOILERPLATE_MIN_PAGES distinct pages is repeated boilerplate.
  const boilerplateKeys = new Set<string>();
  for (const [key, pages] of byKey) {
    if (pages.size >= config.BOILERPLATE_MIN_PAGES) boilerplateKeys.add(key);
  }

  const contentPages = new Set<number>();
  let contentImageCount = 0;
  const dominantPages = new Set<number>();
  for (const r of rects) {
    if (r.coverage >= config.DOMINANT_COVERAGE_RATIO) dominantPages.add(r.page);
    if (boilerplateKeys.has(rectKey(r, config.RECT_QUANTISE_POINTS))) continue; // letterhead — excluded
    contentImageCount++;
    contentPages.add(r.page);
  }

  return {
    pagesWithContentImages: [...contentPages].sort((a, b) => a - b),
    contentImageCount,
    boilerplateImageCount: boilerplateKeys.size,
    imageDominantPages: [...dominantPages].sort((a, b) => a - b),
  };
}

// ---------------------------------------------------------------------------
// Operator-list → image rectangles (CTM reconstruction)
// ---------------------------------------------------------------------------

/** Recover every image draw on one page as a page-space rectangle. */
async function pageImageRects(
  page: PdfPageProxyLike,
  pageNumber: number,
  ops: PdfImageOps,
): Promise<ImageRect[]> {
  if (typeof page.getOperatorList !== 'function') return [];
  const ol = await page.getOperatorList();
  const viewport = page.getViewport ? page.getViewport({ scale: 1 }) : { width: 0, height: 0 };
  const pageArea = viewport.width * viewport.height;

  const paintOps = new Set<number>([
    ops.paintImageXObject,
    ops.paintInlineImageXObject,
    ops.paintImageMaskXObject,
  ]);

  const out: ImageRect[] = [];
  let ctm: Matrix = [1, 0, 0, 1, 0, 0];
  const stack: Matrix[] = [];
  for (let i = 0; i < ol.fnArray.length; i++) {
    const fn = ol.fnArray[i]!;
    if (fn === ops.save) {
      stack.push([...ctm]);
    } else if (fn === ops.restore) {
      ctm = stack.pop() ?? ctm;
    } else if (fn === ops.transform) {
      const a = ol.argsArray[i];
      if (isMatrixArgs(a)) ctm = multiply(ctm, a as Matrix);
    } else if (paintOps.has(fn)) {
      // The image occupies the unit square [0,1]² under the current matrix.
      const w = Math.hypot(ctm[0], ctm[1]);
      const h = Math.hypot(ctm[2], ctm[3]);
      const x = ctm[4];
      const y = ctm[5];
      out.push({ page: pageNumber, x, y, w, h, coverage: pageArea > 0 ? (w * h) / pageArea : 0 });
    }
  }
  return out;
}

type Matrix = [number, number, number, number, number, number];

/** 2D affine matrix product (PDF [a,b,c,d,e,f] convention). */
function multiply(m1: Matrix, m2: Matrix): Matrix {
  return [
    m1[0] * m2[0] + m1[2] * m2[1],
    m1[1] * m2[0] + m1[3] * m2[1],
    m1[0] * m2[2] + m1[2] * m2[3],
    m1[1] * m2[2] + m1[3] * m2[3],
    m1[0] * m2[4] + m1[2] * m2[5] + m1[4],
    m1[1] * m2[4] + m1[3] * m2[5] + m1[5],
  ];
}

function isMatrixArgs(a: unknown): a is Matrix {
  return Array.isArray(a) && a.length === 6 && a.every((n) => typeof n === 'number');
}

/** Quantised rectangle identity key for boilerplate dedupe. */
function rectKey(r: ImageRect, q: number): string {
  const round = (n: number): number => Math.round(n / q) * q;
  return `${round(r.x)}:${round(r.y)}:${round(r.w)}:${round(r.h)}`;
}
