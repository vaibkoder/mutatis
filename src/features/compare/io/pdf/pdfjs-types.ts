/**
 * Minimal STRUCTURAL types for the slice of the PDF.js API the Compare PDF path
 * consumes. Declaring our own shapes (rather than importing
 * `pdfjs-dist` types) keeps the extraction/reconstruction layers decoupled from
 * the heavy bundle, so they unit-test against a tiny Node-opened doc or a hand
 * built stub — no browser, no worker. The real `PDFDocumentProxy` /
 * `PDFPageProxy` are structurally assignable to these.
 */

/** One text item from `page.getTextContent()`. */
export interface PdfTextItem {
  /** The glyph run's string (may be '' for a pure line-break marker). */
  str: string;
  /** PDF.js end-of-line flag for this item (producer-dependent). */
  hasEOL: boolean;
  /**
   * 6-element text transform [a, b, c, d, e, f]. For unrotated text a≈d≈fontSize,
   * e = x (PDF user space), f = y (origin bottom-left). b/c encode skew/rotation.
   */
  transform: number[];
  /** Advance width of the run in text-space units. */
  width: number;
  /** Glyph height in text-space units. */
  height: number;
  /** Internal font id (e.g. 'g_d0_f1') — only used as a weak grouping signal. */
  fontName?: string;
}

/** A marked-content / non-text item — present in the items array; we skip it. */
export interface PdfMarkedItem {
  type: string;
}

export type PdfContentItem = PdfTextItem | PdfMarkedItem;

/** Narrow a content item to a real text item. */
export function isTextItem(item: PdfContentItem): item is PdfTextItem {
  return typeof (item as PdfTextItem).str === 'string';
}

export interface PdfTextContent {
  items: PdfContentItem[];
}

export interface PdfPageProxyLike {
  getTextContent(): Promise<PdfTextContent>;
  /** Operator list — used by D3 image-presence detection (no rendering). */
  getOperatorList?(): Promise<{ fnArray: number[]; argsArray: unknown[] }>;
  /** Page size at scale 1 — used by D3 to measure image coverage (no rendering). */
  getViewport?(params: { scale: number }): { width: number; height: number };
  /** Page resources accessor (alternative image-presence path). */
  readonly objs?: unknown;
}

/**
 * The subset of PDF.js `OPS` constants the D3 image-presence pass needs. Injected
 * (not hardcoded) so images.ts stays pure of the heavy bundle and unit-testable
 * with a tiny stub. The production loader + the Node test opener both source
 * these from the real `pdfjs.OPS` map.
 */
export interface PdfImageOps {
  save: number;
  restore: number;
  transform: number;
  paintImageXObject: number;
  paintInlineImageXObject: number;
  paintImageMaskXObject: number;
}

export interface PdfDocumentProxyLike {
  readonly numPages: number;
  getPage(pageNumber: number): Promise<PdfPageProxyLike>;
  destroy(): Promise<void>;
}

/**
 * Opens raw PDF bytes into a document proxy. The production implementation
 * (`pdfjs-loader.ts`) is browser/Vite-specific and confidentiality-hardened;
 * tests inject a Node-configured opener instead. Dependency-injected wherever
 * extraction needs to open a document, so no layer hardcodes the browser opener.
 */
export type PdfOpener = (data: Uint8Array) => Promise<PdfDocumentProxyLike>;
