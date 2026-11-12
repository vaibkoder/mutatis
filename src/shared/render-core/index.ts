/**
 * RenderCore public API (feature 16, ADR-0024) — host-free, DOM-free engine.
 *
 * One shared layout result from a Compare change-set; Phase 2 paints it as a
 * preview and a vector PDF and wires it behind the Compare/preview surfaces.
 * The engine takes font BYTES + a Shaper by injection (no asset fetching here).
 */
export * from './types.js';
export { buildRenderModel, type RenderInput, type RenderModelInput } from './input/index.js';
export { createFontResolver, type FontResolver, type FaceInput, type ResolvedFace, CLONE_MAP } from './fonts/registry.js';
export { readLineMetrics, singleLineHeightPt } from './fonts/sfnt-metrics.js';
export { deobfuscateOdttf, parseFontKey } from './fonts/odttf.js';
export { createHarfBuzzShaper, type Shaper } from './layout/shape.js';
export { layoutParagraph, type ParagraphLayout } from './layout/paragraph.js';
export { paginate } from './paginate/paginator.js';
export { decorateLayout } from './decorate/decorate.js';
export { renderPdf, type PdfMetadata } from './output/pdf.js';
export { renderPreview, type PreviewOptions, type PreviewHandle } from './output/preview.js';
export * from './worker/index.js';

import type { RenderInput } from './input/index.js';
import { buildRenderModel } from './input/index.js';
import { paginate } from './paginate/paginator.js';
import { decorateLayout } from './decorate/decorate.js';
import type { FontResolver } from './fonts/registry.js';
import type { Shaper } from './layout/shape.js';
import { DEFAULT_THEME, type LayoutResult, type RenderTheme } from './types.js';

export interface RenderOptions {
  theme?: RenderTheme;
}

/**
 * Build → paginate → decorate, returning the single shared layout result that
 * both painters (preview, PDF) consume so they paginate identically.
 */
export function renderToLayout(
  input: RenderInput,
  fonts: FontResolver,
  shaper: Shaper,
  opts: RenderOptions = {},
): LayoutResult {
  const model = buildRenderModel(input);
  const layout = paginate(model, shaper, fonts);
  decorateLayout(layout, opts.theme ?? DEFAULT_THEME);
  return layout;
}
