import type { RenderModel } from '@shared/render-core/types.js';

/**
 * The PDF redline of the assistant's own changes.
 *
 * Lazy on purpose. This path pulls in ~1 MB of vendored font bytes, the
 * HarfBuzz wasm module and a PDF writer — none of which the flowed redline
 * needs, and none of which anyone should pay for until they ask for a PDF. Every
 * import below is dynamic for that reason; hoisting one to the top of the file
 * would put the whole subsystem in the main chunk.
 *
 * The fonts are the four metric clones vendored in the engine (Liberation
 * Sans/Serif, Carlito, Caladea), embedded at BUILD time through the `?bytes`
 * plugin. Nothing is fetched at runtime — a legal document must not cause a
 * network request just to be typeset.
 */

export interface PdfRedlineMeta {
  readonly title: string;
  readonly author: string;
}

export async function renderPdfRedline(
  model: RenderModel,
  meta: PdfRedlineMeta,
): Promise<Uint8Array> {
  const [{ loadRenderFaces }, { createFontResolver }, { createHarfBuzzShaper }, { paginate }, { decorateLayout }, { DEFAULT_THEME }, { renderPdf }] =
    await Promise.all([
      import('@shared/render-core/assets/faces.js'),
      import('@shared/render-core/fonts/registry.js'),
      import('@shared/render-core/layout/shape.js'),
      import('@shared/render-core/paginate/paginator.js'),
      import('@shared/render-core/decorate/decorate.js'),
      import('@shared/render-core/types.js'),
      import('@shared/render-core/output/pdf.js'),
    ]);

  const fonts = createFontResolver(loadRenderFaces());
  const shaper = await createHarfBuzzShaper();

  // paginate + decorate rather than renderToLayout: the model is already built,
  // and rebuilding it from the OOXML would be the same work twice.
  const layout = paginate(model, shaper, fonts);
  decorateLayout(layout, DEFAULT_THEME);

  return renderPdf(layout, fonts, { title: meta.title, author: meta.author });
}

export function savePdf(bytes: Uint8Array, fileName: string): void {
  const blob = new Blob([bytes as unknown as BlobPart], { type: 'application/pdf' });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  URL.revokeObjectURL(url);
}

/** An object URL for an in-app viewer. The caller revokes it. */
export function pdfObjectUrl(bytes: Uint8Array): string {
  return URL.createObjectURL(new Blob([bytes as unknown as BlobPart], { type: 'application/pdf' }));
}
