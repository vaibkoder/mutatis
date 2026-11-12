/**
 * `fontkit`, as this project's compiler sees it.
 *
 * Not `@types/fontkit`. That package's `Font` type is structurally incompatible
 * with the `Fontkit` interface `@pdfme/pdf-lib` declares for
 * `PDFDocument.registerFontkit`, which makes `render-core/output/pdf.ts` — a
 * Mutatis file this repo may never edit — fail to compile. Upstream builds
 * clean only because it has no fontkit types at all.
 *
 * So we reproduce that deliberately, with an explicit `any` rather than an
 * implicit one. Upstream does `import * as fontkit`, so the NAMESPACE has to be
 * assignable, which an export-default shim is not. `fontkit` is passed straight
 * through to pdf-lib and never called by our own code, so there is nothing here
 * for types to protect.
 */
declare module 'fontkit' {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const value: any;
  export = value;
}
