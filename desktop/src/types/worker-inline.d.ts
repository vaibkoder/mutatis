/**
 * Vite's `?worker&inline` import, as TypeScript sees it. Used only by the
 * single-file web build (src/web/create-worker.ts), where the worker has to
 * live inside the HTML rather than beside it.
 */
declare module '*?worker&inline' {
  const WorkerConstructor: new (options?: { name?: string }) => Worker;
  export default WorkerConstructor;
}

declare module 'pdfjs-dist/build/pdf.worker.min.mjs';
