/**
 * The one place the compare worker is constructed.
 *
 * Electron loads it as a separate chunk next to the renderer. The single-file web
 * build (vite.web.config.ts) aliases this module to src/web/create-worker.ts,
 * which inlines the worker into the HTML — a separate file cannot be fetched
 * from file://.
 */
export function createCompareWorker(): Worker {
  return new Worker(new URL('./compare.worker.ts', import.meta.url), { type: 'module' });
}
