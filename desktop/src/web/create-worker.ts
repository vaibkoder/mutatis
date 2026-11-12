// Web build's replacement for renderer/src/compare/create-worker.ts (aliased in
// vite.web.config.ts). `?worker&inline` embeds the worker bundle in the HTML as
// a blob, so nothing is fetched from file://.
import CompareWorker from '../renderer/src/compare/compare.worker.ts?worker&inline';

export function createCompareWorker(): Worker {
  return new CompareWorker();
}
