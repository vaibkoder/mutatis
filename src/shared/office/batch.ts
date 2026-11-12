/**
 * context.sync batching helpers.
 * Rule: never call ctx.sync() inside a per-item loop.
 * Batch multiple loads and sync once. (spec/ARCHITECTURE.md §Performance Guidelines)
 */

/** Run a Word.run context, with a single sync for all loads queued inside fn. */
export async function wordRun<T>(
  fn: (ctx: Word.RequestContext) => Promise<T>
): Promise<T> {
  return Word.run(fn);
}

/**
 * Load multiple proxy objects in one sync. Pass an array of [proxy, properties] pairs.
 * Syncs once at the end.
 */
export async function batchLoad(
  ctx: Word.RequestContext,
  loads: Array<{ proxy: { load(props: string): void }; props: string }>
): Promise<void> {
  for (const { proxy, props } of loads) {
    proxy.load(props);
  }
  await ctx.sync();
}

/**
 * Process items in chunks to avoid holding too many ranges in memory and to
 * keep sync counts bounded. (spec/ARCHITECTURE.md §Performance Guidelines)
 */
export async function processInChunks<T, R>(
  items: T[],
  chunkSize: number,
  process: (chunk: T[], ctx: Word.RequestContext) => Promise<R[]>
): Promise<R[]> {
  const results: R[] = [];
  for (let i = 0; i < items.length; i += chunkSize) {
    const chunk = items.slice(i, i + chunkSize);
    const chunkResults = await wordRun(ctx => process(chunk, ctx));
    results.push(...chunkResults);
  }
  return results;
}
