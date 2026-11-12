import { readFileSync } from 'node:fs';

/**
 * Vite/vitest plugin: resolve `<asset>?bytes` imports to an in-process
 * `Uint8Array` module (base64 → bytes), decoded with `atob` at module-eval time.
 *
 * RenderCore (feature 16 / ADR-0024) embeds its binary assets — the OFL clone
 * fonts and the HarfBuzz wasm — at BUILD time so the render path never issues a
 * runtime `fetch`/XHR for them (zero network egress; ADR-0001/0003, and the
 * `no-network-egress` lint bans `fetch` in `src/`). The same plugin runs in
 * `vite build` (both rollup inputs) and in `vitest`, so tests exercise the real
 * byte-loading path. Assets are imported behind a lazy dynamic `import()` so
 * non-render flows never pull the bytes into their chunk (acceptance 10).
 */
const SUFFIX = '?bytes';

export function bytesPlugin() {
  return {
    name: 'mutatis-bytes',
    enforce: 'pre',
    async resolveId(id, importer) {
      if (!id.endsWith(SUFFIX)) return null;
      const clean = id.slice(0, -SUFFIX.length);
      const resolved = await this.resolve(clean, importer, { skipSelf: true });
      return (resolved?.id ?? clean) + SUFFIX;
    },
    load(id) {
      if (!id.endsWith(SUFFIX)) return null;
      const file = id.slice(0, -SUFFIX.length);
      const b64 = readFileSync(file).toString('base64');
      this.addWatchFile?.(file);
      return {
        code: `export default Uint8Array.from(atob(${JSON.stringify(b64)}), (c) => c.charCodeAt(0));`,
        map: null,
      };
    },
  };
}
