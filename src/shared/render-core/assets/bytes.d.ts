/**
 * Ambient type for the `?bytes` asset imports (RenderCore, ADR-0024 Phase 2).
 * Resolved by the Vite/vitest `bytesPlugin` (`scripts/bytes-plugin.mjs`) to a
 * `Uint8Array` decoded in-process — no runtime `fetch` (zero egress).
 */
declare module '*?bytes' {
  const bytes: Uint8Array;
  export default bytes;
}
