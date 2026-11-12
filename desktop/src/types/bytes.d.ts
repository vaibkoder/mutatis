/**
 * The `?bytes` import suffix, as TypeScript sees it.
 *
 * `scripts/bytes-plugin.mjs` turns `import x from '@fonts/Foo.ttf?bytes'` into a
 * `Uint8Array` inlined at BUILD time — the mechanism that lets RenderCore embed
 * sixteen font faces with zero runtime fetch. The compiler needs telling; the
 * bundler already knows.
 */
declare module '*?bytes' {
  const bytes: Uint8Array;
  export default bytes;
}
