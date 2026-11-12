import type { Plugin } from 'vite';

/**
 * Resolves `<asset>?bytes` imports to an in-process `Uint8Array` module.
 *
 * Copied from Mutatis's scripts/bytes-plugin.mjs. RenderCore embeds its
 * binary assets (the OFL clone fonts, the HarfBuzz wasm) at BUILD time so the
 * render path never issues a runtime fetch for them. Every `?bytes` import
 * fails without this plugin, and it fails silently — so it is wired into the
 * renderer build, the worker build and the test runner.
 */
export declare function bytesPlugin(): Plugin;
