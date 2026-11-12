import { defineConfig, type Plugin } from 'vite';
import react from '@vitejs/plugin-react';
import tailwindcss from '@tailwindcss/vite';
import { viteSingleFile } from 'vite-plugin-singlefile';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { bytesPlugin } from './scripts/bytes-plugin.mjs';
import { engineAliases, engineDedupe } from './scripts/engine-aliases.mjs';

/**
 * The single-file web build: out/web/Mutatis.html, no AI, no chat, no network.
 * Sibling of the `renderer` block in electron.vite.config.ts, which it must not
 * disturb.
 */

const here = import.meta.dirname;

/**
 * Swap two renderer modules for their web versions without editing either.
 * Matched on the RESOLVED file, so every spelling of the import is caught and
 * the shared compare hook and the web settings dialog end up on one module.
 */
const REDIRECTS = new Map<string, string>([
  [
    resolve(here, 'src/renderer/src/settings/SettingsProvider.tsx'),
    resolve(here, 'src/web/WebSettingsProvider.tsx'),
  ],
  [
    resolve(here, 'src/renderer/src/compare/create-worker.ts'),
    resolve(here, 'src/web/create-worker.ts'),
  ],
  [
    resolve(here, 'src/renderer/src/pdf/open-pdf.ts'),
    resolve(here, 'src/web/open-pdf.ts'),
  ],
]);

function redirectPlugin(): Plugin {
  return {
    name: 'mutatis-web-redirect',
    enforce: 'pre',
    async resolveId(id, importer, options) {
      if (!importer || !id.startsWith('.')) return null;
      const resolved = await this.resolve(id, importer, { ...options, skipSelf: true });
      const target = resolved ? REDIRECTS.get(resolved.id) : undefined;
      return target ?? null;
    },
  };
}

/**
 * harfbuzzjs finds its wasm with `new URL("harfbuzz.wasm", import.meta.url).href`,
 * a sibling file that cannot be fetched from file://. Replace that expression
 * with a data: URL of the same bytes. Scoped to harfbuzzjs's own module, so
 * nothing else is touched; node_modules and the engine stay as they are.
 */
const HB_WASM_EXPRESSION = 'new URL("harfbuzz.wasm",import.meta.url).href';

function harfbuzzWasmPlugin(): Plugin {
  return {
    name: 'mutatis-web-harfbuzz-wasm',
    enforce: 'pre',
    transform(code, id) {
      const file = id.split('?')[0] ?? id;
      if (!/[\\/]harfbuzzjs[\\/]dist[\\/]harfbuzz\.js$/.test(file)) return null;
      if (!code.includes(HB_WASM_EXPRESSION)) {
        this.error(
          'harfbuzzjs no longer contains the wasm URL expression this build rewrites; ' +
            'update harfbuzzWasmPlugin in vite.web.config.ts',
        );
      }
      const wasm = readFileSync(resolve(file, '../harfbuzz.wasm')).toString('base64');
      return {
        code: code.replace(HB_WASM_EXPRESSION, JSON.stringify(`data:application/wasm;base64,${wasm}`)),
        map: null,
      };
    },
  };
}

/** vite-plugin-singlefile leaves the page as index.html; the deliverable is Mutatis.html. */
function renameOutputPlugin(): Plugin {
  return {
    name: 'mutatis-web-rename',
    enforce: 'post',
    generateBundle: {
      order: 'post',
      handler(_options, bundle) {
        const page = bundle['index.html'];
        if (!page) this.error('expected index.html in the bundle');
        delete bundle['index.html'];
        page.fileName = 'Mutatis.html';
        bundle['Mutatis.html'] = page;
      },
    },
  };
}

export default defineConfig({
  root: resolve(here, 'src/web'),
  // Relative, so nothing in the page is addressed from "/" — there is no server.
  base: './',
  plugins: [
    redirectPlugin(),
    harfbuzzWasmPlugin(),
    react(),
    tailwindcss(),
    bytesPlugin(),
    viteSingleFile(),
    renameOutputPlugin(),
  ],
  resolve: {
    alias: {
      '@': resolve(here, 'src/renderer/src'),
      '@common': resolve(here, 'src/common'),
      ...engineAliases,
    },
    dedupe: engineDedupe,
  },
  worker: {
    format: 'iife',
    plugins: () => [redirectPlugin(), harfbuzzWasmPlugin(), bytesPlugin()],
  },
  build: {
    target: 'es2022',
    outDir: resolve(here, 'out/web'),
    emptyOutDir: true,
    // Everything is inlined; there is nothing to preload from.
    modulePreload: false,
    chunkSizeWarningLimit: 20_000,
  },
});
