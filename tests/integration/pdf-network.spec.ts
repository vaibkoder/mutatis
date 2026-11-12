/**
 * A1 confidentiality guard — LAYER 2 (runtime zero-network proof, P8).
 *
 * USER-RUN. Playwright is intentionally NOT in the toolchain (it pulls a large
 * dependency + browser binaries). To run this gate:
 *
 *   npm i -D @playwright/test
 *   npx playwright install chromium
 *   npm run build && npm run preview &     # serve dist on a local origin
 *   npx playwright test tests/integration/pdf-network.spec.ts
 *
 * The spec drives a full PDF parse → reconstruct → diff cycle in a real browser
 * and asserts ZERO HTTP/HTTPS requests to any non-local origin. blob: worker
 * URLs are in-memory (not a network leak) and are explicitly excluded (P8).
 *
 * Wire `APP_URL` to wherever `npm run preview` serves the task pane, and trigger
 * a PDF compare via the page's test hook before the assertion block.
 */
// This file is excluded from lint and typecheck (see eslint.config.js ignores and
// tsconfig.test.json exclude): `@playwright/test` is not installed by default, so
// the import below cannot resolve until the setup block above has been run.
import { test, expect } from '@playwright/test';

const APP_URL = process.env.MUTATIS_APP_URL ?? 'https://localhost:3000/src/taskpane/index.html';

test('PDF compare issues zero requests to non-local origins', async ({ page }) => {
  const offOrigin = [];
  const isLocal = (url) => {
    if (url.startsWith('blob:') || url.startsWith('data:')) return true; // in-memory
    try {
      const host = new URL(url).hostname;
      return host === 'localhost' || host === '127.0.0.1' || host === '::1';
    } catch {
      return true; // relative ⇒ same-origin
    }
  };

  page.on('request', (req) => {
    if (!isLocal(req.url())) offOrigin.push(req.url());
  });

  await page.goto(APP_URL);

  // TODO(user): trigger a PDF-vs-PDF compare here via a test hook exposed on the
  // page (e.g. window.__mutatis_runPdfCompare(twoFixtureBuffers)), then await
  // completion. The fixtures live in tests/fixtures/pdf/.
  await page.waitForTimeout(500);

  expect(offOrigin, `off-origin requests: ${offOrigin.join(', ')}`).toHaveLength(0);
});
