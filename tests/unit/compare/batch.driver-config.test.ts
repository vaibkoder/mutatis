/**
 * The variant cap the panel shows is the cap the runner enforces (spec §7).
 *
 * Its own file because it mocks the config module, and `vi.mock` is
 * file-scoped: hoisting it into `batch.driver.test.ts` would give every other
 * test in there a config it did not ask for.
 *
 * This exists because a review replaced the driver's resolution with a
 * hard-coded `const maxVariants = 100;` — precisely the defect the wiring was
 * added to fix — and all 353 compare and web tests still passed.
 * `resolveBatchConfig` was tested in isolation and the runner's cap was tested,
 * but nothing pinned that the DRIVER reads the resolved value. A panel that
 * promises a hundred while the runner refuses past three is a batch that fails
 * at document four with no warning anyone could have read.
 */
import { describe, it, expect, vi } from 'vitest';

vi.mock('@shared/config/index.js', async importOriginal => {
  const actual = await importOriginal<typeof import('@shared/config/index.js')>();
  return {
    ...actual,
    getCompareConfig: (profile: Parameters<typeof actual.getCompareConfig>[0]) => ({
      ...actual.getCompareConfig(profile),
      batch: {
        maxVariants: 3,
        // Relaxed far enough that a contested pair pre-ticks, which is
        // something the defaults will not do — see the bundle-driver test.
        pairing: { autoPairScore: 0.5, autoPairMargin: 0.01 },
      },
    }),
  };
});

const { createBatchDriver } = await import('@features/compare/batch/driver.js');
// From `limits.ts`, which owns it. Imported from `constants.js` this bound
// `undefined`, and the guard below became `expect(undefined).not.toBe(3)` — a
// test that could not fail, invisible because tests are outside `tsconfig`.
const { MAX_BATCH_VARIANTS } = await import('@features/compare/batch/limits.js');

describe('the driver resolves its cap from config, not from a constant', () => {
  it('reports the configured cap', () => {
    expect(createBatchDriver().maxVariants).toBe(3);
  });

  it('and the configured cap is not merely the default wearing a hat', () => {
    // Without this the assertion above would still pass if the driver ignored
    // config and the fixture happened to match the built-in default.
    expect(MAX_BATCH_VARIANTS).not.toBe(3);
  });

  it('refuses the run at the same number the panel promises', async () => {
    // The two halves have to agree. The cap is enforced in the runner and
    // displayed from the driver; if they came from different places, the only
    // place a user would find out is a failed batch.
    const { synthesiseDocx } = await import('@features/compare/io/text/synthesise-docx.js');
    const file = (name: string, text: string): File =>
      new File([new Uint8Array(synthesiseDocx(text).toUint8Array())], name);

    const driver = createBatchDriver();
    driver.setTemplate(file('house.docx', 'The Seller shall pay on Completion.'));
    driver.addVariants(
      Array.from({ length: driver.maxVariants + 1 }, (_, i) =>
        file(`v${i}.docx`, `The Seller shall pay within ${i + 1} days.`)),
    );

    await driver.runBatch();

    expect(driver.status.value).toBe('error');
    expect(driver.errorMessage.value).toContain(String(driver.maxVariants));
  });
});

describe('the bundle driver hands its config to the pairing', () => {
  // The other half of the same gap. `proposePairing` accepts a config and the
  // pairing modules consume it, but a driver that simply never passed it would
  // leave every profile setting inert again — and the import wall in the
  // acceptance sweep cannot see a missing ARGUMENT, only a forbidden import.
  it('proposes under the profile\'s thresholds, not the built-in ones', async () => {
    const { createBundleDriver } = await import('@features/compare/batch/bundle-driver.js');
    const { synthesiseDocx } = await import('@features/compare/io/text/synthesise-docx.js');
    const { proposePairing } = await import('@features/compare/batch/pairing/propose.js');

    const body = (extra: string): string => [
      'This Share Purchase Agreement is made between the Seller and the Buyer.',
      'The Seller shall sell and the Buyer shall purchase the Sale Shares.',
      'The Purchase Price shall be paid in full on the Completion Date.',
      'The Seller warrants that it has good title to the Sale Shares.',
      extra,
    ].join('\n');

    const file = (name: string, text: string): File =>
      new File([new Uint8Array(synthesiseDocx(text).toUint8Array())], name, {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });

    const driver = createBundleDriver();
    driver.addDocuments('a', [file('SPA v6.docx', body('This Agreement is governed by English law.'))]);
    driver.addDocuments('b', [
      file('SPA v7.docx', body('This Agreement is governed by the laws of England and Wales.')),
      file('SPA early draft.docx', body('The parties may agree a different completion mechanic.')),
    ]);

    await driver.proposePairs();

    const row = driver.proposal.value?.rows.find(r => r.a !== null && r.b !== null);
    expect(row?.band).toBe('auto-pair');

    // Control: the SAME documents proposed under the feature defaults do not
    // pre-tick. Without this, "the band was auto-pair" could just mean the
    // documents were an easy match, and the test would prove nothing about the
    // config reaching the pairing.
    const { buildPairingSketch } = await import('@features/compare/batch/pairing/sketch.js');
    const { getCompareConfig, UK_PROFILE } = await import('@shared/config/index.js');
    const { DEFAULT_PAIRING_SETTINGS } = await import('@features/compare/batch/config.js');

    // The mocked config with its `batch` block removed — i.e. the profile as it
    // would be if it said nothing about pairing at all.
    const plain = { ...getCompareConfig(UK_PROFILE) };
    delete (plain as { batch?: unknown }).batch;

    const sketchOf = async (f: File, side: 'a' | 'b', id: string) => buildPairingSketch({
      docId: id,
      side,
      fileName: f.name,
      bytes: await f.arrayBuffer(),
      config: plain,
      profile: UK_PROFILE,
      flatten: 'acceptAll',
    });

    const a6 = file('SPA v6.docx', body('This Agreement is governed by English law.'));
    const b7 = file('SPA v7.docx', body('This Agreement is governed by the laws of England and Wales.'));
    const bDraft = file('SPA early draft.docx', body('The parties may agree a different completion mechanic.'));

    const underDefaults = proposePairing({
      sideA: [await sketchOf(a6, 'a', 'a1')],
      sideB: [await sketchOf(b7, 'b', 'b1'), await sketchOf(bDraft, 'b', 'b2')],
      generatedAt: '2026-09-11T00:00:00.000Z',
    });

    expect(DEFAULT_PAIRING_SETTINGS.autoPairScore).toBeGreaterThan(0.5);
    expect(underDefaults.rows.find(r => r.a !== null && r.b !== null)?.band).toBe('propose-with-rivals');
  });
});
