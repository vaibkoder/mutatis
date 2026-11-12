/**
 * Packaging configuration.
 *
 * This was `electron-builder.yml` until phase 8. It is JavaScript now because
 * signing is a DECISION taken from the environment, and a decision taken in
 * YAML can only be inspected by running a release build. The decision itself
 * lives in `scripts/signing.cjs`, which is pure and unit-tested; this file only
 * assembles it.
 *
 * Note the shape of `.yml` vs `.cjs` in electron-builder's own discovery order:
 * yml is checked FIRST. So the old file is deleted, not left beside this one —
 * had both existed, every edit here would have been silently ignored.
 */
const { resolveSigning } = require('./scripts/signing.cjs');

const signing = resolveSigning(process.env);
for (const note of signing.notes) console.log(`[packaging] ${note}`);

module.exports = {
  appId: 'com.mutatis.desktop',
  productName: 'Mutatis Desktop',
  copyright: 'Copyright © 2026 the Mutatis authors',

  directories: {
    output: 'release',
    buildResources: 'build',
  },

  // `out` is electron-vite's output. Everything else in the repo — sources,
  // tests, the sibling engine checkout — stays out of the package.
  files: [
    'out/**/*',
    'package.json',
    // Apache-2.0 §4 requires the licence text to travel WITH the distribution,
    // and the AI SDK packages ship only the short-form header. See
    // THIRD-PARTY-NOTICES.md.
    'licenses/**/*',
    'THIRD-PARTY-NOTICES.md',
  ],

  asar: true,

  // Only present when MUTATIS_UPDATE_URL is set. With no publish config
  // electron-builder writes no app-update.yml, and the shipped app therefore
  // holds no feed URL at all — which is the state we want by default.
  ...(signing.publish ? { publish: signing.publish } : {}),

  mac: {
    category: 'public.app-category.productivity',
    target: signing.macTarget,
    // `null` is not "use the default" — it is an explicit refusal to sign, and
    // it stops electron-builder picking up whatever Developer ID happens to be
    // in the local keychain. See scripts/signing.cjs for why that matters.
    identity: signing.macSign ? undefined : null,
    hardenedRuntime: true,
    gatekeeperAssess: false,
    // Narrower than @electron/osx-sign's defaults, deliberately. Read the
    // comment at the top of either file.
    entitlements: 'build/entitlements.mac.plist',
    entitlementsInherit: 'build/entitlements.mac.inherit.plist',
    notarize: signing.notarize,
    // 11.0 is Electron 44's own floor; stating it keeps the Info.plist honest
    // rather than letting it drift with the Electron bump.
    minimumSystemVersion: '11.0',
  },

  win: {
    target: signing.winTarget,
    ...(signing.winAzure ? { azureSignOptions: signing.winAzure } : {}),
    // electron-updater re-verifies the Authenticode publisher of a downloaded
    // installer before running it. Left at its default (true) on purpose: an
    // update path that does not check who signed the thing it is about to
    // execute is worse than having no update path.
    verifyUpdateCodeSignature: true,
  },

  linux: {
    target: ['dir'],
    category: 'Office',
  },
};
