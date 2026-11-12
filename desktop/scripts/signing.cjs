/**
 * What this build is allowed to sign, decided from the environment ALONE.
 *
 * Pure and CommonJS on purpose: `electron-builder.config.cjs` requires it, and
 * `tests/signing.test.ts` imports it. A signing policy that can only be checked
 * by running a release build is a signing policy nobody checks.
 *
 * The governing rule: SIGNING IS OPT-IN, never inferred.
 *
 * electron-builder will happily find a Developer ID certificate sitting in the
 * local keychain and sign with it. That is the wrong default here. A build that
 * is signed but not notarised still trips Gatekeeper, and it does so with a
 * message that reads like corruption rather than like "unsigned" — so an
 * accidental half-signed artefact is WORSE than an honest unsigned one. It also
 * means a release build differs depending on whose laptop it ran on, which is
 * the opposite of reproducible. So: nothing is signed unless MUTATIS_SIGN is
 * set, and then only what the credentials actually present can carry.
 */

/** The three credential sets @electron/notarize accepts, in its own order. */
function notaryCredentials(env) {
  if (env.APPLE_API_KEY && env.APPLE_API_KEY_ID && env.APPLE_API_ISSUER) {
    return 'App Store Connect API key';
  }
  if (env.APPLE_ID && env.APPLE_APP_SPECIFIC_PASSWORD && env.APPLE_TEAM_ID) {
    return 'Apple ID + app-specific password';
  }
  if (env.APPLE_KEYCHAIN && env.APPLE_KEYCHAIN_PROFILE) {
    return 'notarytool keychain profile';
  }
  return null;
}

/**
 * Azure Artifact Signing (the service formerly called Trusted Signing).
 *
 * Four config values name the account; the Entra credentials that authenticate
 * to it are read from the environment by the Azure SDK itself, never by us and
 * never written into the config. See docs/signing.md for who may enrol.
 */
function azureSigning(env) {
  const endpoint = env.AZURE_CODE_SIGNING_ENDPOINT;
  const account = env.AZURE_CODE_SIGNING_ACCOUNT;
  const profile = env.AZURE_CODE_SIGNING_PROFILE;
  const publisher = env.AZURE_CODE_SIGNING_PUBLISHER;
  if (!endpoint || !account || !profile || !publisher) return null;
  // EnvironmentCredential's own required trio. Missing it fails deep inside the
  // signing step with an opaque error, so refuse up front instead.
  if (!env.AZURE_TENANT_ID || !env.AZURE_CLIENT_ID || !env.AZURE_CLIENT_SECRET) {
    return { incomplete: 'AZURE_TENANT_ID, AZURE_CLIENT_ID and AZURE_CLIENT_SECRET' };
  }
  return {
    endpoint,
    codeSigningAccountName: account,
    certificateProfileName: profile,
    // Must match the certificate subject EXACTLY. electron-updater re-checks it
    // on the downloaded installer before running it, so a mismatch here does not
    // break the build — it breaks auto-update, later, silently.
    publisherName: publisher,
  };
}

/**
 * The update feed.
 *
 * Absent by default, and absence is meaningful: with no publish config,
 * electron-builder writes no `app-update.yml` into the bundle, so the shipped
 * app has no feed URL to call even if someone wired a check to startup.
 */
function publishConfig(env) {
  const url = env.MUTATIS_UPDATE_URL;
  if (!url) return null;
  if (!/^https:\/\//.test(url)) {
    return { invalid: 'MUTATIS_UPDATE_URL must be https://' };
  }
  return { provider: 'generic', url, channel: 'latest' };
}

function resolveSigning(env = {}) {
  const notes = [];
  const wanted = env.MUTATIS_SIGN === '1';

  if (!wanted) {
    notes.push('MUTATIS_SIGN is not 1 — building UNSIGNED on every platform.');
  }

  // --- macOS ---------------------------------------------------------------
  const macSign = wanted;
  const notary = notaryCredentials(env);
  const notarize = macSign && notary !== null;
  if (wanted && notary === null) {
    notes.push(
      'macOS: signing, NOT notarising — no notary credentials in the environment. ' +
        'Gatekeeper will still block this build on another machine.',
    );
  } else if (notarize) {
    notes.push(`macOS: signing and notarising via ${notary}.`);
  }

  // --- Windows -------------------------------------------------------------
  const azure = wanted ? azureSigning(env) : null;
  let winAzure = null;
  if (wanted) {
    if (azure === null) {
      notes.push(
        'Windows: UNSIGNED — no Azure Artifact Signing account configured. ' +
          'SmartScreen will warn on first run.',
      );
    } else if (azure.incomplete) {
      notes.push(`Windows: UNSIGNED — account is configured but ${azure.incomplete} are missing.`);
    } else {
      winAzure = azure;
      notes.push(`Windows: signing via Azure Artifact Signing (${azure.endpoint}).`);
    }
  }

  // --- Update feed ---------------------------------------------------------
  const publish = publishConfig(env);
  let resolvedPublish = null;
  if (publish === null) {
    notes.push('No MUTATIS_UPDATE_URL — no app-update.yml, so auto-update stays off.');
  } else if (publish.invalid) {
    notes.push(`Update feed IGNORED — ${publish.invalid}.`);
  } else {
    resolvedPublish = publish;
    notes.push(`Update feed: ${publish.url}`);
  }

  // --- Targets -------------------------------------------------------------
  // `dir` unsigned, because that is what the three packaged audit scripts read
  // and it costs seconds rather than minutes. A signed build produces the real
  // artefacts — and on macOS the `zip` is not optional: Squirrel.Mac consumes a
  // zip, so a dmg-only release has no update path at all.
  const macTarget = macSign ? ['dmg', 'zip'] : ['dir'];
  const winTarget = winAzure ? ['nsis'] : ['dir'];

  return {
    macSign,
    notarize,
    notaryMethod: notary,
    winAzure,
    publish: resolvedPublish,
    macTarget,
    winTarget,
    notes,
  };
}

module.exports = { resolveSigning };
