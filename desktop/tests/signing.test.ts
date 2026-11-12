import { describe, expect, it } from 'vitest';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const { resolveSigning } = require('../scripts/signing.cjs') as {
  resolveSigning: (env: Record<string, string>) => {
    macSign: boolean;
    notarize: boolean;
    notaryMethod: string | null;
    winAzure: Record<string, string> | null;
    publish: { provider: string; url: string } | null;
    macTarget: string[];
    winTarget: string[];
    notes: string[];
  };
};

/**
 * The signing policy is the one part of packaging that cannot be checked by
 * looking at the artefact: by the time you can see whether a build was signed,
 * it is already built. So the decision is a pure function, and this is where it
 * gets its adversarial reading.
 */
describe('resolveSigning', () => {
  const SIGN = { MUTATIS_SIGN: '1' };
  const APPLE_KEY = {
    APPLE_API_KEY: '/keys/AuthKey.p8',
    APPLE_API_KEY_ID: 'ABCD1234',
    APPLE_API_ISSUER: '00000000-0000-0000-0000-000000000000',
  };

  it('signs nothing by default', () => {
    const result = resolveSigning({});
    expect(result.macSign).toBe(false);
    expect(result.notarize).toBe(false);
    expect(result.winAzure).toBeNull();
    expect(result.macTarget).toEqual(['dir']);
  });

  // The load-bearing one. A developer machine with a Developer ID in its
  // keychain must not quietly produce a different artefact from CI.
  it('does not sign just because Apple credentials are lying around', () => {
    const result = resolveSigning({ ...APPLE_KEY, APPLE_TEAM_ID: 'TEAMID', CSC_LINK: '/c.p12' });
    expect(result.macSign).toBe(false);
    expect(result.notarize).toBe(false);
  });

  it('signs but refuses to claim notarisation without notary credentials', () => {
    const result = resolveSigning({ ...SIGN });
    expect(result.macSign).toBe(true);
    expect(result.notarize).toBe(false);
    expect(result.notes.join(' ')).toMatch(/Gatekeeper will still block/);
  });

  it('notarises with an App Store Connect API key', () => {
    const result = resolveSigning({ ...SIGN, ...APPLE_KEY });
    expect(result.notarize).toBe(true);
    expect(result.notaryMethod).toBe('App Store Connect API key');
  });

  it('accepts the Apple ID and keychain-profile credential sets too', () => {
    expect(
      resolveSigning({
        ...SIGN,
        APPLE_ID: 'a@b.c',
        APPLE_APP_SPECIFIC_PASSWORD: 'x',
        APPLE_TEAM_ID: 'T',
      }).notarize,
    ).toBe(true);
    expect(
      resolveSigning({ ...SIGN, APPLE_KEYCHAIN: 'login', APPLE_KEYCHAIN_PROFILE: 'p' }).notarize,
    ).toBe(true);
  });

  it('will not half-configure Azure signing', () => {
    // Account named, but the Entra credentials the Azure SDK reads are absent.
    // Without this guard the failure happens minutes later, inside the signer.
    const result = resolveSigning({
      ...SIGN,
      AZURE_CODE_SIGNING_ENDPOINT: 'https://neu.codesigning.azure.net',
      AZURE_CODE_SIGNING_ACCOUNT: 'acct',
      AZURE_CODE_SIGNING_PROFILE: 'profile',
      AZURE_CODE_SIGNING_PUBLISHER: 'CN=Example',
    });
    expect(result.winAzure).toBeNull();
    expect(result.notes.join(' ')).toMatch(/AZURE_TENANT_ID/);
  });

  it('configures Azure signing when the whole set is present', () => {
    const result = resolveSigning({
      ...SIGN,
      AZURE_CODE_SIGNING_ENDPOINT: 'https://neu.codesigning.azure.net',
      AZURE_CODE_SIGNING_ACCOUNT: 'acct',
      AZURE_CODE_SIGNING_PROFILE: 'profile',
      AZURE_CODE_SIGNING_PUBLISHER: 'CN=Example',
      AZURE_TENANT_ID: 't',
      AZURE_CLIENT_ID: 'c',
      AZURE_CLIENT_SECRET: 's',
    });
    expect(result.winAzure).toMatchObject({
      codeSigningAccountName: 'acct',
      certificateProfileName: 'profile',
      publisherName: 'CN=Example',
    });
    expect(result.winTarget).toEqual(['nsis']);
  });

  it('never puts a secret into the build config', () => {
    const result = resolveSigning({
      ...SIGN,
      ...APPLE_KEY,
      AZURE_CODE_SIGNING_ENDPOINT: 'https://neu.codesigning.azure.net',
      AZURE_CODE_SIGNING_ACCOUNT: 'acct',
      AZURE_CODE_SIGNING_PROFILE: 'profile',
      AZURE_CODE_SIGNING_PUBLISHER: 'CN=Example',
      AZURE_TENANT_ID: 't',
      AZURE_CLIENT_ID: 'c',
      AZURE_CLIENT_SECRET: 'super-secret-value',
      CSC_KEY_PASSWORD: 'another-secret',
    });
    const serialised = JSON.stringify(result);
    expect(serialised).not.toContain('super-secret-value');
    expect(serialised).not.toContain('another-secret');
    expect(serialised).not.toContain('/keys/AuthKey.p8');
  });

  it('ships a zip alongside the dmg, because Squirrel.Mac cannot read a dmg', () => {
    expect(resolveSigning({ ...SIGN }).macTarget).toEqual(['dmg', 'zip']);
  });

  it('refuses a non-https update feed', () => {
    const result = resolveSigning({ MUTATIS_UPDATE_URL: 'http://updates.example.com' });
    expect(result.publish).toBeNull();
    expect(result.notes.join(' ')).toMatch(/must be https/);
  });

  it('configures a generic feed when one is given', () => {
    const result = resolveSigning({ MUTATIS_UPDATE_URL: 'https://updates.example.com/dp' });
    expect(result.publish).toMatchObject({ provider: 'generic', url: 'https://updates.example.com/dp' });
  });

  it('omits publish entirely by default, so no app-update.yml is written', () => {
    expect(resolveSigning({}).publish).toBeNull();
    expect(resolveSigning({}).notes.join(' ')).toMatch(/no app-update\.yml/);
  });
});
