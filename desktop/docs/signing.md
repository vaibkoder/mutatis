# Signing, notarisation and updates

Everything in this file was checked against primary sources on **11 September 2026**.
Two points are easy to get wrong: Windows signing is available to organisations (not
individuals) in Singapore, and Microsoft's service has been renamed. Read the first two
sections before spending any money.

Nothing here is blocking the app. The build works today; it is simply **unsigned**, and
`npm run check:signing` will tell you so every time you package it.

---

## What "unsigned" costs, concretely

| | macOS | Windows |
|---|---|---|
| First launch | Gatekeeper refuses. The user must open **System Settings › Privacy & Security** and press **Open Anyway**, then confirm again. | SmartScreen shows "Windows protected your PC" with a **More info › Run anyway** behind one click. |
| Repeat launches | Fine after the first override. | Fine. |
| Auto-update | **Impossible.** Squirrel.Mac replaces the bundle only if it can verify the signature, so an unsigned copy can never update itself. | Works, but `verifyUpdateCodeSignature` has no publisher to check against. |
| Cost | £0 | £0 |

For a tool handed to a few colleagues, that friction is survivable. For anything
sent to a client, the macOS dialog reads like a virus warning and it will be the first
thing anyone says about the product.

---

## macOS — the only path is Apple, and it is $99/yr

There is no free alternative and no workaround worth building. The Apple Developer Program
is the sole issuer of Developer ID certificates and the sole route to notarisation.

What you need, in order:

1. **Apple Developer Program membership**, $99/yr. Enrol as an organisation (needs a D-U-N-S
   number, which takes a few days) or as an individual — the certificate's subject differs,
   and users see it.
2. A **Developer ID Application** certificate in the build machine's keychain, or exported
   as a `.p12` and given to CI as `CSC_LINK` + `CSC_KEY_PASSWORD`.
3. **Notarisation credentials.** Use an **App Store Connect API key** — a `.p8` file, a key
   id and an issuer id. Apple's own guidance and electron-builder's both prefer it to an
   Apple ID + app-specific password, because the password is a full-account credential and
   the key is not.

Then:

```bash
export MUTATIS_SIGN=1
export APPLE_API_KEY=/secure/AuthKey_XXXXXXXX.p8
export APPLE_API_KEY_ID=XXXXXXXX
export APPLE_API_ISSUER=00000000-0000-0000-0000-000000000000
export APPLE_TEAM_ID=XXXXXXXXXX
npm run dist:mac
```

`altool` is dead; electron-builder 26 uses `notarytool` through `@electron/notarize`. You
do not configure that — it is picked from the environment variables above.

**`MUTATIS_SIGN=1` is required even when a certificate is present.** See
`scripts/signing.cjs`: a build signed but not notarised is *worse* than an unsigned one,
because Gatekeeper's message for it reads like file corruption. Opt in deliberately.

---

## Windows — Azure Artifact Signing

Eligibility of a Singapore-based publisher for Azure Trusted Signing was checked before
designing around it. Two things to know:

**1. The service was renamed.** It is **Azure Artifact Signing** now (the docs live under
`/azure/artifact-signing/`). Older write-ups all say "Trusted Signing".

**2. Singapore is eligible — as an ORGANISATION, not as an individual.** From Microsoft's
current prerequisites, verbatim in substance:

> Public Trust certificates are available to organizations in the United States, Canada,
> the European Union, the United Kingdom, Australia, New Zealand, Japan, South Korea,
> **Singapore**, Switzerland, Norway, and Israel. **Individual developers must be located
> in the United States or Canada.**

So the individual route is closed and the organisation route is open. If you publish from
Singapore, **enrol a legal entity** (company or firm), not yourself as an individual. That
is not a workaround; it is the eligible path.

What organisation validation asks for:

- A Microsoft Entra tenant and an Azure subscription whose **billing account type is
  Organization** — an individual billing account cannot validate an organisation identity,
  and the legal name on the billing profile becomes the certificate subject. Fix the
  billing profile *before* starting, because a correction means a whole new request.
- The legal entity name, a **website on a domain the entity owns**, a business identifier,
  the registered address, and two monitored email addresses on that domain.
- A named individual who completes a personal ID check (passport or driving licence,
  through AU10TIX) on behalf of the entity.
- **1 to 20 business days** to process. Start this well before you need a release.

Pricing: **Basic ≈ $9.99/month**, Premium $99.99/month for up to 100,000 signatures. Basic
is the right tier for one desktop app.

**A bonus: no renewal date.** Certificates issued on or after 1 March 2026 carry a
460-day maximum validity, so a traditional OV certificate bought from a CA needs renewal
in the release calendar.
Artifact Signing issues **very short-lived certificates per signing operation** and
timestamps the signature, so the signature stays valid after the certificate expires and
there is no renewal date to diarise. If you go this route, the calendar item disappears.

Configuration, once the account exists:

```bash
export MUTATIS_SIGN=1
export AZURE_CODE_SIGNING_ENDPOINT=https://neu.codesigning.azure.net   # match your region
export AZURE_CODE_SIGNING_ACCOUNT=<account name>
export AZURE_CODE_SIGNING_PROFILE=<certificate profile name>
export AZURE_CODE_SIGNING_PUBLISHER="CN=Your Company Pte. Ltd., O=..., C=SG"
export AZURE_TENANT_ID=... AZURE_CLIENT_ID=... AZURE_CLIENT_SECRET=...
npm run dist:win
```

`AZURE_CODE_SIGNING_PUBLISHER` must match the certificate subject **exactly**. Getting it
wrong does not break the build — it breaks auto-update months later, when
`verifyUpdateCodeSignature` rejects a legitimate installer.

**Configuration note:** electron-builder does not select the Windows backend from
`win.sign.type`. In 26.15.3 the schema has `win.azureSignOptions`
and `win.signtoolOptions`, which are mutually exclusive. The config uses the former.

**Build Windows on Windows.** Cross-compiling a Windows target from a macOS runner does not
Authenticode-sign properly. This is why `dist:win` exists as its own script.

---

## Updates

Wired, tested, and **off** — because this build cannot satisfy the preconditions, and says
so rather than failing quietly.

The design decisions are in `src/main/updater.ts` and they are deliberate:

- **No startup check.** The usual `checkForUpdatesAndNotify()` fires an HTTPS request the
  moment the app opens. The egress audit's central claim is that the only thing leaving
  this machine is an AI chat turn the user sent, and a version beacon would make that claim
  false. Both the audit and `check:signing` assert that call is absent.
- **Two presses to install.** Check tells you a version exists. Download is separate.
  Install happens on a restart you chose.
- **`electron-updater` is imported lazily**, so a session that never asks for an update
  never loads it.

To turn updates on you need all three of: a signed macOS build (Squirrel.Mac requirement),
a place to host the feed, and the update URL:

```bash
export MUTATIS_UPDATE_URL=https://updates.example.com/mutatis
```

which configures a `generic` provider. With no such variable, electron-builder writes no
`app-update.yml`, so the shipped app has **no feed URL in it at all** — not a disabled one,
none. The audit checks that too.

Hosting can be anything that serves static files over https: an S3 bucket, a Cloudflare R2
bucket, or a folder on your own web server. electron-builder writes `latest-mac.yml`,
`latest.yml` and the packages into `release/`; upload that directory.

Note the macOS targets change when you sign: `dmg` **and** `zip`. The zip is not optional —
Squirrel.Mac cannot read a dmg, so a dmg-only release has no update path.

---

## The decision, stated plainly

| Option | Cost | What it buys |
|---|---|---|
| Ship unsigned | £0 | Works. Two scary dialogs. No auto-update on macOS. |
| macOS only | $99/yr | Clean install on Mac, auto-update possible. Windows still warns. |
| Both | $99/yr + ~$10/mo | Clean on both, auto-update on both, no renewal calendar. |

If Mutatis is only used by people who know to expect the dialogs, unsigned is defensible.
If you distribute it more widely, sign macOS at minimum — the Gatekeeper dialog is the one that costs you
the conversation.

---

## Verifying whatever you build

```bash
npm run dist && npm run check:signing
```

It passes on an unsigned build and describes it honestly. It fails on a **misdescribed**
one: entitlements broader than declared, a signature without the hardened runtime, a
notarisation claim with no stapled ticket, an update feed baked in by accident, or update
machinery wired to startup.

## Sources

- [Quickstart: Set up Artifact Signing — Microsoft Learn](https://learn.microsoft.com/en-us/azure/artifact-signing/quickstart)
- [Azure Artifact Signing pricing](https://azure.microsoft.com/en-us/pricing/details/artifact-signing/)
- [Code signing options for Windows app developers — Microsoft Learn](https://learn.microsoft.com/en-us/windows/apps/package-and-deploy/code-signing-options)
- [macOS Notarization — electron-builder](https://www.electron.build/docs/features/code-signing/notarization/)
- [electron/notarize](https://github.com/electron/notarize)
