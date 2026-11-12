# Mutatis

Legal document comparison and proofreading for **Singapore and England & Wales drafting**.

Its flagship is **Compare**: a Word-vs-Word, Text or PDF redline that categorises every change and never suppresses one (accept-all gives the revised document, reject-all the original).

**Source-available, view only — not open source.** All rights reserved. See [`LICENSE`](LICENSE) before doing anything other than reading.

## Three ways to run it

All three use the same comparison engine (`src/`), which reads both documents and produces the redline on your own computer.

- **Word add-in** — runs inside Word's task pane.
- **Web app** — the same comparison in a browser tab. Your files are read by the page on your own machine and are not uploaded.
- **Desktop app** (`desktop/`) — a stand-alone Windows and macOS program built with Electron. It compares Word and PDF documents side by side and adds an optional **AI chat panel** that can discuss the changes and propose tracked edits.

## What leaves your machine

| Product | Where processing happens | What leaves the machine | To whom |
|---|---|---|---|
| Word add-in | On your computer, in Word | Nothing | Nobody |
| Web app | On your computer, in the browser | Nothing (no document is uploaded) | Nobody |
| Desktop app, AI chat **not used** | On your computer | Nothing. No update check at start-up, no telemetry, no crash reports, no web fonts | Nobody |
| Desktop app, AI chat **used** | Comparison on your computer; each chat message is answered by the provider you chose | The message you typed, the conversation so far, the two document names, and the changed wording (before and after text of each change, clause references, category counts). Unchanged text is not sent. Very large comparisons are cut off at about 120,000 characters of change text, and the app tells the model how much was left out | DeepSeek (a China-based provider, `api.deepseek.com`) or Alibaba Qwen (international endpoint, `dashscope-intl.aliyuncs.com`), whichever you select |
| Desktop app, update check (only if you press the button, and only in a build packaged with an update address) | Your computer | A request for the latest version file | The update server named at build time. Builds have none by default |

Before you use AI chat in the desktop app:

- It is **off until you add your own API key** for DeepSeek or Qwen. You pay the provider directly; Mutatis has no server.
- Once you send a message, the changed wording of the open comparison goes to the provider you selected, and is resent with each later message. **Do not use chat on documents you are not permitted to share with that provider.** Check your engagement terms and the provider's data policy first.
- Your key is encrypted with the operating-system keychain (macOS Keychain or Windows DPAPI) and stored in the app's user-data folder in encrypted form. The app window never sees it, and the app refuses to save it if encryption is unavailable.
- Conversations are saved on your computer only, in a database file (`threads.db`) in the app's user-data folder. That file is **not encrypted**: it holds your messages, the assistant's replies and any document wording they quote, so anyone with access to your user account can read it. Delete a conversation with the bin control beside it. Uninstalling the app does not remove the user-data folder.
- The provider sees what you send under its own terms. Mutatis cannot see or control how long the provider keeps it.

## Build and test

Requires Node 22 or later.

**Word add-in and web app** (repository root):

```bash
npm install
npm run dev        # Vite dev server on https://localhost:3000 (sideload manifest/manifest.dev.xml in Word)
npm run build
npm test
npm run lint
npm run typecheck
```

**Desktop app** — install the root first (the desktop app reads the engine from `src/`), then from `desktop/`:

```bash
npm install
npm run dev        # run the app in development
npm test
npm run build      # type-check and build
npm run dist:mac   # package a macOS app (run on a Mac)
npm run dist:win   # package a Windows installer (run on Windows)
```

Packaged desktop builds are **unsigned** unless you opt in. Signing and notarisation are driven by environment variables (`MUTATIS_SIGN=1` plus your Apple or Azure credentials), and the update feed by `MUTATIS_UPDATE_URL`. See [`desktop/docs/signing.md`](desktop/docs/signing.md). Unsigned builds work, but macOS and Windows show a warning on first launch.

## Where things live

| Path | What it is |
|---|---|
| `src/` | The engine and the Word add-in / web app. |
| `desktop/` | The Electron desktop app. Its third-party notices are in `desktop/THIRD-PARTY-NOTICES.md`. |
| `spec/` | The product specification. Start at `spec/README.md`, then `spec/ARCHITECTURE.md` and `spec/CONVENTIONS.md` (the single source of truth for SG/UK legal conventions). |
| `docs/adr/` | Architecture Decision Records — *why* each major choice was made, and what was rejected. |
| `public/fonts/` | Bundled open-licence fonts (SIL OFL 1.1; licence texts alongside). |

## Licence

Proprietary, source-available, **view only**. All rights reserved — see [`LICENSE`](LICENSE). It covers the whole repository, including `desktop/`. Third-party dependencies remain under their own licences.

No contributions are solicited; see clause 4 of the licence before opening a pull request.

Mutatis is not legal advice and is not a substitute for professional legal judgment.
