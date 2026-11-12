# Standalone browser try-out (experimental — not a product target)

Repackages the existing Office-free web app (`src/web/`, ADR-0022) so you can poke
at it in a plain browser without Word. **Not an end product**; the Word add-in is
unaffected. Still fully client-side — no backend, nothing leaves the machine.

Build configs live in `vite.standalone.config.ts`; outputs are git-ignored.

## What works where

| Feature | Single file (double-click) | Served folder (http) |
|---|:--:|:--:|
| Proof checkers (Defined Terms, Cross-Ref, Party-Role, Factual) | ✅ | ✅ |
| Word-vs-Word redline `.docx` download | ✅ | ✅ |
| Text-vs-text redline `.docx` download | ✅ | ✅ |
| Rendered redline preview + **vector-PDF export** | ✅ | ✅ |
| PDF-vs-PDF compare | ❌ | ✅ |

The rendered redline (preview + the **PDF download**) normally runs in a Web Worker,
which `file://` blocks. The single-file build instead runs that engine on the main
thread (the wasm shaper ships inlined as a data URL), so the **PDF redline works
from a plain double-clicked file — no server**. A big render briefly freezes the UI;
fine for trying it.

The one feature the single file can't do is **PDF-vs-PDF compare**: it relies on
PDF.js, which needs its worker + vendored font/cmap assets fetched over `http://`.
For that one, use the served folder.

## Option A — single self-contained file (quick double-click)

```bash
npm run build:single
```

Copy **`dist-single/index.html`** (~3.6 MB, everything inlined — JS, CSS and the
shaper wasm) to the Windows box, double-click it → opens in Chrome/Edge. Rename it
to anything you like, e.g. `mutatis.html`. That one file is all you need,
including the rendered redline **PDF download**. No server, no `.bat`, no install.

## Option B — served folder (everything works)

```bash
npm run build:web
```

Copy the whole **`dist-web/`** folder over, then from inside it serve at the root
and open the printed URL:

```bash
# Python (ships with Windows' python.org installer)
python -m http.server 8000 --directory dist-web
#   → http://localhost:8000/

# or Node, if you have it
npx serve dist-web
```

Serve from the folder **root** — the app fetches `/assets/*` and `/pdf-assets/*`
as same-origin paths.
