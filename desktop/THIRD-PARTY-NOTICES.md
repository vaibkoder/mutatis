# Third-party notices

Every entry below was read from the installed package's own metadata, not from a badge
or a project's marketing claim.

**Audit status:** 47 direct dependencies (28 shipped, 19 build-only), zero GPL / AGPL / BUSL /
SSPL, zero missing licence declarations. One dual-licensed package (`pizzip`, MIT OR
GPL-3.0) — **we take the MIT option**, which is the user's choice to make under a dual
grant.

## Obligations we actually carry

- **Apache-2.0 §4.** The obligation is the FULL licence text, not
  the short-form header. Every Vercel AI SDK package we installed (`ai`, `@ai-sdk/react`,
  `@ai-sdk/deepseek`) ships a 12-line LICENSE containing only the header, and none ships a
  NOTICE file — checked, not assumed. The complete text is therefore reproduced by us at
  [`licenses/Apache-2.0.txt`](licenses/Apache-2.0.txt), and that file is included in the
  packaged application (`electron-builder.config.cjs` → `files`) so it reaches recipients with the
  binary rather than only living in the repository.
  Apache-2.0 components in this build: `ai`, `@ai-sdk/react`, `@ai-sdk/deepseek`,
  `@ai-sdk/alibaba`, `diff-match-patch`, `fast-diff`, `fuse.js`, `pdfjs-dist` (shipped); `typescript`
  (build only).
  Copyright holders, taken from each package's own files: Vercel, Inc. (2023) for the AI
  SDK; The diff-match-patch Authors (2018) — their LICENSE leaves the appendix placeholder
  unfilled, so this comes from the source header; Jason Chen (2014–2023) for `fast-diff`;
  Kirollos Risk (2017) for `fuse.js`; Mozilla Foundation (2024) for `pdfjs-dist` (PDF.js —
  its LICENSE is the full text; it also ships no NOTICE file). The single-file HTML build
  embeds these libraries without the licence text; add it before redistributing that file.
- **MIT / ISC** require the copyright notice and permission notice to be included.
- **SIL Open Font License 1.1** — the bundled fonts: Liberation Sans and Liberation Serif
  (metric-compatible with Arial and Times New Roman), Carlito (Calibri-compatible) and
  Caladea (Cambria-compatible). They are embedded in the application for the PDF/layout
  path. The licence texts ship with the app in
  [`licenses/`](licenses/): `OFL-LiberationSans.txt`, `OFL-LiberationSerif.txt`,
  `OFL-Carlito.txt` and `OFL-Caladea.txt` (copies of those in the repository root's
  `public/fonts/`). The fonts are not sold on their own
  and are not renamed.

---

## Installed direct dependencies

### Runtime dependencies (shipped in the app)

| Package | Version | Licence |
|---|---|---|
| `@ai-sdk/alibaba` | 2.0.44 | Apache-2.0 |
| `@ai-sdk/deepseek` | 3.0.42 | Apache-2.0 |
| `@ai-sdk/react` | 4.0.100 | Apache-2.0 |
| `@base-ui/react` | 1.8.0 | MIT |
| `ai` | 7.0.97 | Apache-2.0 |
| `pdfjs-dist` | 5.7.284 | Apache-2.0 |
| `chrono-node` | 2.9.1 | MIT |
| `cn` | 0.2.6 | MIT |
| `diff-match-patch` | 1.0.5 | Apache-2.0 |
| `electron-updater` | 6.8.9 | MIT |
| `fast-diff` | 1.3.0 | Apache-2.0 |
| `fast-xml-parser` | 5.8.0 | MIT |
| `@pdfme/pdf-lib` | 6.1.5 | MIT |
| `fontkit` | 2.0.4 | MIT |
| `fuse.js` | 7.4.2 | Apache-2.0 |
| `harfbuzzjs` | 1.2.1 | MIT |
| `lucide-react` | 1.44.0 | ISC |
| `@preact/signals` | 2.9.1 | MIT |
| `pizzip` | 3.2.0 | MIT **OR** GPL-3.0 — we take MIT |
| `preact` | 10.29.2 | MIT |
| `react` | 19.3.0 | MIT |
| `react-dom` | 19.3.0 | MIT |
| `react-markdown` | 10.1.0 | MIT |
| `react-resizable-panels` | 4.12.4 | MIT |
| `remark-gfm` | 4.0.1 | MIT |
| `remark-parse` | 11.0.0 | MIT |
| `unified` | 11.0.5 | MIT |
| `words-to-numbers` | 1.5.1 | MIT |

### Build and development dependencies (not shipped)

| Package | Version | Licence |
|---|---|---|
| `@eslint/js` | 10.0.1 | MIT |
| `@tailwindcss/vite` | 4.3.3 | MIT |
| `@types/diff-match-patch` | 1.0.36 | MIT |
| `@types/node` | 22.18.1 | MIT |
| `@types/office-js` | 1.0.377 | MIT |
| `@types/react` | 19.3.0 | MIT |
| `@types/react-dom` | 19.3.0 | MIT |
| `@vitejs/plugin-react` | 5.1.0 | MIT |
| `electron` | 44.3.0 | MIT |
| `electron-builder` | 26.15.3 | MIT |
| `electron-vite` | 5.0.0 | MIT |
| `eslint` | 10.10.0 | MIT |
| `happy-dom` | 21.0.3 | MIT |
| `tailwindcss` | 4.3.3 | MIT |
| `typescript` | 5.9.3 | Apache-2.0 |
| `typescript-eslint` | 8.70.0 | MIT |
| `vite` | 7.3.6 | MIT |
| `vite-plugin-singlefile` | 2.3.3 | MIT |
| `vitest` | 5.0.0 | MIT |

---

## How this table is kept honest

`node -e` over `package.json` + each installed `node_modules/<name>/package.json`, so the
licence recorded is the one in the package that is actually on disk — not the one on a
registry page, a README badge, or a project's marketing site.

Re-run after any dependency change:

```
node -e "const fs=require('fs');const pkg=JSON.parse(fs.readFileSync('package.json','utf8'));for(const g of ['dependencies','devDependencies']){console.log('##',g);for(const [n] of Object.entries(pkg[g])){const p=JSON.parse(fs.readFileSync('node_modules/'+n+'/package.json','utf8'));console.log('|','`'+n+'`','|',p.version,'|',typeof p.license==='string'?p.license:JSON.stringify(p.license),'|')}}"
```
