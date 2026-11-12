# 17 — Compare at Scale: one-to-many & many-to-many (standalone/web target)

> **Status:** ✅ v1 — two features in one spec. **scope row 16** = one-to-many (1:N); **scope row 17** = many-to-many (M:N). Governed by **ADR-0037**. Extends feature **08-compare** (Word-vs-Word) and inherits **ADR-0008** (categorise, never suppress) unchanged.
>
> *(Numbering: the project scope rows and `spec/features/NN-*.md` file numbers are different sequences. This one file is spec **17** and covers scope rows **16 and 17**.)*
>
> Heading scheme: `#` doc title, `##` numbered top sections, `###` sub-sections.

---

## 1. Purpose & user value

Feature 08 compares one document against one document, one comparison at a time. Two routine SG/UK transactional workflows do not fit that shape.

**One-to-many (1:N).** A firm holds a house precedent or an agreed template and needs to know how each of *N* returned or executed versions departs from it — a standard NDA against the forty variants a client has actually signed; an LMA-based facility agreement against each borrower's negotiated version; a template SPA against every version returned during a bidding round. Today this is *N* manual Compare runs, each re-uploading and re-preprocessing the same template.

**Many-to-many (M:N).** Two bundles — a signing set against a closing set, a data-room folder against the version counsel is told is current. The documents correspond, but the correspondence is not recoverable from filenames alone: `Project_Falcon_SPA_v7_exec.docx` is the counterpart of `SPA (clean) 2026-08-30.docx`, while the disclosure letter, the side letter and the SPA all share the same parties and most of the same defined terms.

Both are done today either by hand or by uploading a whole deal bundle to a cloud comparison service — exactly the confidentiality problem this product exists to remove.

**User value in one line:** "Redline a whole bundle against a template — or one bundle against another — without a single document leaving the machine."

---

## 2. Scope

### 2.1 In scope ✅ v1

| # | Capability | Feature |
|---|---|---|
| 1 | **Batch runner** — ingest *N* variants one file at a time, preprocess the template **once**, run each pair through the shipped Compare pipeline in a bounded worker pool, stream results. | 16 + 17 |
| 2 | ***N* separate redlines** — for each pair, the **editable tracked-changes `.docx`**: the primary artefact a single Compare run produces today, unchanged. The read-only RenderCore snapshot is **not** produced per pair in v1 (§2.2, §10). | 16 + 17 |
| 3 | **A compact batch index** — one row per pair (identity, status, change counts by classification bucket, artefact handles, timing) plus batch totals and disclosures. Flat, filter-independent, built the way `export/report-model.ts` is built. | 16 + 17 |
| 4 | **Batch progress + cancellation** — two-level progress (batch `done/total`, per-pair %), cancel at any point; completed pairs are **kept and labelled**, not discarded. | 16 + 17 |
| 5 | **Batch input caps** — named constants, typed errors, guards that fail before any partial work. | 16 + 17 |
| 6 | **Document pairing** — content Jaccard over existing paragraph hashes ▸ version-token-stripped filename similarity ▸ defined-term/party-cast overlap as a **tie-breaker only**; Hungarian assignment padded with no-match dummies. | 17 |
| 7 | **A confirmable pairing grid** — every proposed pair with its score and top rivals, in three confidence bands, **confirmed by the user before anything runs**. Bands change presentation only; nothing is hidden. | 17 |
| 8 | **Drill-down** — open one pair's full change list by **re-running that single pair**. | 16 + 17 |

**Both features live in the standalone/web target only** (`src/web/**`, `src/host-web/**`, plus host-free code under `src/features/compare/**`). **The Word add-in is untouched** — no new tab, no new panel, no behaviour change, no bundle growth. ADR-0003's two-documents-in-memory `getFileAsync` ceiling is the reason, and it is structural.

**Sources are `.docx` only in v1** (both sides, both features). The pairing layer is built **format-agnostically** — it consumes a `PairingSketch` that any source kind can produce — so a future PDF sketch builder drops in without redesign. That is the "designed for cross-format pairing" seam, and it is the whole of it.

### 2.2 Cut / out ❌ and deferred 🔜

- **Aggregate "deviation matrix" / portfolio dashboard** ❌ (in this product) — *N* separate redlines plus the compact index is the deliverable. The matrix is a different product with its own UI, persistence and performance envelope; deferred to a **separate desktop product** by owner decision (2026-09-10, ADR-0037).
- **Batch compare inside the Word add-in** ❌ — ADR-0003, ADR-0037.
- **Silent / automatic pairing** ❌ — a mis-pairing produces a plausible, fully-populated, entirely worthless whole-document rewrite that a reader has no cue to distrust. Pairing is always proposed and confirmed. This is ADR-0008 one level up.
- **A confidence threshold that removes candidates from the pairing grid** ❌ — confidence is a **badge**, never a hidden filter, exactly as ADR-0017 §Invariant analysis already requires for alignment confidence.
- **AI / embedding / model-based similarity, anywhere** ❌ — ADR-0001 (network) and ADR-0008 (opaque judgement). Pairing signals are lexical and structural, and every score shown to the user is explainable.
- **Cross-format (Word-vs-PDF) comparison** ❌ here — ADR-0017 gated Phase E on a measured alignment bar that was **never measured** (deferred). Reviving it inside a batch feature would ship an ungated capability by the back door.
- **Network, Electron, chat** ❌ — ADR-0001, ADR-0022.
- 🔜 **A rendered RenderCore snapshot per pair in a batch.** A single Compare run produces the editable `.docx` *and* a read-only rendered snapshot (ADR-0020, ADR-0024). The batch produces only the first. The snapshot's value is in *reading one redline* — which is what the drill-down is for — while rendering it costs a full layout pass per pair on top of the comparison, and holding a hundred of them reopens the memory ceiling §4.3 exists to close. So the batch spends its budget on the artefact a batch is actually for, and the snapshot stays where it earns its cost: one pair at a time. Revisit if the paginated renderer becomes cheap enough that N of them are free.
- 🔜 **PDF-vs-PDF batching** — running the in-scope PDF mode across a bundle. Not built in v1; the format-agnostic `PairingSketch` seam is what it will attach to.
- 🔜 **Cross-bundle move detection** (a clause relocated from the SPA into the side letter). Out — move detection stays within a pair, as in feature 08.

### 2.3 Invariant check

- **Confidentiality (ADR-0001/0003).** Everything is client-side. No new dependency, no new asset fetch, no telemetry. All new code sits under paths the `no-network-egress` lint already polices, and none of it is on the `no-host-imports` allow-list — so it is lint-proven both network-free and host-free.
- **Jurisdiction (ADR-0002).** The classifier and `CompareConfig` are reused verbatim; the batch layer adds no locale logic. The one place jurisdiction *does* surface is the filename version-token vocabulary (§6.3), which is SG/UK drafting vocabulary (*engrossment*, *conformed*, *initialled*) and day-first dates — **never** `MM/DD/YYYY`.
- **Scope (ADR-0005).** ADR-0037 + scope rows 16/17 are the gate.
- **Categorise, never suppress (ADR-0008).** Extended to pairing, not weakened: every pairing candidate stays reachable, every completed comparison carries its full change set, and a cancelled batch keeps what it finished.

---

## 3. Dependencies & reuse

This feature adds **no diff logic and no classification logic**. It adds a runner, a pool, a pairing layer, and two panels around the shipped engine.

**Reused unchanged:**

| What | Where | How it is used |
|---|---|---|
| `runComparePipeline(input)` | `compare/pipeline.ts:107` | The batch entry point. Already takes **pre-built `ParaRecord[]`** and already supports `onProgress` / `abortSignal`. **Never call `runPipeline` (`pipeline.ts:55`) per variant** — it re-preprocesses both sides every time, which is precisely the waste this feature exists to remove. |
| `preprocessSide(...)` | `compare/preprocess/index.ts:42` | Called **once** for the template side in 1:N; once per document in M:N. |
| `fnv1a` + `ParaRecord.hash` | `compare/preprocess/hash.ts:19`, `model.ts` | Paragraph hashes are computed during preprocessing regardless — they are free input for the pairing sketch. (Note: `fnv1a` is **32-bit / 8 hex**, despite the `ParaRecord.hash` JSDoc saying 64-bit. Fine for Jaccard estimation; **do not** use a single paragraph hash as a document identity key.) |
| The whole classifier, diff core, `CompareConfig`, `model.ts` | `compare/classify/**`, `compare/diff/**` | Untouched. |
| `output/author.ts`, `output/writeout.ts` | `compare/output/**` | Per-pair redline materialisation, exactly as today. |
| RenderCore snapshot path | `compare/render/**`, `shared/render-core/**` | **Not used by the batch in v1** (§2.2). Listed because the drill-down and the single-pair surface still reach it, and because a future batched snapshot attaches here. |
| Worker protocol/client/handler shape | `shared/render-core/worker/{protocol,client,handler}.ts` | The **pattern** to copy — including the `WorkerLike` structural interface that makes the worker path unit-testable, the `VITE_DP_INPROCESS` short-circuit, and the `try/catch` + in-process fallback in `compare/render/index.ts:103-134`. |
| `buildCompareReport` | `compare/export/report-model.ts:112` | The **precedent** for the batch index: pure, date-injected, takes **no filter flags**, so `rows.length` always equals the true total. |
| `loadDocSnapshotFromOoxml` | `host-web/ooxml-loader.ts:54` | Available for host-free sketch enrichment; note it produces the checker `ParaNode` model, **not** Compare's `ParaRecord`. |

**New:** a batch runner, a bounded worker pool (the repo has **none** today — RenderCore runs one shared worker with job ids), a pairing module, a batch `limits.ts`, and two web panels.

---

## 4. Architecture

### 4.1 Prerequisite — the driver must become instantiable

`compare/driver.ts` holds its state in **module-level Preact signals** plus eight **module-level mutables** (`_originalSource`, `_revisedSource`, `_abortSignal`, `_heldRedline`, `_textAbort`, `_pdfOriginalFile`, `_pdfRevisedFile`, `_pdfAbort`) and guards re-entry with `if (compareStatus.value === 'running') return;`. That is one run at a time, by construction, and it blocks all batch work.

**The refactor is contained and behaviour-neutral.** The factory precedent already exists in-repo — `makeStubDriverState()` at `compare/ui-types.ts:267` returns a fresh object of signals satisfying the same `CompareDriverState` interface — and **the UI layer is already instance-agnostic**: every panel receives its driver as a prop, and only two files consume the module singletons (`taskpane/shell/Shell.tsx:25`, `web/WebCompare.tsx:23`).

Therefore: wrap the state and the module mutables in `createCompareDriver(): CompareDriverState` (mutables become closure locals), and **keep `export const compareDriverState = createCompareDriver()`** so both consumers and the five tests that import the driver directly keep working unchanged. Same for the text and PDF blocks. A test must assert that **two instances do not share state** — that is the whole point of the change, and nothing else in the suite would catch a regression.

### 4.2 The worker pool

Copy the RenderCore transport shape into `compare/batch/worker/{protocol,client,handler}.ts` + `src/workers/compare.worker.ts`.

- **Protocol** — two discriminated unions carrying a job `id`: `ClientToWorker = CompareJob | Cancel`, `WorkerToClient = Progress | Result | Error | Cancelled`. `ParaRecord` is plain data and structured-clone-safe; nothing is transferable, so nothing is transferred.
- **Cancellation is finer-grained here than in RenderCore.** `runComparePipeline` *awaits* between word-diff chunks (`WORD_DIFF_CHUNK = 25`), so a `{ aborted: boolean }` token flipped by an incoming `cancel` message is genuinely checked mid-run — the render handler can only check at stage boundaries. Rejection uses the same contract the pipeline already throws: `DOMException('…cancelled', 'AbortError')`.
- **Pool size** — `BATCH_POOL_SIZE = clamp((navigator.hardwareConcurrency ?? 2) - 1, 1, BATCH_POOL_MAX)`, `BATCH_POOL_MAX = 4`. One core is left for the UI. Overridable via `compare.batch.poolSize` (`0` = auto).
- **Fallback ladder, mirroring `compare/render/index.ts:103-134`** — honour `VITE_DP_INPROCESS` (the single-file standalone build tree-shakes workers out), `try/catch` the `new Worker(new URL(...), { type: 'module' })` construction (a `file://` page blocks module workers), and fall back to a **sequential in-process client of the same shape**. A batch must run, slower, wherever a single compare runs.
- **Keep the `WorkerLike` structural interface** so tests inject a fake worker that pipes straight into the handler in-process. This is why the render worker is testable without jsdom `Worker` support, and it is non-negotiable for a pool.

### 4.3 Memory policy — the thing that actually breaks first

Each `PipelineResult` holds `originalParas` **and** `revisedParas`, and every `ParaRecord` carries the paragraph's `rawXml`. A result is therefore roughly document-sized on both sides. **Retaining *N* results is an out-of-memory crash, not a slowdown.**

The policy, in order:

1. **Ingest one file at a time** (`file.arrayBuffer()`), parse, preprocess, release the buffer. Never hold *N* `ArrayBuffer`s.
2. **The template side is preprocessed exactly once** and its `ParaRecord[]` is shared, read-only, across all *N* jobs. (*One* document reused across *N* jobs. The many-to-many analogue would be *M+N* documents retained at once, which is the opposite of this policy rather than an application of it — see §6.7.) The pipeline stages read the input arrays and return them by reference; they do not mutate them. A test must pin this — deep-equal the template `ParaRecord[]` before and after a batch.
3. **On each pair completing: materialise the artefacts, then release the result.** A `Blob` is browser-managed storage the engine may spill to disk; a `ParaRecord[]` is JS heap. So the redline `.docx` is materialised **eagerly into a `Blob`** as each pair finishes, and the `PipelineResult` — ops, both `ParaRecord[]`, `authoredBodyXml` — is dropped immediately.
4. **Retain only the compact row** (§5.2). Peak retained JS heap is therefore *one* pair's working set times the pool size, not *N*.
5. **Drill-down re-runs the single pair.** Opening a variant's full change list re-runs it — the ops and paras are exactly what step 3 dropped. Bounded, fast, and correct.

   *Amended during 17b:* this holds on the many-to-many path too. A pairing sketch carries hashes, not records (§6.2, §6.7), so opening a pair re-reads and re-preprocesses its two documents exactly as this step describes — and, exactly as §6.7 requires, it stops at the pipeline and never packages a redline.
6. **A cumulative artefact ceiling.** Past `MAX_BATCH_ARTEFACT_BYTES`, stop materialising eagerly, mark subsequent rows `artefactOnDemand`, and **say so in the batch disclosures**. Their artefacts are produced by re-running that pair on download. Degrading silently is not an option.

### 4.4 Input caps — `compare/batch/limits.ts`

Compare's only caps today are `io/pdf/limits.ts` (PDF, per run) and `MAX_TEXT_PASTE_CHARS`. **The `.docx` path has no size cap at all** — survivable at one pair, not at *N*. Following the established convention (named constant ▸ typed error ▸ guard that fails **before** any partial work):

| Constant | Default | Applies to |
|---|---|---|
| `MAX_BATCH_VARIANTS` | `100` | 1:N variant count |
| `MAX_BUNDLE_DOCUMENTS` | `50` | M:N documents **per side** |
| `MAX_DOCX_BYTES` | `50 MiB` | any single `.docx`, either feature |
| `MAX_BATCH_TOTAL_BYTES` | `1 GiB` | sum across a batch |
| `MAX_BATCH_ARTEFACT_BYTES` | `512 MiB` | cumulative retained artefacts (§4.3 step 6) |

`BatchCapKind = 'variants' | 'documents' | 'docxBytes' | 'totalBytes' | 'artefactBytes'`; `BatchInputTooLargeError` names the cap, the limit and the offending document. **A batch that would exceed a cap is refused, never silently truncated** — truncating a batch is suppression by another name.

---

## 5. Feature 16 — one-to-many (1:N)

### 5.1 Flow

```
pick template (.docx)  ─┐
pick N variants (.docx) ─┴─▶ cap check (§4.4, fails before any work)
        │
        ├─▶ preprocess template ONCE  ──────────────────┐  (shared, read-only)
        │                                               │
        └─▶ for each variant, streamed one at a time:   │
              read → preprocess → enqueue job ──────────┴─▶ [worker pool, §4.2]
                                                              │
              ◀── progress (batch done/total + per-pair %) ────┤
              ◀── result ──▶ materialise Blobs ──▶ compact row ┘  (result released)
        │
        └─▶ batch index (§5.2) · per-row download · "download all" (streamed ZIP) · drill-down (re-run one pair)
```

The template is always the **Original** side and each variant the **Revised** side, so every redline reads "what this version did to our template". This is fixed, not a toggle — a batch of redlines that disagree about direction is unreadable.

### 5.2 Data model

Lives in `compare/batch/model.ts`. Plain data, structured-clone-safe, no filter flags anywhere.

```ts
type BatchRowStatus =
  | 'queued' | 'running' | 'done'
  | 'failed'      // this pair errored; the batch continues
  | 'cancelled'   // never ran, or ran partially, because the batch was cancelled
  | 'unrunnable'; // pair identified, comparison out of scope (e.g. cross-format)

interface BatchRow {
  id: string;                     // stable within the batch
  label: string;                  // variant file name (or "A ▸ B" for M:N)
  status: BatchRowStatus;
  reason?: string;                // required when status is 'failed' | 'unrunnable'
  changeCount: number;            // === summary.total; never a filtered count
  byClassification: Readonly<Record<ChangeClass, number>>;
  quarantinedOriginal: number;
  quarantinedRevised: number;
  artefacts: readonly ArtefactHandle[];  // v1: at most one, kind 'redlineDocx'
  artefactOnDemand: boolean;      // §4.3 step 6 — nothing retained; re-run on download
  durationMs: number;
}

interface BatchIndex {
  kind: 'compare-batch';
  mode: 'one-to-many' | 'many-to-many';
  templateLabel?: string;         // 1:N only
  generatedAt: string;            // INJECTED — never Date.now() inside the builder
  rows: readonly BatchRow[];
  totals: { pairs: number; done: number; failed: number; cancelled: number; unrunnable: number; changes: number };
  flatten: { template: FlattenPolicy; variant: FlattenPolicy };  // §8 case 7 — per batch, stated here
  disclosures: readonly string[];
}
```

`buildBatchIndex(input): BatchIndex` is **pure and takes no filter flags**, mirroring `buildCompareReport`. `rows.length` always equals the number of pairs the batch was asked to run — a failed, cancelled or unrunnable pair is a **row with a reason**, never an absence.

### 5.3 Output artefacts

One artefact kind per pair in v1: the **editable tracked-changes `.docx`**, exactly as a single Compare run produces it. The RenderCore rendered snapshot is not batched (§2.2); the drill-down and the single-pair surface are where it is still produced. Batch adds two exits and no new artefact kind:

- **Per-row download** of that pair's redline.
- **"Download all"** — a ZIP assembled by streaming each `Blob` in, one at a time, so peak memory stays at one artefact. Filenames are derived from the variant labels, de-duplicated with a numeric suffix, and sanitised.

**A row past the artefact ceiling is still downloadable.** `artefactOnDemand` means *not retained*, never *not available*: the redline is produced by re-running that one pair, both for a single download and — lazily, one at a time — as the archive is written. Re-running them all up front would rebuild the exact memory peak that made them on-demand. A ceiling that cost the user a hundred redlines would be a silent loss of results, which is the batch-scale shape of the thing ADR-0008 forbids.

### 5.4 UI

A new `View` member and one conditional render in `src/web/WebApp.tsx` (the web shell is a small `useState` over the views, not a router — two when this was written, four once both batch surfaces landed). **The batch panel must be handed its driver explicitly** — `ComparePanel` and friends default to **no-op stub drivers**, so a panel that is not injected silently does nothing and no unit test catches it (`web/WebCompare.tsx:15-22` records exactly this). Mirror that injection.

The panel shows: template slot, variant drop-zone with the running count against `MAX_BATCH_VARIANTS`, a two-level progress bar, a cancel button, and the row list (label · status · change count · classification chips · download · open). Selecting a row opens the existing single-pair change-list UI, driven by **a pipeline re-run of that one pair** — `runComparePipeline` only, never `materialise`. Going through the single-pair driver would materialise the redline, and on web `materialise` writes back through the host, which means triggering a download: an absurd side effect for "show me what changed". The instantiable driver (§4.1) is still the prerequisite that makes a second concurrent comparison possible at all; the drill-down simply does not need the driver's write-back half. The change list is rendered **unfiltered** — a drill-down arriving pre-filtered would show less than the row's own count claims.

### 5.5 Progress & cancellation

Batch progress is `done/total` plus the in-flight pairs' percentages, aggregated **by the runner** — `onProgress` is a plain per-call `(pct: number) => void` with no job id, so the runner owns the aggregation.

On cancel: post `cancel` to every busy worker, clear the queue, mark queued rows `cancelled`. **Completed pairs and their artefacts are kept and labelled** ("batch cancelled after 60 of 100"). Feature 08 §9.12's "partial state is discarded cleanly" applies to the *in-flight pair*, not to the batch — discarding sixty finished redlines because the user stopped the sixty-first would be user-hostile and is not what that clause means.

---

## 6. Feature 17 — many-to-many (M:N)

Everything in §4 and §5 applies. This section adds only the pairing layer.

### 6.1 Flow

```
bundle A (M docs) + bundle B (N docs)
   → cap check → ingest + preprocess each doc ONCE → build a PairingSketch per doc
   → score every A×B pair (§6.2–6.3) → Hungarian assignment with no-match padding (§6.4)
   → RENDER THE PAIRING GRID AND STOP (§6.6)
   → user confirms / re-pairs / marks unmatched
   → only then: run the confirmed pairs through the §5 batch runner
```

**Nothing is compared before the user confirms the pairing.** That is the feature's load-bearing rule.

### 6.2 The pairing sketch

```ts
interface PairingSketch {
  docId: string;
  side: 'a' | 'b';
  sourceKind: 'docx' | 'pdf' | 'other';   // the format-agnostic seam (§2.1)
  fileName: string;
  byteLength: number;
  paraHashes: ReadonlySet<string>;        // fnv1a of normText, boilerplate filtered
  definedTerms: ReadonlySet<string>;
  partyNames: ReadonlySet<string>;
  quarantinedCount: number;
  unreadable?: string;                    // §8 case 2 — shown, never paired
}
```

**Amended during 17b — three fields this block first omitted, and one it must
never regain.** `sourceKind` is the union today, not `'docx'` alone: a `.pdf` in
a bundle is sketched on its filename so it can be paired and shown with its
reason (§6.6, acceptance M8), and `'other'` exists so an unexpected extension is
a shown row rather than an exception. `unreadable` carries §8 case 2 into the
pairing stage. And the records themselves are **not** here: they were, so that a
confirmed run would not preprocess a document twice, and at the §4.4 cap that
holds a hundred document-sized `ParaRecord[]` live from the moment a pairing is
proposed — the shape §4.3 opens by naming as an out-of-memory crash. See §6.7.

`paraHashes` comes free from preprocessing. **Paragraphs below `MIN_SKETCH_TOKENS` (4) content tokens are excluded** — empty paragraphs, single-word headings and numbering stubs are identical across every document in a bundle and would inflate every Jaccard score toward each other. Four tokens matches the guard ADR-0032 already uses for absorbed paragraphs.

### 6.3 Signals

**1. Content — dominant.** Jaccard over `paraHashes`: `|A ∩ B| / |A ∪ B|`. At M, N ≤ 50 with a few thousand paragraphs each, the exact set intersection over all M×N pairs is trivially cheap, so **compute it exactly**; a MinHash sketch is the bounded-memory optimisation if a bundle ever exceeds `MINHASH_THRESHOLD` paragraphs, and it estimates the same quantity.

**2. Filename — secondary.** Normalise, then score with the Sørensen–Dice coefficient over character trigrams (reusing the Dice implementation ADR-0026 already added). Normalisation lowercases, drops the extension, collapses separators, and strips **version tokens**:

> `v1`…`v99`, `(2)`, `copy`, `draft`, `final`, `clean`, `blackline`, `redline`, `exec`, `execution`, `signed`, `initialled`, `conformed`, `engrossed`, `engrossment`, and dates in **day-first** forms (`2026-08-30`, `30-08-2026`, `30.08.26`, `30 Aug 2026`, `30 August 2026`).

The vocabulary is SG/UK drafting vocabulary and the date forms are day-first per `CONVENTIONS.md` → Dates. **`MM/DD/YYYY` is not in the list and must not be added** — it is a `no-us-literals` lint violation and a jurisdiction bug.

**3. Defined-term and party-cast overlap — tie-breaker only, and only that.** Inside one deal bundle the SPA, the disclosure letter and the side letter share the same cast and most defined terms, so this signal carries almost no discriminating power. It is therefore **not in the weighted sum at all**. It is applied only to order rivals whose combined scores differ by less than `PAIRING_TIE_EPSILON`. Putting it in the sum would let a near-useless signal move a pair across a band boundary.

```
score(a, b) = W_CONTENT * jaccard(a.paraHashes, b.paraHashes)
            + W_FILENAME * dice(normalise(a.fileName), normalise(b.fileName))

W_CONTENT = 0.80   W_FILENAME = 0.20   PAIRING_TIE_EPSILON = 0.02
```

**Amended during 17b — when the content signal is UNAVAILABLE.** A document with
no paragraph hashes (a `.pdf`, whose text is deliberately not extracted for
pairing, or a `.docx` with nothing long enough to be sketch-worthy) yields a
Jaccard of 0. Fed into the sum above, that scores it as *positively dissimilar*,
caps the pair at 0.20 and pushes it under the no-match price — so a Word and a
PDF of the very same agreement could never be paired at all, and acceptance M8
would be unreachable rather than merely unmet.

So a pair with no content evidence is scored on its filename alone, **into a
ceiling**:

```
score(a, b) = NO_CONTENT_CEILING * dice(normalise(a.fileName), normalise(b.fileName))
            when either side has no paragraph hashes

NO_CONTENT_CEILING = min(1, PAIRING_NO_MATCH_SCORE + PAIRING_NO_CONTENT_HEADROOM)   (0.25 + 0.05)
```

The `min` is not decoration: `noMatchScore` is a unit value, so a profile
setting it to 1 would put the ceiling at 1.05 and the grid would report a pair
at 105%. A score is a proportion, and nothing downstream is entitled to see one
above its own maximum because a threshold was pushed to the edge of its range.

The ceiling is the correction to a first attempt that renormalised outright. A
review found what that does on exactly the mixed bundle it was written for: a
`.pdf` whose name matched perfectly scored 1.000, won the *global* assignment,
and took a side-A document away from its genuine `.docx` counterpart — which was
then displayed as "no counterpart in the other bundle". Nothing was hidden, and
the proposal was still wrong.

**The ceiling governs COMPARISON, not formation.** A second review found the
first version of this paragraph claiming more than the ceiling can deliver: that
scaling into a narrow band above the no-match price is what lets a cross-format
pair form "when it is the only candidate". It is not, and cannot be. A score is
one number on one scale, and the two duties asked of it pull in opposite
directions — pitched high enough to beat a dummy it displaces real pairs, pitched
low enough not to displace them it beats no dummy either. Under the ceiling
alone, §1's own motivating example (`Project_Falcon_SPA_v7_exec.docx` against
`SPA (clean) 2026-08-30.pdf`, Dice 0.261) stopped pairing altogether and both
documents were labelled "no counterpart in the other bundle" — untrue, and M8
unreachable again by the opposite route.

So the ceiling keeps only the duty a score can discharge: **ordering**. It is
what a filename-only candidate is worth *relative to other candidates* — where it
sits in a row's rival list, and how `bandFor` measures the margin between the
proposed pair and its nearest rival. An uncapped PDF rival at 1.000 would
otherwise suppress the auto-pair of a perfectly good `.docx` pair merely by being
in the same bundle. Whether such a pair FORMS is answered structurally instead,
by the two-phase assignment in §6.4.

`contentAvailable` is carried on the score rather than inferred, so the grid can
say which of the two formulas produced a number.

### 6.4 Assignment

Cost matrix `cost(a, b) = 1 − score(a, b)`, padded with dummy rows and columns priced at `NO_MATCH_COST`, solved with the **Hungarian algorithm**.

**Corrected during 17b — the dummy price is HALF what this section first said.**
It was written as `NO_MATCH_COST = 1 − PAIRING_NO_MATCH_SCORE`, which is wrong,
and wrong in the direction that does not announce itself. Leaving a real A and a
real B both unmatched costs `2 · NO_MATCH_COST` — each pays once — while pairing
them costs `1 − score`. The solver therefore prefers the pair exactly when
`1 − score < 2 · NO_MATCH_COST`. For that to mean `score > PAIRING_NO_MATCH_SCORE`,
which is what the constant is *defined* to mean, the dummy price must be:

```
NO_MATCH_COST = (1 − PAIRING_NO_MATCH_SCORE) / 2      (PAIRING_NO_MATCH_SCORE = 0.25)
```

The original figure would have doubled the effective threshold and paired
documents sharing almost nothing, which looks exactly like a working pairing
until someone reads one of the redlines. The code has always had the derived
value; this sentence is what was wrong. A document whose best real partner scores below the no-match price is assigned to a dummy and falls out as **added** or **removed** — which is the correct answer for a bundle where one side gained a side letter, and is exactly what a greedy walk gets wrong.

Greedy nearest-neighbour is rejected on the record (ADR-0037): it is order-dependent across a bundle of near-identical documents, so the same two folders could pair differently depending on the order the files were picked. At M, N ≤ 50 the O(n³) global solve costs nothing.

**Amended during 17b — the solve runs in TWO phases.** The single global solve
above cannot both refuse to let filename-only evidence displace a
content-corroborated pair and still let a cross-format pair form; see §6.3 for
why no single score can. The two questions are therefore separated, and the
answer to *formation* is the phase a candidate is eligible for:

1. **Phase 1 — pairs the content corroborates.** The matrix above, but a cell
   whose score has `contentAvailable === false` is priced *strictly above* the
   cost of leaving both its documents unmatched, so a cost-minimising solver can
   never take it. Filename-only evidence cannot take a document away from a
   pairing content supports, because it is not in this phase at all.

2. **Phase 2 — what is left, on the filename alone.** Over phase 1's leftovers
   only, and only over cells where the content signal was *unavailable* rather
   than merely weak: a content-available pair that lost phase 1 lost on evidence,
   and re-running it here on its name would be exactly the force-matching the
   dummy exists to prevent. The dummy price is `(1 − PAIRING_NO_CONTENT_MATCH_SCORE) / 2`
   by the same derivation as above, with `PAIRING_NO_CONTENT_MATCH_SCORE = 0.20`.

Being a *phase* rather than a lower score is what makes both halves true at once:
filename-only evidence is enough to beat **nothing** — so a cross-format pair
forms and is surfaced with its reason (M8) whatever the rest of the bundle looks
like — and never enough to beat **something**, because by the time phase 2 runs
there is nothing left for it to beat.

The residual trade is documented rather than hidden: a pair phase 1 *rejected* on
content (score under the no-match price) can be picked up in phase 2 by a
better-named no-content candidate. That is the intended reading — phase 1 already
said those two documents are not counterparts — and it is covered by a test.

Pairs are emitted in `docId` order, not in phase order: which phase placed a pair
is an implementation detail, and M4 requires the output to be a function of the
two input sets.

### 6.5 Confidence bands

| Band | Condition | Presentation |
|---|---|---|
| **auto-pair** | `score ≥ AUTO_PAIR_SCORE (0.85)` **and** margin over the best rival `≥ AUTO_PAIR_MARGIN (0.15)` | proposed, pre-ticked |
| **propose-with-rivals** | assigned to a real partner but not meeting auto-pair | proposed, **top two rivals shown with their scores** |
| **unmatched** | assigned to a dummy by both phases (§6.4) | shown as added / removed, with its best candidate still listed |

**Amended during 17b — the unmatched condition is not a score threshold.** This
table read `score < 0.25` (`PAIRING_NO_MATCH_SCORE`), which was true of a single
solve and is not true of the two-phase assignment. A phase-2 pair forms on its
*filename* clearing `PAIRING_NO_CONTENT_MATCH_SCORE` while its *score* is
`NO_CONTENT_CEILING × filename` — so a proposed, correctly-formed pair can carry
a score of 0.08 and still be a pair, and the grid shows it as
`Check this · 8%`. Whether a document is unmatched is decided by the assignment,
not readable off its score; the score says how much evidence there is, and for a
filename-only pair the honest answer is *very little*, which is exactly what the
grid should show.

The two bands **above** unmatched are still score-and-margin conditions, and
`auto-pair` additionally requires content evidence: a pair with none can never be
pre-ticked, whatever its filenames look like.

**Bands change presentation only.** No band removes a candidate from the grid, and there is no filter that hides low-confidence pairs — the same stance ADR-0017 takes on alignment confidence, for the same reason.

### 6.6 The pairing grid — the confirmation gate

Rendered before any comparison runs. For every document on both sides it shows the proposed counterpart, the score, the band, and (outside auto-pair) the top two rivals with their scores. The user can re-point any pair to any document on the other side, mark either side unmatched, and drop a pair from the run. **The "Run comparisons" action is disabled until the grid is confirmed.**

**Amended during 17b — where the re-point control lives, and when it is live.**
The control is rendered on rows that have a side-A document, and offers every
side-B document it can actually pair with — an *unreadable* one is excluded,
because the assignment never sees it (§8 case 2), so choosing it would be a
selection silently discarded — plus *"Nothing — leave unmatched"*. A side-B document with no
counterpart therefore has no control of its own: it is paired by choosing it in
some side-A row, and unpaired by choosing something else there. Every reachable
pairing is reachable that way, and one control per correspondence is one place
for the grid and the user to disagree instead of two. The control, and the
"change the pairing" action beside it, are **withdrawn for the duration of a
run**: a run consumes the pairing it was given, and an edit landing mid-batch
leaves finished rows and a written index describing a pairing the grid above them
no longer shows. Cancel first.

The grid is also the honest place for the two things the run cannot do:

- A pair whose two sides are **different formats** is shown, paired, and marked `unrunnable` with the reason *"cross-format comparison is out of scope (ADR-0017 Phase E ungated)"*. The pairing is still surfaced — identifying the correspondence is useful even when the comparison cannot run — and it is never silently skipped.
- **Amended during 17b:** a pair whose two sides **agree** on a format this build does not batch — today PDF against PDF (§2.2) — is marked `unrunnable` too, but with its own reason: *"only Word documents can be compared in a batch"*. It is deliberately NOT called cross-format. Two PDFs are not a mismatch, and a reason that is untrue is worse than no reason: it sends the reader looking for a discrepancy that does not exist. The block kinds are `cross-format`, `unsupported-format` and `cap`, each with its own sentence and its own batch-level disclosure.
- A pair that would breach a cap is shown with the cap named. **Amended during 17b:** the cap is checked on *both* sides — side B's bytes are never read, but its size is still a limit the pair has to satisfy — and it demotes **its own pair only**. An over-cap document must never end a confirmed run: the grid has already shown that pair with the cap named, so the user confirmed the run knowing it would not be compared, and refusing the whole batch at that point is the many-to-many form of losing ninety-nine good redlines to one bad file (§8 case 2). The aggregate `MAX_BATCH_TOTAL_BYTES` cap stays a whole-batch refusal, because it is about how much the run would hold at once rather than about one unusable document.

### 6.7 Running

Confirmed pairs go straight into the §5 batch runner, with `mode: 'many-to-many'` and row labels `"A ▸ B"`.

**Amended during 17b — a pair's records are built by the lane that runs it.**
This section first said both sides were preprocessed once during sketch building
so *no document is preprocessed twice* — the saving 1:N gets on its template,
applied to both bundles. A review measured what that costs on the many side: the
sketch has to hold each document's `ParaRecord[]`, so at the §4.4 cap of fifty
documents a side, **a hundred document-sized record arrays are live from the
moment the pairing is proposed until the driver is discarded**. §4.3 opens by
naming that exact shape — "retaining *N* results is an out-of-memory crash, not a
slowdown" — and step 4 promises a peak of one pair's working set times the pool
size.

The two claims cannot both hold, and they are not comparable: one is a slowdown,
the other is a crash. So the sketch keeps the **hashes** and drops the records,
and `ConfirmedPairInput` takes a `records()` thunk the lane calls when it picks
the pair up. Each document is preprocessed twice — once for its hashes, once for
the pair it lands in — and never more than the pool's worth of record arrays
exist at a time. The 1:N saving is untouched: there the reused side is *one*
document, and one is not *N*.

A sketch is reused only while the flatten policy it was built under still holds. Reusing one across a flatten change would leave records preprocessed under the old policy while the index states the new one — an index claiming a flatten choice its redlines were not produced under, which is the thing §8 case 7 puts flatten on the index to prevent.

**Opening a pair** (§2.1 capability 8) re-runs that single pair's *pipeline* and nothing else. The batch drops every pair's ops the moment it completes, so the change list is recovered rather than remembered; the drill-down must never take the artefact path, because packaging a redline on web means triggering a download, and "show me what changed" must not save a file. Both modes have this.

---

## 7. Config keys (feature-local)

All under `compare.batch.*`, resolved through the existing `CompareConfig` resolver. Every default is a named constant in `compare/batch/constants.ts` — **no magic literals**.

**"Resolved" means CONSUMED.** A review found the pairing keys resolved
correctly and then ignored: every pairing module imported the raw constants, so
setting `autoPairScore` in a profile did nothing whatsoever, and the resolver's
own test passed either way. The resolved settings are now threaded through
`assignPairs`, `scorePair`, `rankCandidates`, `bandFor` and `isSketchWorthy`, and
the tunable constants may be imported by `batch/config.ts` — which owns the
defaults — and nowhere else under the batch tree. That import wall is enforced by
a test, because "resolved but unread" is a failure no amount of resolver testing
can see.

`compare.batch.pairing.noMatchScore` also determines two derived values, which
are functions of it rather than constants beside it: the dummy price
`NO_MATCH_COST` (§6.4) and the filename-only ceiling `NO_CONTENT_CEILING` (§6.3,
clamped to 1).
A profile that moves the threshold moves both, or the relationships they encode
come apart.

| Key | Default | Notes |
|---|---|---|
| `compare.batch.maxVariants` | `100` | §4.4 |
| `compare.batch.maxBundleDocuments` | `50` | §4.4 |
| `compare.batch.poolSize` | `0` | `0` = auto (§4.2) |
| `compare.batch.pairing.weightContent` | `0.80` | §6.3 |
| `compare.batch.pairing.weightFilename` | `0.20` | §6.3 |
| `compare.batch.pairing.tieEpsilon` | `0.02` | §6.3 |
| `compare.batch.pairing.autoPairScore` | `0.85` | §6.5 |
| `compare.batch.pairing.autoPairMargin` | `0.15` | §6.5 |
| `compare.batch.pairing.noMatchScore` | `0.25` | §6.4 |
| `compare.batch.pairing.noContentMatchScore` | `0.20` | Phase-2 bar (§6.4) |
| `compare.batch.pairing.minSketchTokens` | `4` | §6.2 |

No new **locale** keys. The classifier reads the same SG/UK `compare.*` keys it always has.

---

## 8. Edge cases

1. **A variant identical to the template** — a real, valid result: zero changes, `status: 'done'`, `changeCount: 0`. Never an error, never omitted from the index.
2. **A corrupt or non-`.docx` file in the batch** — that row is `failed` with a reason; **the batch continues**. One bad file must not lose ninety-nine good redlines.

**Amended during 17b — an unreadable document (§8 case 2) in a many-to-many
bundle.** The sketch build catches per DOCUMENT, not per bundle: a file named
`.docx` that is not a Word package gets a sketch carrying its name and a reason,
is kept out of the assignment entirely — there is nothing to pair on and nothing
to compare — and appears in the grid as its own row with that reason, alongside a
batch-level disclosure. Every other document in both bundles is paired as
normal. A single `try` around the whole ingest loop made one such file empty the
grid, which is the pairing-stage form of losing ninety-nine good redlines to one
bad file.

**Amended during 17b — a document over the per-document cap is never read.** It
is going to be shown with the cap named and never compared, so it gets a
filename-only sketch built from its size alone: reading and preprocessing it
would spend every byte and then retain its records for the driver's life, which
is exactly the peak §4.3 exists to bound and the partial work §4.4 asks the guard
to precede. It is still paired, still shown, and still told why it will not run.

3. **Duplicate file names** — rows are keyed by a batch-stable id, not the name; download filenames are de-duplicated with a numeric suffix.
4. **Two genuinely near-identical documents in a bundle** (a `_v6` and a `_v7` of the same agreement) — the exact case the `AUTO_PAIR_MARGIN` exists for: high absolute score, low margin, so it lands in **propose-with-rivals** and the user sees both candidates. This must have a dedicated fixture and test.
5. **Bundles of unequal size** — the dummy padding handles it; the surplus documents come out `unmatched`.
6. **An empty bundle or a single-document bundle** — valid; M:N with one document each degenerates to a single pair, still via the grid.

   **Amended during 17b — an EMPTY side is refused at the surface, not paired.**
   The pairing layer handles it correctly (every document on the non-empty side
   comes back unmatched, and there is a test), but the driver asks for documents
   on both sides before it will propose. A grid of five rows all reading "no
   counterpart in the other bundle" is a true answer to a question nobody asked:
   the user has not finished loading. Nothing is hidden — the documents are all
   still listed on the surface, and the message says exactly what is missing.
7. **Documents with pre-existing tracked changes** — the per-side flatten choice (feature 08 §5.0) applies **per batch**, not per pair, and is stated on the batch index. A batch that silently flattened differently across variants would be incomparable with itself.
8. **A pair breaching a cap mid-batch** (cumulative total bytes) — the batch stops accepting new work, keeps completed rows, and reports the cap; it does not roll back.
9. **The worker pool unavailable** (`file://`, `VITE_DP_INPROCESS`, construction throws) — the sequential in-process client runs the same batch, slower, with the same results. A test must assert result-identity between pooled and in-process runs.
10. **Cancellation mid-pair** — that pair's partial state is discarded (feature 08 §9.12); completed pairs are kept (§5.5).

---

## 9. Test cases / acceptance criteria

**One-to-many — B1…B13**

| # | Acceptance |
|---|---|
| B1 | Two `createCompareDriver()` instances hold independent state; mutating one does not affect the other, and the exported singleton still behaves exactly as before (existing Compare tests unchanged). |
| B2 | A 1:N run over *N* fixtures preprocesses the template **exactly once** (spy on `preprocessSide`) and produces *N* rows. |
| B3 | The template's `ParaRecord[]` is deep-equal before and after the batch — the shared read-only array is not mutated. |
| B4 | Every `PipelineResult` is released after its pair completes: no reachable reference to any variant's `ParaRecord[]` once its row is `done` (assert on the runner's retained state, not on GC). |
| B5 | Each row's `changeCount` equals that pair's `summary.total`, and `buildBatchIndex` accepts **no filter flags** — the index is filter-independent by construction (mirrors `export.report-model.test.ts`). |
| B6 | Cancelling mid-batch leaves completed rows `done` with their artefacts intact, queued rows `cancelled`, and the in-flight pair discarded. |
| B7 | A corrupt file yields one `failed` row with a reason and does not stop the batch. |
| B8 | Each cap in §4.4 raises `BatchInputTooLargeError` naming the cap, **before** any pair runs. |
| B9 | Pooled and in-process (`WorkerLike` fake / `VITE_DP_INPROCESS`) runs produce identical rows and identical authored output. |
| B10 | A variant identical to the template yields `done` with `changeCount: 0` — present in the index, not omitted. |
| B11 | Past the artefact ceiling every pair is still compared, still counted, and still **downloadable** by re-running that one pair, with a batch-level disclosure. A row marked `artefactOnDemand` is `isDownloadable`. |
| B12 | "Download all" re-runs on-demand rows **lazily, one at a time** — never all up front, which would rebuild the memory peak the ceiling exists to prevent. |
| B13 | The flatten choice is stated on the index, applies to the whole batch, and two variants sharing an id are **refused** rather than collapsed into one row. |

**Many-to-many — M1…M8**

| # | Acceptance |
|---|---|
| M1 | On a bundle whose correct pairing is known, the Hungarian assignment recovers it exactly. |
| M2 | **The deliberately near-identical pair fixture** (`_v6` / `_v7` of one agreement) lands in **propose-with-rivals**, not auto-pair, with both rivals and their scores present. |
| M3 | Unequal bundles: the surplus documents come out `unmatched` (assigned to dummies), not force-matched. |
| M4 | Shuffling the input order of both bundles produces an **identical** assignment (the property greedy fails). |
| M5 | A bundle where every document shares the same cast and defined terms still pairs correctly — proving the cast signal is not dominating (assert it is applied only within `PAIRING_TIE_EPSILON`). |
| M6 | Filename normalisation strips each version token in §6.3 and each day-first date form; a `MM/DD/YYYY` pattern appears nowhere in the source (`no-us-literals` passes). |
| M7 | **No comparison runs before confirmation** — spy on the batch runner and assert zero invocations until the grid is confirmed. |
| M8 | A cross-format pair is surfaced as `unrunnable` **with its pairing shown** and its reason set; it is never silently dropped. |

**Invariants — I1…I3** (both features)

| # | Acceptance |
|---|---|
| I1 | `npm run lint` passes with zero `no-network-egress` and zero `no-host-imports` findings across all new paths — none of which is on the host allow-list. |
| I2 | The Word add-in bundle and behaviour are unchanged: the add-in's tests pass untouched and the batch panels are not reachable from `taskpane/**`. |
| I3 | No suppression anywhere: `rows.length` equals the pairs requested; `changeCount` is never a filtered count; every pairing candidate remains reachable in the grid. |

---

## 10. Known limitations (state these in-product)

- Pairing is **lexical and structural**. Two documents that mean the same thing but share no wording will not pair. This is deliberate — the alternative is opaque judgement (ADR-0008) — and it is why the grid is confirmable.
- `fnv1a` is 32-bit. Jaccard estimates are unaffected in practice, but a paragraph-hash collision is possible; hashes are never used as document identity.
- Cross-format pairs are identified but not compared, pending ADR-0017's unmeasured Phase-E gate.
- Move detection does not cross a pair boundary: a clause relocated from one document into another shows as a deletion in one redline and an insertion in another.
- The batch index is a **list of comparisons, not an analysis across them**. Any question of the form "which clause do most variants change?" is the deferred deviation matrix, and is not answerable here.
- A batch produces the **editable redline only**. The read-only rendered snapshot is produced one pair at a time, not across a batch (§2.2).
- The **flatten choice applies to the whole batch**, not per pair, and is stated on the index because it changes what every redline in that batch says.

---

## 11. Build milestones

| M | Deliverable | Plan |
|---|---|---|
| **M1** | `createCompareDriver()` — the driver becomes instantiable; singleton preserved; concurrency test added. **Blocks everything.** | `17a` |
| **M2** | `compare/batch/limits.ts` + `constants.ts` + `model.ts` + pure `buildBatchIndex`. | `17a` |
| **M3** | The worker pool (protocol / handler / client / entry) with the `WorkerLike` seam and the in-process fallback. | `17a` |
| **M4** | The 1:N batch runner: streamed ingest, preprocess-once, pool dispatch, eager artefact materialisation, result release, progress, cancel. | `17a` |
| **M5** | The 1:N web panel + `WebApp` view + explicit driver injection; per-row download, streamed "download all", drill-down re-run. | `17a` |
| **M6** | `PairingSketch` builder + the three signals + exact Jaccard + filename normalisation. | `17b` |
| **M7** | Hungarian assignment with no-match padding + the three confidence bands. | `17b` |
| **M8** | The pairing grid and its confirmation gate; cross-format and cap surfacing. | `17b` |
| **M9** | M:N runner wiring (both bundles preprocessed once) + acceptance M1–M8. | `17b` |

Delivery: 1:N first, then M:N. Each ends green (lint / typecheck / test) and reviewed against this spec.
