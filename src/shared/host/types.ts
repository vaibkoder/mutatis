/**
 * The HostPort seam (ports-and-adapters / hexagonal architecture — ADR-0022).
 *
 * The feature core depends on THIS interface, never on `Office.*` / `Word.*`
 * directly. Each delivery surface supplies an adapter:
 *   - `host-office` — Office JS inside the Word task pane (today's add-in).
 *   - `host-web`    — browser file-picker + OOXML parse + in-browser view
 *                     (the standalone web app / PWA; spec/features/12 phase 2+).
 *
 * This module is host-NEUTRAL by construction: it must never reference an
 * Office or Word symbol. The `no-host-imports` lint rule enforces that
 * (eslint-plugins/invariants.cjs) — `src/shared/host/**` is deliberately NOT on
 * the allow-list.
 *
 * (spec/features/12-standalone-implementation.md §3.)
 */
import type { DocSnapshot } from '@shared/doc-model/index.js';
import type { DocxBundle, ReadOptions, WriteOptions } from '@shared/ooxml/index.js';
import type { StorageProvider } from '@shared/storage/index.js';

/**
 * Host-neutral document location — replaces `Word.Range` in `Finding.location`.
 * Carries only data (a paragraph id + an optional character span + a label);
 * each adapter resolves it to something navigable its own way (office:
 * `resolveAnchor(...).select()`; web: scroll/highlight in its own view).
 */
export interface DocLocation {
  /** Paragraph id from the DocModel (`ParaNode.id`). Absent for document-level findings. */
  paraId?: string;
  /** Inclusive start offset into the paragraph's normalised text. */
  charStart?: number;
  /** Exclusive end offset into the paragraph's normalised text. */
  charEnd?: number;
  /** Human-readable label, e.g. "Clause 4.2(a)". */
  label: string;
}

/** A cancellable subscription. */
export interface Disposable {
  dispose(): void;
}

/**
 * The in-memory document model the core queries. Today this is the single-pass
 * `DocSnapshot` (`ParaNode[]` + footnotes); the alias keeps the port's vocabulary
 * host-neutral and lets the model evolve without changing the seam.
 */
export type DocModel = DocSnapshot;

/**
 * Where a free-text search looks (ADR-0047). A header or footer names the
 * section (0-based, in the compared documents) and the reference type; a note
 * names only its kind — a host searches every note of that kind.
 */
export type SearchScope =
  | { story: 'header' | 'footer'; section: number; type: 'default' | 'first' | 'even' }
  | { story: 'footnote' | 'endnote' };

/**
 * The single seam between the host-neutral feature core and its environment.
 * The small set of capabilities the core needs: ingest the document into the
 * model, observe document changes, reveal a host-neutral location, read/write
 * the .docx package, and persist key/value + document-scoped storage.
 */
export interface HostPort {
  // ingest + model
  /** Load the document into the in-memory model. */
  loadDocument(): Promise<DocModel>;
  /** Subscribe to document-change notifications (debounced by the adapter). */
  onDocumentChanged(cb: () => void): Disposable;

  // navigation (replaces range.select())
  /** Navigate to / highlight a host-neutral location. */
  reveal(loc: DocLocation): Promise<void>;

  /**
   * Navigate to a change by free-text search of the live host document.
   * Compare's Word-vs-Word `goToChange` locates a change by its verbatim text
   * (not by a `paraId` anchor), which {@link reveal} cannot express. Optional:
   * only hosts with a live, navigable document implement it.
   *   - office: `body.search(text).select()`.
   *   - web: omitted (undefined) — the web target has no live host doc to
   *     navigate; Word-vs-Word output is a download, so this is a clean no-op.
   *
   * `scope` names the story the change lives in (ADR-0047): a header or footer
   * of a section, or the footnotes/endnotes. Absent ⇒ the main text.
   */
  findAndReveal?(searchText: string, scope?: SearchScope): Promise<void>;

  // OOXML read-modify-write (the unzip/parse/edit middle is already host-agnostic)
  /** Read the live document into an editable bundle. */
  readDocx(opts?: ReadOptions): Promise<DocxBundle>;
  /** Write a modified bundle back (office: insert in place; web: Blob download). */
  writeBack(bundle: DocxBundle, opts?: WriteOptions): Promise<void>;

  // storage (mirrors the existing StorageProvider)
  storage: StorageProvider;
}
