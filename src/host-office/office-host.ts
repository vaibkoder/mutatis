/**
 * host-office — the Office JS adapter implementing `HostPort` (ADR-0022).
 *
 * This is today's add-in behaviour, expressed behind the seam. It COMPOSES the
 * existing shared modules (it does not relocate them — the lighter-boundary
 * structure): the single-pass loader drives the DocModel, `reveal`
 * reuses the `resolveAnchor` search-and-select machinery (the exact sequence
 * `FindingRow` used inline), and ingest/storage delegate to the OOXML Service
 * and the storage provider. Swapping this for `host-web` (phase 2+) is the only
 * thing the standalone target changes.
 *
 * (spec/features/12-standalone-implementation.md §3.)
 */
import { isCapable, wordRun } from '@shared/office/index.js';
import { DocModel, resolveAnchor } from '@shared/doc-model/index.js';
import type { TextAnchor } from '@shared/doc-model/index.js';
import { ooxmlService } from '@shared/ooxml/index.js';
import type { DocxBundle, ReadOptions, WriteOptions } from '@shared/ooxml/index.js';
import { getStorageProvider } from '@shared/storage/index.js';
import type { StorageProvider } from '@shared/storage/index.js';
import type { DocLocation, DocModel as DocModelType, Disposable, HostPort, SearchScope } from '@shared/host/index.js';

/** `HostPort` backed by Office JS inside the Word task pane. */
export class OfficeHostPort implements HostPort {
  /** Load the document into the shared DocModel and return its snapshot. */
  async loadDocument(): Promise<DocModelType> {
    await DocModel.load();
    const snapshot = DocModel.snapshot.value;
    if (!snapshot) {
      throw new Error('OfficeHostPort.loadDocument: DocModel produced no snapshot');
    }
    return snapshot;
  }

  /** Notify on document change. Wraps the DocModel's debounced invalidation. */
  onDocumentChanged(cb: () => void): Disposable {
    const off = DocModel.onInvalidate(cb);
    return { dispose: off };
  }

  /**
   * Navigate to a host-neutral location by re-resolving a live Range and
   * selecting it. Reproduces the prior `FindingRow.navigateTo` sequence exactly:
   * a first sync materialises the deferred search hit, then `select()` + a second
   * sync. A stale location after a document edit is a silent no-op (as before).
   */
  async reveal(loc: DocLocation): Promise<void> {
    if (!loc.paraId) return;
    const paras = DocModel.allParagraphs.value;
    const anchor: TextAnchor = {
      paragraphId: loc.paraId,
      startOffset: loc.charStart ?? 0,
      endOffset: loc.charEnd ?? 0,
    };
    try {
      await wordRun(async ctx => {
        const range = resolveAnchor(ctx, paras, anchor);
        await ctx.sync();
        range.select();
        await ctx.sync();
      });
    } catch {
      // Range may be stale after a document edit; silently ignore.
    }
  }

  /**
   * Navigate to a change by free-text search of the live document. Moved
   * verbatim from Compare's Word-vs-Word `goToChange` (driver.ts): search
   * case-insensitively, materialise the deferred hits with a sync, then select
   * the first match and sync again. Empty search text is a no-op.
   *
   * A `scope` (ADR-0047) searches that story first: the section's header or
   * footer of that type, then every other header or footer (a section break
   * added or removed shifts the redline's section numbers from the compared
   * documents'), and the main text last. Notes are searched through
   * `body.footnotes`/`endnotes` where WordApi 1.5 offers them; without it the
   * search falls back to the main text — the change list's tooltip says so.
   */
  async findAndReveal(searchText: string, scope?: SearchScope): Promise<void> {
    const text = searchText.trim();
    if (!text) return;
    await wordRun(async ctx => {
      for (const body of await storiesFor(ctx, scope)) {
        const results = body.search(text, { matchCase: false });
        results.load('items/text');
        await ctx.sync();
        if (results.items.length > 0 && results.items[0]) {
          results.items[0].select();
          await ctx.sync();
          return;
        }
      }
    });
  }

  /** Read the live .docx into an editable bundle (Office getFileAsync loop). */
  readDocx(opts?: ReadOptions): Promise<DocxBundle> {
    return ooxmlService.readDocx(opts);
  }

  /** Write a modified bundle back in place (insertFileFromBase64). */
  writeBack(bundle: DocxBundle, opts?: WriteOptions): Promise<void> {
    return ooxmlService.writeBack(bundle, opts);
  }

  /** Local + document-settings storage (lazily initialised). */
  get storage(): StorageProvider {
    return getStorageProvider();
  }
}

const HEADER_FOOTER_TYPE = { default: 'Primary', first: 'FirstPage', even: 'EvenPages' } as const;
const ALL_TYPES = ['default', 'first', 'even'] as const;

/** The bodies to search for a scope, in order; the main text always last. */
async function storiesFor(ctx: Word.RequestContext, scope: SearchScope | undefined): Promise<Word.Body[]> {
  const main = ctx.document.body;
  if (scope === undefined) return [main];
  if (scope.story === 'header' || scope.story === 'footer') {
    const { story, section: home, type: homeType } = scope;
    const sections = ctx.document.sections;
    sections.load('items');
    await ctx.sync();
    const of = (k: number, type: (typeof ALL_TYPES)[number]): Word.Body | undefined => {
      const section = sections.items[k];
      if (section === undefined) return undefined;
      const kind = HEADER_FOOTER_TYPE[type];
      return story === 'header' ? section.getHeader(kind) : section.getFooter(kind);
    };
    const rest = sections.items.flatMap((_, k) =>
      ALL_TYPES.filter(t => k !== home || t !== homeType).map(t => of(k, t)));
    return [...[of(home, homeType), ...rest].filter((b): b is Word.Body => b !== undefined), main];
  }
  if (!isCapable('wordApi15')) return [main];
  const notes = scope.story === 'footnote' ? main.footnotes : main.endnotes;
  notes.load('items');
  await ctx.sync();
  return [...notes.items.map(n => n.body), main];
}
