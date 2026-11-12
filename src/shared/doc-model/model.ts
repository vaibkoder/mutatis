import { signal, computed } from '@preact/signals';
import { loadDocSnapshot } from './loader.js';
import type { DocSnapshot, ParaNode } from './types.js';

type InvalidateCallback = () => void;

const DEBOUNCE_MS = 500;

/**
 * DocModel — the single in-memory document representation.
 * All features read from here; none traverse the document independently.
 * (spec/ARCHITECTURE.md §Document Model / Paragraph Index)
 */
class DocModelImpl {
  private _snapshot = signal<DocSnapshot | null>(null);
  private _loading = signal(false);
  private _invalidateCallbacks = new Set<InvalidateCallback>();
  private _debounceTimer: ReturnType<typeof setTimeout> | null = null;
  private _eventHandlers: (() => void)[] = [];

  /** Reactive current snapshot. null before first load. */
  readonly snapshot = computed(() => this._snapshot.value);

  /** All body paragraphs. Empty array before first load. */
  readonly paragraphs = computed<ReadonlyArray<ParaNode>>(
    () => this._snapshot.value?.paragraphs ?? []
  );

  /** All footnote paragraphs. */
  readonly footnotes = computed<ReadonlyArray<ParaNode>>(
    () => this._snapshot.value?.footnotes ?? []
  );

  /** All paragraphs (body + footnotes). */
  readonly allParagraphs = computed<ReadonlyArray<ParaNode>>(() => [
    ...(this._snapshot.value?.paragraphs ?? []),
    ...(this._snapshot.value?.footnotes ?? []),
  ]);

  /** Whether a load is in progress. */
  readonly loading = computed(() => this._loading.value);

  private _loadPromise: Promise<void> | null = null;

  /** Register a callback to be called when the model is invalidated. */
  onInvalidate(cb: InvalidateCallback): () => void {
    this._invalidateCallbacks.add(cb);
    return () => this._invalidateCallbacks.delete(cb);
  }

  /** Load (or reload) the document. Concurrent callers share the same in-flight promise. */
  async load(): Promise<void> {
    if (this._loadPromise) return this._loadPromise;
    this._loading.value = true;
    this._loadPromise = (async () => {
      try {
        this._snapshot.value = await loadDocSnapshot();
      } finally {
        this._loading.value = false;
        this._loadPromise = null;
      }
      this._notifyInvalidate();
    })();
    return this._loadPromise;
  }

  /** Subscribe to Office paragraph-change events and debounce model invalidation. */
  subscribeToEvents(): void {
    if (typeof Word === 'undefined') return;

    try {
      const onChanged = async (_event: Word.ParagraphChangedEventArgs): Promise<void> => {
        this._scheduleReload();
      };
      const onAdded = async (_event: Word.ParagraphAddedEventArgs): Promise<void> => {
        this._scheduleReload();
      };
      const onDeleted = async (_event: Word.ParagraphDeletedEventArgs): Promise<void> => {
        this._scheduleReload();
      };

      Word.run(async ctx => {
        ctx.document.onParagraphChanged.add(onChanged);
        ctx.document.onParagraphAdded.add(onAdded);
        ctx.document.onParagraphDeleted.add(onDeleted);
        await ctx.sync();
      }).catch(() => { /* event registration not supported on this host */ });
    } catch { /* noop */ }
  }

  /** Find a paragraph by ID. */
  getById(id: string): ParaNode | undefined {
    return this.allParagraphs.value.find(p => p.id === id);
  }

  /**
   * All paragraphs whose style matches one of `styleNames`, in document order.
   * Style comparison is case-insensitive. Pure filter over `allParagraphs`
   * (body then footnotes); returns a new array, never mutates the snapshot.
   * Used by the Content feature for the Index of Schedules (custom-style scan).
   */
  getParagraphsByStyle(styleNames: string[]): ParaNode[] {
    const wanted = new Set(styleNames.map(s => s.toLowerCase()));
    if (wanted.size === 0) return [];
    return this.allParagraphs.value.filter(p => wanted.has(p.styleName.toLowerCase()));
  }

  /** Dispose all event handlers. */
  dispose(): void {
    for (const h of this._eventHandlers) h();
    this._eventHandlers = [];
  }

  private _scheduleReload(): void {
    if (this._debounceTimer !== null) {
      clearTimeout(this._debounceTimer);
    }
    this._debounceTimer = setTimeout(() => {
      this._debounceTimer = null;
      this.load().catch(() => { /* silently retry next invalidation */ });
    }, DEBOUNCE_MS);
  }

  private _notifyInvalidate(): void {
    for (const cb of this._invalidateCallbacks) {
      try { cb(); } catch { /* don't let a bad listener break the model */ }
    }
  }
}

/** Singleton DocModel instance. */
export const DocModel = new DocModelImpl();
