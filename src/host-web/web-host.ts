/**
 * host-web — the browser `HostPort` adapter (ADR-0022, spec/features/12 §3, phase 2).
 *
 * The standalone counterpart to `host-office`: instead of Office JS it backs the
 * seam with browser primitives —
 *   - ingest:   a file-picked `.docx` ArrayBuffer → `PizZipBundle` (held in memory).
 *   - model:    `loadDocument()` parses the held bundle into the shared DocSnapshot
 *               (`loadDocSnapshotFromOoxml`).
 *   - reveal:   scroll + transient highlight of the paragraph in the app's own
 *               document view (the in-app surface keyed by `data-para-id`).
 *   - read/write: `readDocx()` returns the held bundle; `writeBack()` serialises it
 *               to a `.docx` Blob and triggers a download (no in-place mutation —
 *               the web target is analyse-and-export, ADR-0022).
 *   - storage:  localStorage-backed (web-storage.ts).
 *
 * Host-NEUTRAL by construction: no `Office.*`/`Word.*` token appears here, so this
 * file stays OFF the `no-host-imports` allow-list and is lint-proven Office-free.
 */
import { PizZipBundle } from '@shared/ooxml/bundle.js';
import type { DocxBundle, WriteOptions } from '@shared/ooxml/index.js';
import { loadDocSnapshotFromOoxml } from './ooxml-loader.js';
import { createWebStorageProvider } from './web-storage.js';
import type {
  DocLocation,
  DocModel,
  Disposable,
  HostPort,
} from '@shared/host/index.js';
import type { StorageProvider } from '@shared/storage/index.js';

const DOCX_MIME =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

/** CSS class the document view styles as the active (revealed) paragraph. */
const ACTIVE_PARA_CLASS = 'dp-web-doc__para--active';

/** `HostPort` backed by the browser (file-picker ingest + in-browser view + download). */
export class WebHostPort implements HostPort {
  private bundle: PizZipBundle | null = null;
  private fileName = 'document.docx';
  private readonly listeners = new Set<() => void>();
  readonly storage: StorageProvider = createWebStorageProvider();

  /**
   * host-web-specific: take a picked `.docx` as bytes. The web app calls this from
   * its file-input handler; the rest of the `HostPort` surface then operates on the
   * held bundle. Notifies document-changed subscribers.
   */
  ingest(buffer: ArrayBuffer, fileName?: string): void {
    this.bundle = new PizZipBundle(buffer);
    if (fileName) this.fileName = fileName;
    for (const cb of this.listeners) cb();
  }

  /** Whether a document has been ingested yet. */
  hasDocument(): boolean {
    return this.bundle !== null;
  }

  /** Parse the held bundle into the shared DocSnapshot model. */
  async loadDocument(): Promise<DocModel> {
    if (!this.bundle) {
      throw new Error('WebHostPort.loadDocument: no document ingested. Call ingest() first.');
    }
    return loadDocSnapshotFromOoxml(this.bundle);
  }

  /** Subscribe to document-change (re-ingest) notifications. */
  onDocumentChanged(cb: () => void): Disposable {
    this.listeners.add(cb);
    return { dispose: () => { this.listeners.delete(cb); } };
  }

  /**
   * Navigate to a host-neutral location by scrolling its paragraph into view and
   * applying a transient highlight in the app's own document view. Best-effort: a
   * missing element (no view rendered, or a stale id) is a silent no-op — matching
   * the office adapter's stale-range tolerance.
   */
  async reveal(loc: DocLocation): Promise<void> {
    if (!loc.paraId || typeof document === 'undefined') return;
    const target = document.querySelector<HTMLElement>(`[data-para-id="${cssEscape(loc.paraId)}"]`);
    if (!target) return;
    for (const prev of document.querySelectorAll(`.${ACTIVE_PARA_CLASS}`)) {
      prev.classList.remove(ACTIVE_PARA_CLASS);
    }
    target.classList.add(ACTIVE_PARA_CLASS);
    if (typeof target.scrollIntoView === 'function') {
      target.scrollIntoView({ block: 'center', behavior: 'smooth' });
    }
  }

  /** Return the held bundle (the web counterpart of getFileAsync). */
  async readDocx(): Promise<DocxBundle> {
    if (!this.bundle) {
      throw new Error('WebHostPort.readDocx: no document ingested. Call ingest() first.');
    }
    return this.bundle;
  }

  /**
   * Write a bundle back as a `.docx` download (no in-place mutation). The web target
   * is analyse-and-export: the modified/round-tripped package is handed to the user
   * as a file, never sent anywhere (ADR-0001/0003/0022).
   */
  async writeBack(bundle: DocxBundle, _opts?: WriteOptions): Promise<void> {
    void _opts;
    if (!(bundle instanceof PizZipBundle)) {
      throw new Error('WebHostPort.writeBack: expected a PizZipBundle.');
    }
    // Copy into a fresh ArrayBuffer-backed view so the Blob part type is exact.
    const bytes = new Uint8Array(bundle.toUint8Array());
    const blob = new Blob([bytes], { type: DOCX_MIME });
    triggerDownload(blob, this.fileName);
  }
}

/** Trigger a browser download of a Blob via a transient anchor. */
function triggerDownload(blob: Blob, fileName: string): void {
  if (typeof document === 'undefined' || typeof URL.createObjectURL !== 'function') return;
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

/** Escape a paragraph id for use inside a CSS attribute selector. */
function cssEscape(value: string): string {
  if (typeof CSS !== 'undefined' && typeof CSS.escape === 'function') return CSS.escape(value);
  return value.replace(/["\\]/g, '\\$&');
}
