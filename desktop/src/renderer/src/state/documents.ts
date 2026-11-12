import {
  MAX_BUNDLE_DOCUMENTS,
  MAX_DOCX_BYTES,
  exceedsDocxByteCap,
} from '@features/compare/batch/limits.js';

/**
 * The document set behind the landing page.
 *
 * Deliberately shaped for 1:1, 1-to-many AND many-to-many from the start: each
 * side is a LIST. A 1:1 comparison is just the case where both lists have one
 * entry, so nothing has to be reshaped later to support batches.
 *
 * Caps are imported from the engine rather than restated here. Restating them
 * would let the two drift, and the UI would then promise something the runner
 * refuses (or refuse something it would have accepted).
 */

export type DocSide = 'original' | 'modified';

/** What the compare engine can actually consume. */
export type DocKind = 'docx' | 'pdf' | 'text';

export interface LoadedDoc {
  readonly id: string;
  readonly name: string;
  readonly byteLength: number;
  readonly kind: DocKind;
  /** Kept unread. Bytes are pulled lazily when a comparison actually runs. */
  readonly file: File;
  /**
   * Set when the document is over the per-document cap. It is still listed —
   * never silently dropped — and the reason travels with it.
   */
  readonly overCap: boolean;
  /**
   * How the document was read, when that is worth telling the user — set for a
   * PDF, which is rebuilt from page layout. Shown above any redline it is in.
   */
  readonly notes?: readonly string[];
  /** Appended to the name the assistant sees, e.g. "PDF: text rebuilt from page layout". */
  readonly aiNote?: string;
}

export interface DocumentSet {
  readonly original: readonly LoadedDoc[];
  readonly modified: readonly LoadedDoc[];
}

export const emptyDocumentSet: DocumentSet = { original: [], modified: [] };

const EXTENSIONS: Record<string, DocKind> = {
  docx: 'docx',
  pdf: 'pdf',
  txt: 'text',
  md: 'text',
};

/** A rejected file, with the reason kept so the UI can say it out loud. */
export interface RejectedFile {
  readonly name: string;
  readonly reason: string;
}

export interface AcceptResult {
  readonly accepted: readonly LoadedDoc[];
  readonly rejected: readonly RejectedFile[];
}

function extensionOf(name: string): string {
  const dot = name.lastIndexOf('.');
  return dot === -1 ? '' : name.slice(dot + 1).toLowerCase();
}

function describeBytes(n: number): string {
  const mib = n / (1024 * 1024);
  return mib >= 1 ? `${mib.toFixed(1)} MiB` : `${Math.max(1, Math.round(n / 1024))} KiB`;
}

let idCounter = 0;

/**
 * Classify dropped files for one side.
 *
 * Nothing is dropped silently. A file we cannot use comes back in `rejected`
 * with a sentence explaining why, and an oversize `.docx` is ACCEPTED but
 * flagged — the engine's own cap helpers do the same thing, because a document
 * the user handed us must stay visible even when it cannot be compared.
 */
export function acceptFiles(
  files: readonly File[],
  existing: readonly LoadedDoc[],
): AcceptResult {
  const accepted: LoadedDoc[] = [];
  const rejected: RejectedFile[] = [];
  let room = MAX_BUNDLE_DOCUMENTS - existing.length;

  for (const file of files) {
    const ext = extensionOf(file.name);
    const kind = EXTENSIONS[ext];

    if (!kind) {
      rejected.push({
        name: file.name,
        reason: ext
          ? `.${ext} is not a supported format — use .docx, .pdf, .txt or .md`
          : 'no file extension, so the format could not be determined',
      });
      continue;
    }

    if (existing.some((d) => d.name === file.name && d.byteLength === file.size)) {
      rejected.push({ name: file.name, reason: 'already added to this side' });
      continue;
    }

    if (room <= 0) {
      rejected.push({
        name: file.name,
        reason: `over the ${MAX_BUNDLE_DOCUMENTS}-document limit for one side`,
      });
      continue;
    }

    const overCap = kind === 'docx' && exceedsDocxByteCap(file.size);
    accepted.push({
      id: `doc-${++idCounter}`,
      name: file.name,
      byteLength: file.size,
      kind,
      file,
      overCap,
    });
    room -= 1;
  }

  return { accepted, rejected };
}

/** The sentence shown beside a document that is over the per-document cap. */
export function overCapReason(doc: LoadedDoc): string {
  return (
    `${describeBytes(doc.byteLength)} exceeds the ${describeBytes(MAX_DOCX_BYTES)} ` +
    `per-document cap, so this document will not be compared.`
  );
}

/** How the current selection would be run, once both sides have documents. */
export type ComparisonShape = 'none' | 'tracked-changes' | 'one-to-one' | 'one-to-many' | 'many-to-many';

export function comparisonShape(set: DocumentSet): ComparisonShape {
  const o = set.original.length;
  const m = set.modified.length;
  if (o + m === 1) return 'tracked-changes';
  if (o === 0 || m === 0) return 'none';
  if (o === 1 && m === 1) return 'one-to-one';
  if (o === 1 || m === 1) return 'one-to-many';
  return 'many-to-many';
}

export function describeShape(shape: ComparisonShape): string {
  switch (shape) {
    case 'one-to-one':
      return 'One-to-one comparison';
    case 'one-to-many':
      return 'One-to-many batch';
    case 'many-to-many':
      return 'Many-to-many bundle — pairing will be proposed for review';
    case 'tracked-changes':
      return 'One document: its tracked changes will be compared';
    case 'none':
      return 'Add a document to each side';
  }
}

export { MAX_BUNDLE_DOCUMENTS, describeBytes };
