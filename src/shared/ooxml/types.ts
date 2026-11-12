/**
 * The read-modify-write OOXML loop, behind a clean API.
 * No document bytes ever leave the machine. (ADR-0001/ADR-0003, spec/ARCHITECTURE.md §The OOXML Service)
 * Consumed by Repair (v1) and Compare (v2).
 */

/** An in-memory editable .docx package. */
export interface DocxBundle {
  /** Read a part as a string, e.g. 'word/document.xml'. Undefined if absent. */
  part(path: string): string | undefined;
  /**
   * Read a part as raw bytes, e.g. an embedded `word/fonts/font1.odttf`. Undefined
   * if absent. Optional: not every `DocxBundle` carries binary parts (text-only
   * stubs may omit it); callers that need bytes must degrade gracefully when it is
   * absent or returns undefined.
   */
  partBytes?(path: string): Uint8Array | undefined;
  /** Replace a part's XML. */
  setPart(path: string, xml: string): void;
  /** Whether a part exists. */
  hasPart(path: string): boolean;
  /** List all part paths. */
  listParts(): string[];
  /** Remove a part. */
  removePart(path: string): void;
  /** Strip named attributes from a part (e.g. w:rsid*) by targeted regex, preserving fidelity. */
  stripAttributes(partPath: string, attrNames: string[]): void;
}

export interface ReadOptions {
  /** Slice size in bytes for getFileAsync. */
  sliceSize?: number;
}

export interface WriteOptions {
  /** 'replace' inserts in place; 'newDocument' opens a fresh document. */
  mode?: 'replace' | 'newDocument';
  /** What to do if an in-place artifact is detected. */
  onArtifactDetected?: 'fallbackToNewDoc' | 'throw';
}

export interface OoxmlService {
  /** Read the live document into an editable bundle. */
  readDocx(opts?: ReadOptions): Promise<DocxBundle>;
  /** Write a modified bundle back. */
  writeBack(bundle: DocxBundle, opts?: WriteOptions): Promise<void>;
}
