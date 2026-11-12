/**
 * Load both sides of a Compare operation into DocxBundles.
 * Spec: spec/features/08-compare.md §3.2, Architecture Office JS rows 10/12.
 *
 * Invariants:
 * - Original = active document only (read via the active HostPort's readDocx —
 *   office: getFileAsync; web: the in-memory ingested bundle).
 * - Revised  = File-API ArrayBuffer → PizZipBundle (no host read).
 * - The ArrayBuffer is consumed (not retained) after the bundle is constructed
 *   so the GC can reclaim it; the caller must not reuse the buffer.
 * - The 'text' discriminator (feature 12, ADR-0018) synthesises a minimal valid
 *   .docx from a pasted plain-text blob via the T1 synthesiser, so the text path
 *   rides this same loader → pipeline chain.
 * - A 'pdf' discriminator (feature 12, ADR-0017) is wired as a seam (not built here).
 */
import type { DocxBundle } from '@shared/ooxml/index.js';
import { PizZipBundle } from '@shared/ooxml/bundle.js';
import { getActiveHost } from '@shared/host/index.js';

import { synthesiseDocx } from './text/synthesise-docx.js';

// ---------------------------------------------------------------------------
// Source discriminator — 'text' (feature 12) built; 'pdf' seam lives here
// ---------------------------------------------------------------------------

export type CompareSourceKind = 'activeDocument' | 'docxArrayBuffer' | 'text' | 'pdf';

export interface ActiveDocumentSource {
  kind: 'activeDocument';
}

export interface DocxArrayBufferSource {
  kind: 'docxArrayBuffer';
  /** The ArrayBuffer from File API (file.arrayBuffer()). Consumed on load; do not reuse. */
  buffer: ArrayBuffer;
}

/**
 * PDF source (feature 12, ADR-0017). Unlike the docx/text kinds it does NOT
 * resolve to a DocxBundle — PDFs have no OOXML body. The PDF path opens these
 * bytes with the offline PDF.js loader, reconstructs paragraphs, and feeds
 * `runPipelinePdf` directly (see io/pdf/parts.ts → extractPdfParaRecords). So
 * `loadSource` does not handle this kind; the driver routes PDF sides via the
 * dedicated PDF loader BEFORE `loadBothSides` is reached.
 */
export interface PdfSource {
  kind: 'pdf';
  /** The ArrayBuffer from File API (file.arrayBuffer()). Consumed on load; do not reuse. */
  buffer: ArrayBuffer;
}

/**
 * Text-vs-text source (feature 12, ADR-0018): a pasted plain-text blob, adapted
 * into a synthesised .docx bundle so it feeds the unchanged Compare pipeline.
 */
export interface TextSource {
  kind: 'text';
  /** Raw pasted plain text for this side (Original or Revised). */
  text: string;
}

// v2 seam — not implemented; type is present so the discriminated union is extensible.
// export interface PdfSource { kind: 'pdf'; buffer: ArrayBuffer; }

export type CompareSource =
  | ActiveDocumentSource
  | DocxArrayBufferSource
  | TextSource
  | PdfSource;

/** Which .docx is Original (base) and which is Revised (compared). */
export interface CompareDirection {
  /** Source for the Original (base) side. */
  original: CompareSource;
  /** Source for the Revised (compared) side. */
  revised: CompareSource;
}

/** Loaded bundles for both sides, ready for pre-processing. */
export interface LoadedSides {
  originalBundle: DocxBundle;
  revisedBundle: DocxBundle;
}

/**
 * Load both sides from the given direction description.
 *
 * Constraint: the active document may be used as EITHER side (not both).
 * If both sources are 'activeDocument', only one host read is made and the same
 * bundle is used for both sides — the caller's responsibility is to not mutate it.
 */
export async function loadBothSides(direction: CompareDirection): Promise<LoadedSides> {
  const [originalBundle, revisedBundle] = await Promise.all([
    loadSource(direction.original),
    loadSource(direction.revised),
  ]);
  return { originalBundle, revisedBundle };
}

async function loadSource(source: CompareSource): Promise<DocxBundle> {
  switch (source.kind) {
    case 'activeDocument':
      return getActiveHost().readDocx();

    case 'docxArrayBuffer': {
      // Construct the bundle from the ArrayBuffer synchronously; the PizZipBundle
      // constructor copies the data into PizZip's internal representation so the
      // buffer can be freed by the GC after this call.
      const bundle = new PizZipBundle(source.buffer);
      // Explicitly let the GC collect the original buffer.
      // (source.buffer is not nulled here since TS doesn't allow it on a readonly
      //  field — callers should not hold a reference to the buffer after this call.)
      return bundle;
    }

    case 'text':
      // Feature 12 (ADR-0018): synthesise a minimal valid .docx from plain text.
      // Pure + synchronous; wrapped here so all source kinds share one async seam.
      return synthesiseDocx(source.text);

    case 'pdf':
      // Feature 12 (ADR-0017): PDFs have no DocxBundle — the driver loads PDF
      // sides via io/pdf/extractPdfParaRecords → runPipelinePdf, never here.
      throw new Error(
        'loadSource: PDF sources do not resolve to a DocxBundle; load them via ' +
          'extractPdfParaRecords and runPipelinePdf (the PDF path), not loadBothSides.',
      );

    default: {
      const exhaustiveCheck: never = source;
      throw new Error(`loadSource: unsupported source kind '${(exhaustiveCheck as CompareSource).kind}'`);
    }
  }
}
