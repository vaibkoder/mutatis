/**
 * Feature 08 — Compare: styles.xml diff (I2).
 * Spec: spec/features/08-compare.md §9.9.
 *
 * This module's scope is a change to a style DEFINITION — `styles.xml` itself, a
 * style's font or spacing redefined — reported as 'formatting' with a "style
 * definition changed" note.
 *
 * A change to which style a PARAGRAPH uses is a different thing and belongs to
 * the body pipeline. It is handled there by `diff/property.ts` (ADR-0042), which
 * turns a text-identical pair with a differing `styleId` into a `property` op.
 *
 * **Correction, 2026-09-12.** This header used to claim per-paragraph styleId
 * changes were "handled in the main body pipeline (A1 sets styleId on
 * ParaRecord; I1 classifier B1 detects formatting-only changes)". That was
 * false for the whole time it stood: A1 did set `styleId`, but nothing read it,
 * and B1 only LABELS ops that already exist — a style-only change produced no op
 * to label, so it reached no classifier and appeared nowhere. The sentence is
 * true now because ADR-0042 made it true, not because it ever was.
 */
import type { DiffOp } from './model.js';

export interface StylesDiffResult {
  /** True when styles.xml differs between the two sides. */
  changed: boolean;
  /** A single synthetic DiffOp summarising the styles change, or null if unchanged. */
  op: DiffOp | null;
  /** Human-readable note for the summary panel. */
  note: string | null;
}

let _idSeq = 9_000_000; // start well outside body-op range

/**
 * Compare the styles.xml parts of two packages.
 * Returns a synthetic formatting-only DiffOp when styles differ.
 *
 * Pure: no Office JS.
 */
export function diffStyles(
  originalStylesXml: string | null,
  revisedStylesXml:  string | null,
  author: string,
  date: string,
): StylesDiffResult {
  if (originalStylesXml === null || revisedStylesXml === null) {
    return { changed: false, op: null, note: null };
  }
  if (originalStylesXml === revisedStylesXml) {
    return { changed: false, op: null, note: null };
  }

  const op: DiffOp = {
    id: `styles-${_idSeq++}`,
    type: 'replace',
    origRange: null,
    revRange:  null,
    partId: 'styles',
    wordDiffs: [],
    classification: 'formatting',
    classes: ['formatting'],
    author,
    date,
  };

  return {
    changed: true,
    op,
    note: 'Style definitions changed (formatting-only) — check styles.xml for font/spacing changes.',
  };
}
