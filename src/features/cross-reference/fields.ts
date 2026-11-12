/**
 * Feature A — feature-local REF-field / bookmark read (ADR-0014).
 *
 * Performs a bounded batched Word.run over only the paragraphs that carry
 * ReferenceMatch entries, reads paragraph.fields (REF/PAGEREF field codes),
 * loads document.bookmarks once, and sets isFieldBacked + staticBadge on each
 * ReferenceMatch, plus detects broken-REF findings (§5.1.4).
 *
 * MANDATORY FALLBACK (ADR-0014): if paragraph.fields is unavailable/unreliable
 * on the host (e.g. older Word builds), degrade gracefully by treating ALL
 * matches as un-backed (staticBadge=true) and emitting a single non-blocking
 * capability banner. NEVER throw; NEVER block the rest of the check.
 *
 * All reads are client-side typed-API reads. No bytes leave the machine.
 */

import type { ParaNode } from '@shared/doc-model/types.js';
import type { ReferenceMatch } from './types.js';

export interface FieldReadResult {
  /** Whether the fields API was available on this host. False triggers the banner. */
  fieldsApiAvailable: boolean;
}

/**
 * In the Office JS environment, run a batched Word.run to read fields/bookmarks.
 * Updates `matches` in-place.
 *
 * Should be called only inside a Word.run block (or creates its own).
 * If called outside Word context (e.g. in unit tests with mock), the caller
 * provides a mock ctx.
 */
export async function readFieldsAndBookmarks(
  matches: ReferenceMatch[],
  paras: ReadonlyArray<ParaNode>,
  ctx: Word.RequestContext,
): Promise<FieldReadResult> {
  if (matches.length === 0) return { fieldsApiAvailable: true };

  // Index paragraphs by ID for fast lookup.
  const paraById = new Map<string, ParaNode>();
  for (const p of paras) paraById.set(p.id, p);

  // Collect unique paragraph IDs that have matches.
  const matchedParaIds = new Set(matches.map(m => m.paragraphId));

  try {
    // Load bookmarks once for the whole document.
    const bookmarks = ctx.document.bookmarks;
    bookmarks.load('items/name');

    // Sync 1: load each matched paragraph's field collection.
    const paraFieldsMap = new Map<string, Word.FieldCollection>();
    for (const paraId of matchedParaIds) {
      const para = paraById.get(paraId);
      if (!para) continue;
      const paraRange = para.getRange(ctx);
      const fields = paraRange.fields;
      fields.load('items');
      paraFieldsMap.set(paraId, fields);
    }
    await ctx.sync();

    // Sync 2: load code + result for every field across all matched paragraphs in one batch.
    for (const fields of paraFieldsMap.values()) {
      for (const field of fields.items) {
        field.load('code,result/text');
      }
    }
    await ctx.sync();

    // Build a set of valid bookmark names.
    const validBookmarks = new Set<string>();
    for (const bm of bookmarks.items) {
      if (bm.name) validBookmarks.add(bm.name);
    }

    // Build, per paragraph, the list of REF/PAGEREF fields with target + result text.
    interface RefField { target: string; resultText: string; }
    const refFieldsByPara = new Map<string, RefField[]>();
    for (const [paraId, fields] of paraFieldsMap) {
      const list: RefField[] = [];
      for (const field of fields.items) {
        const code: string = (field as unknown as { code?: string }).code ?? '';
        const upper = code.trim().toUpperCase();
        if (!upper.startsWith('REF ') && !upper.startsWith('PAGEREF ')) continue;
        const target = code.replace(/^REF\s+|^PAGEREF\s+/i, '').trim().split(/\s/)[0] ?? '';
        if (!target) continue;
        const resultText: string =
          (field as unknown as { result?: { text?: string } }).result?.text ?? '';
        list.push({ target, resultText });
      }
      refFieldsByPara.set(paraId, list);
    }

    // Correlate each match to a field IN THE SAME PARAGRAPH by content overlap
    // (spec §5.1.4 requires the field to back THIS reference, not merely share the
    // paragraph). v1 heuristic: the field's displayed result text matches the
    // match's rawText (either contains the other, case-insensitively). When a
    // paragraph has exactly one REF field and one reference, that trivially holds;
    // for multi-reference paragraphs only the field whose result matches is backed.
    for (const match of matches) {
      const refFields = refFieldsByPara.get(match.paragraphId);
      if (!refFields || refFields.length === 0) {
        match.isFieldBacked = false;
        match.staticBadge = true;
        continue;
      }

      const raw = match.rawText.trim().toLowerCase();
      const backing = refFields.find(rf => {
        const res = rf.resultText.trim().toLowerCase();
        if (!res) return false;
        return res === raw || res.includes(raw) || raw.includes(res);
      });

      if (backing) {
        match.isFieldBacked = true;
        match.staticBadge = false;
        match.fieldId = backing.target;
        // Broken REF: target bookmark absent → Not Found, high confidence (spec §5.1.4 / A13).
        if (!validBookmarks.has(backing.target)) {
          match.classification = 'notFound';
          match.severity = 'error';
        }
      } else {
        // A REF field exists in the paragraph but none matches this reference → manual.
        match.isFieldBacked = false;
        match.staticBadge = true;
      }
    }

    return { fieldsApiAvailable: true };
  } catch {
    // ADR-0014 mandatory fallback: fields API unavailable.
    // Leave all matches as staticBadge=true, isFieldBacked=false.
    return { fieldsApiAvailable: false };
  }
}
