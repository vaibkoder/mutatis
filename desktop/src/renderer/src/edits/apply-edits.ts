import { PizZipBundle } from '@shared/ooxml/bundle.js';

/**
 * Applying an AI's proposed edits to a `.docx`.
 *
 * THE RULE THIS FILE EXISTS TO ENFORCE: an edit is applied only when its `find`
 * text occurs EXACTLY ONCE in the document. Not zero times, not twice. A model
 * that quotes text slightly wrong, or quotes a phrase that appears in both a
 * clause and its schedule, gets a refusal with a reason — never a guess at which
 * one it meant. In a legal document the cost of editing the wrong clause is not
 * symmetrical with the cost of asking again.
 *
 * Nothing here writes revision markup. The edited document is a CLEAN document;
 * tracked changes come from running the comparison engine between the original
 * and this edited copy, which is what guarantees that every AI edit reaches the
 * user as a tracked change they review — the same machinery, with no second
 * code path to get wrong.
 */

export interface ProposedEdit {
  /** Exact text to replace. Must occur exactly once in the document. */
  readonly find: string;
  /** Replacement text. Empty string deletes. */
  readonly replace: string;
  /** Why, in the model's words. Shown to the user; never applied to the document. */
  readonly reason?: string;
}

export interface AppliedEdit {
  readonly find: string;
  readonly replace: string;
  readonly reason?: string;
}

export interface RefusedEdit {
  readonly find: string;
  readonly replace: string;
  readonly reason?: string;
  readonly refusal: string;
}

export interface EditResult {
  /** The edited document. Clean — no revision markup. */
  readonly bytes: Uint8Array;
  readonly applied: readonly AppliedEdit[];
  readonly refused: readonly RefusedEdit[];
}

const PARA_RE = /<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>|<w:p(?:\s[^>]*)?\/>/g;
const TEXT_RE = /<w:t(\s[^>]*)?>([\s\S]*?)<\/w:t>/g;

export function applyEdits(
  documentBytes: ArrayBuffer | Uint8Array,
  edits: readonly ProposedEdit[],
): EditResult {
  const bundle = new PizZipBundle(documentBytes);
  const xml = bundle.part('word/document.xml');
  if (xml === undefined) {
    return {
      bytes: new PizZipBundle(documentBytes).toUint8Array(),
      applied: [],
      refused: edits.map((edit) => ({
        ...edit,
        refusal: 'This file has no word/document.xml and cannot be edited.',
      })),
    };
  }

  let current = xml;
  const applied: AppliedEdit[] = [];
  const refused: RefusedEdit[] = [];

  for (const edit of edits) {
    const outcome = applyOne(current, edit);
    if (outcome.ok) {
      current = outcome.xml;
      applied.push({ find: edit.find, replace: edit.replace, ...reasonOf(edit) });
    } else {
      refused.push({ find: edit.find, replace: edit.replace, ...reasonOf(edit), refusal: outcome.refusal });
    }
  }

  if (applied.length > 0) bundle.setPart('word/document.xml', current);
  return { bytes: bundle.toUint8Array(), applied, refused };
}

function reasonOf(edit: ProposedEdit): { reason?: string } {
  return edit.reason !== undefined ? { reason: edit.reason } : {};
}

type ApplyOutcome = { ok: true; xml: string } | { ok: false; refusal: string };

function applyOne(xml: string, edit: ProposedEdit): ApplyOutcome {
  // NOT trimmed. Leading and trailing spaces are part of what the model asked to
  // replace — trimming them here turns "delete ' in full'" into "delete 'in
  // full'" and leaves a double space behind.
  const find = edit.find;
  if (find.trim() === '') return { ok: false, refusal: 'The text to replace was empty.' };

  // Paragraphs first, because a replacement must stay inside one paragraph —
  // text spanning a paragraph boundary is a structural edit, not a wording one,
  // and silently flattening two paragraphs into one is not a change anyone asked
  // for.
  const paragraphs = [...xml.matchAll(PARA_RE)];
  const hits: { index: number; start: number; end: number }[] = [];

  paragraphs.forEach((match, index) => {
    const para = match[0];
    const runs = textRuns(para);
    const plain = runs.map((run) => run.text).join('');
    let from = 0;
    for (;;) {
      const at = plain.indexOf(find, from);
      if (at === -1) break;
      hits.push({ index, start: at, end: at + find.length });
      from = at + 1;
    }
  });

  if (hits.length === 0) {
    return {
      ok: false,
      refusal:
        'That exact text does not appear in the document. It may have been paraphrased, or the ' +
        'spacing or punctuation may differ.',
    };
  }
  if (hits.length > 1) {
    return {
      ok: false,
      refusal:
        `That text appears ${hits.length} times. Quote more of the surrounding sentence so there ` +
        'is only one place it could mean.',
    };
  }

  const hit = hits[0]!;
  const match = paragraphs[hit.index]!;
  const rewritten = rewriteParagraph(match[0], hit.start, hit.end, edit.replace);
  if (rewritten === null) {
    return { ok: false, refusal: 'That text could not be located inside the paragraph markup.' };
  }

  const at = match.index;
  return { ok: true, xml: xml.slice(0, at) + rewritten + xml.slice(at + match[0].length) };
}

interface TextRun {
  /** Offset of the whole `<w:t …>…</w:t>` element within the paragraph XML. */
  readonly at: number;
  readonly length: number;
  readonly attrs: string;
  /** Decoded text content. */
  readonly text: string;
}

function textRuns(paragraphXml: string): TextRun[] {
  const runs: TextRun[] = [];
  for (const match of paragraphXml.matchAll(TEXT_RE)) {
    runs.push({
      at: match.index,
      length: match[0].length,
      attrs: match[1] ?? '',
      text: decode(match[2] ?? ''),
    });
  }
  return runs;
}

/**
 * Put the replacement into the FIRST run the match touches and empty the rest.
 *
 * Word splits a sentence across runs for reasons that have nothing to do with
 * meaning — a spellcheck boundary, a stray language tag — so a replacement
 * frequently spans several. Collapsing into the first run means the new text
 * inherits that run's formatting. That is a real limitation and it is the
 * conventional one; it is disclosed to the user rather than papered over.
 */
function rewriteParagraph(
  paragraphXml: string,
  start: number,
  end: number,
  replacement: string,
): string | null {
  const runs = textRuns(paragraphXml);
  if (runs.length === 0) return null;

  let cursor = 0;
  const pieces: { run: TextRun; text: string }[] = [];
  let touched = false;

  for (const run of runs) {
    const runStart = cursor;
    const runEnd = cursor + run.text.length;
    cursor = runEnd;

    if (runEnd <= start || runStart >= end) {
      pieces.push({ run, text: run.text });
      continue;
    }

    const prefix = run.text.slice(0, Math.max(0, start - runStart));
    const suffix = run.text.slice(Math.max(0, end - runStart));
    if (!touched) {
      touched = true;
      pieces.push({ run, text: prefix + replacement + suffix });
    } else {
      pieces.push({ run, text: prefix + suffix });
    }
  }

  if (!touched) return null;

  // Rebuild back-to-front so earlier offsets stay valid.
  let out = paragraphXml;
  for (let i = pieces.length - 1; i >= 0; i -= 1) {
    const piece = pieces[i]!;
    if (piece.text === piece.run.text) continue;
    const element = `<w:t${preserveSpace(piece.run.attrs, piece.text)}>${encode(piece.text)}</w:t>`;
    out = out.slice(0, piece.run.at) + element + out.slice(piece.run.at + piece.run.length);
  }
  return out;
}

/**
 * `xml:space="preserve"` matters: without it Word collapses leading and trailing
 * whitespace, so a replacement ending in a space silently loses it.
 */
function preserveSpace(attrs: string, text: string): string {
  if (attrs.includes('xml:space')) return attrs;
  const needed = text !== text.trim() || text === '';
  return needed ? `${attrs} xml:space="preserve"` : attrs;
}

function decode(value: string): string {
  return value
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

function encode(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}
