import PizZip from 'pizzip';
import type { LoadedDoc } from '../state/documents.js';

/**
 * One Word document with tracked changes becomes two: the text as it stood
 * before them (every change rejected) and after them (every change accepted).
 * Comparing those two is a comparison of the tracked changes themselves.
 *
 * Done as text surgery on the OOXML parts so it needs no Word and no new
 * dependency. Deliberately NOT handled, and disclosed to the user:
 *  - a tracked paragraph mark (two paragraphs merged or split) — the paragraph
 *    break stays on both sides;
 *  - formatting-only tracked changes (`rPrChange`/`pPrChange`) — both sides keep
 *    the current formatting. The engine hides formatting-only changes anyway.
 */

export class NoTrackedChangesError extends Error {
  constructor() {
    super(
      'This document has no tracked changes, so there is nothing to compare it with. ' +
        'Add a second document, or turn on Track Changes in Word and save the edits.',
    );
  }
}

const PARTS = /^word\/(document|footnotes|endnotes|header\d*|footer\d*)\.xml$/;
const TAGS = 'ins|del|moveFrom|moveTo';
// `(?=[\s/>])` so <w:insideH>, <w:delText> and range markers never match.
const SELF_CLOSING = new RegExp(`<w:(?:${TAGS})(?=[\\s/>])[^>]*/>`, 'g');

function elements(names: string): RegExp {
  return new RegExp(`<w:(${names})(?=[\\s>])[^>]*>[\\s\\S]*?</w:\\1>`, 'g');
}
function wrappers(names: string): RegExp {
  return new RegExp(`</?w:(?:${names})(?=[\\s>])[^>]*>`, 'g');
}

function dropFormatChanges(xml: string): string {
  return xml.replace(
    /<w:(rPrChange|pPrChange|sectPrChange|tblPrChange|tcPrChange|trPrChange|tblGridChange|numberingChange)(?=[\s>])[^>]*>[\s\S]*?<\/w:\1>/g,
    '',
  );
}

export function acceptAll(xml: string): string {
  let out = xml.replace(SELF_CLOSING, '');
  out = out.replace(elements('del|moveFrom'), '');
  out = out.replace(wrappers('ins|moveTo'), '');
  return dropFormatChanges(out);
}

export function rejectAll(xml: string): string {
  let out = xml.replace(SELF_CLOSING, '');
  out = out.replace(elements('ins|moveTo'), '');
  out = out.replace(wrappers('del|moveFrom'), '');
  out = out.replace(/<(\/?)w:delText(?=[\s>])/g, '<$1w:t').replace(/<(\/?)w:delInstrText(?=[\s>])/g, '<$1w:instrText');
  return dropFormatChanges(out);
}

export function hasTrackedChanges(xml: string): boolean {
  const stripped = xml.replace(SELF_CLOSING, '');
  return new RegExp(`<w:(?:${TAGS})(?=[\\s>])`).test(stripped);
}

function rewrite(zip: PizZip, fn: (xml: string) => string): Uint8Array {
  const copy = new PizZip(zip.generate({ type: 'uint8array' }));
  for (const name of Object.keys(copy.files)) {
    if (PARTS.test(name)) copy.file(name, fn(copy.file(name)!.asText()));
  }
  return copy.generate({ type: 'uint8array', compression: 'DEFLATE' });
}

const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

export const TRACKED_NOTE =
  'Compared from one document\'s tracked changes: the original is the text with every change rejected, ' +
  'the modified is the text with every change accepted. Tracked paragraph-mark and formatting-only changes ' +
  'are not separated.';

export async function splitTrackedChanges(
  doc: LoadedDoc,
): Promise<{ original: LoadedDoc; modified: LoadedDoc }> {
  if (doc.kind !== 'docx') throw new NoTrackedChangesError();
  const zip = new PizZip(await doc.file.arrayBuffer());
  const hasAny = Object.keys(zip.files).some(
    (n) => PARTS.test(n) && hasTrackedChanges(zip.file(n)!.asText()),
  );
  if (!hasAny) throw new NoTrackedChangesError();

  const base = doc.name.replace(/\.[^.]+$/, '');
  const make = (fn: (x: string) => string, label: string, tag: string): LoadedDoc => {
    const bytes = rewrite(zip, fn);
    const file = new File([bytes as BlobPart], `${base} (${label}).docx`, { type: DOCX });
    return {
      id: `${doc.id}-${tag}`,
      name: `${doc.name} (${label})`,
      byteLength: file.size,
      kind: 'docx',
      file,
      overCap: doc.overCap,
      notes: [TRACKED_NOTE],
      aiNote: 'tracked changes split into before/after',
    };
  };
  return {
    original: make(rejectAll, 'changes rejected', 'before'),
    modified: make(acceptAll, 'changes accepted', 'after'),
  };
}
