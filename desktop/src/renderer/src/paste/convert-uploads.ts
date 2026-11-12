import type { LoadedDoc, RejectedFile } from '../state/documents.js';
import { openPdf } from '../pdf/open-pdf.js';
import { explainPdfError, readPdf } from '../pdf/read-pdf.js';
import { markdownToBlocks } from './markdown-to-blocks.js';
import { plainToBlocks } from './plain-to-blocks.js';
import { toDoc } from './pasted.js';

/**
 * The engine reads only OOXML. Uploaded .txt, .md and .pdf files are turned into
 * a .docx as they are added, with the converters the paste view uses, so every
 * kind of input rides the one compare path. The document keeps the uploaded
 * file's name and id; `notes` carries what the user should know about how it was
 * read (shown above the redline).
 *
 * A file that cannot be read is NOT added: it comes back in `failed` with the
 * reason, so the drop zone says so immediately instead of the comparison failing
 * later with a message about zip files.
 */
export async function convertUploads(
  docs: readonly LoadedDoc[],
  read: typeof readPdf = readPdf,
): Promise<{ docs: LoadedDoc[]; failed: RejectedFile[] }> {
  const out: LoadedDoc[] = [];
  const failed: RejectedFile[] = [];

  for (const doc of docs) {
    if (doc.kind === 'docx') {
      out.push(doc);
      continue;
    }
    try {
      if (doc.kind === 'pdf') {
        // PDF.js detaches the buffer it is given, so it gets its own copy.
        const bytes = new Uint8Array(await doc.file.arrayBuffer());
        const { blocks, notes } = await read(bytes, openPdf);
        out.push(withIdentity(toDoc(blocks, doc.name.replace(/\.[^.]+$/, '')), doc, notes));
      } else {
        const text = await doc.file.text();
        const blocks = /\.md$/i.test(doc.name) ? markdownToBlocks(text) : plainToBlocks(text);
        out.push(withIdentity(toDoc(blocks, doc.name.replace(/\.[^.]+$/, '')), doc, []));
      }
    } catch (error) {
      failed.push({ name: doc.name, reason: explainPdfError(error) });
    }
  }
  return { docs: out, failed };
}

function withIdentity(converted: LoadedDoc, original: LoadedDoc, notes: readonly string[]): LoadedDoc {
  return {
    ...converted,
    id: original.id,
    name: original.name,
    kind: 'docx',
    ...(notes.length > 0 ? { notes, aiNote: 'PDF: text rebuilt from page layout' } : {}),
  };
}
