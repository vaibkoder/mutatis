import { exceedsDocxByteCap } from '@features/compare/batch/limits.js';
import type { LoadedDoc } from '../state/documents.js';
import { htmlToBlocks } from './html-to-blocks.js';
import { markdownToBlocks } from './markdown-to-blocks.js';
import { plainToBlocks } from './plain-to-blocks.js';
import { blocksToDocx } from './blocks-to-docx.js';
import { hasContent, type Block, type PasteFormat } from './model.js';

/** One side of the paste view: what arrived, and how it is being read. */
export interface PasteSide {
  /** The plain-text version of the paste, and what the user types or edits. */
  readonly text: string;
  /** Clipboard HTML, kept only while it carries formatting worth keeping. */
  readonly html: string | null;
  readonly format: PasteFormat;
}

export const emptyPasteSide: PasteSide = { text: '', html: null, format: 'plain' };

export function sideBlocks(side: PasteSide): Block[] {
  if (side.format === 'formatted' && side.html !== null) return htmlToBlocks(side.html);
  if (side.format === 'markdown') return markdownToBlocks(side.text);
  return plainToBlocks(side.text);
}

let counter = 0;

/** Turn the two sides into documents the ordinary compare path can run. */
export function buildPastedPair(
  original: PasteSide,
  modified: PasteSide,
): { original: LoadedDoc; modified: LoadedDoc } | null {
  const a = sideBlocks(original);
  const b = sideBlocks(modified);
  if (!hasContent(a) || !hasContent(b)) return null;
  const n = ++counter;
  return {
    original: toDoc(a, `Pasted original ${n}`),
    modified: toDoc(b, `Pasted modified ${n}`),
  };
}

export function toDoc(blocks: readonly Block[], name: string): LoadedDoc {
  const bytes = blocksToDocx(blocks);
  const file = new File([bytes as BlobPart], `${name}.docx`, {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  return {
    id: `pasted-${name.replace(/\s+/g, '-').toLowerCase()}`,
    name,
    byteLength: file.size,
    kind: 'docx',
    file,
    overCap: exceedsDocxByteCap(file.size),
  };
}
