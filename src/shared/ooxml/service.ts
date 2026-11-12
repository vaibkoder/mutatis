import type { DocxBundle, OoxmlService, ReadOptions, WriteOptions } from './types.js';
import { PizZipBundle } from './bundle.js';

// Back-compat re-export: `PizZipBundle` moved to its own host-free module
// (bundle.ts) in 12f so host-neutral code (Compare, host-web) can import it
// without dragging this Office-coupled module into their bundle graph. Existing
// `@shared/ooxml/service.js` importers keep working unchanged.
export { PizZipBundle } from './bundle.js';

const DEFAULT_SLICE_SIZE = 65536;

/** Concatenate getFileAsync slices into a single Uint8Array. */
function readAllSlices(file: Office.File): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    const sliceCount = file.sliceCount;
    const slices: Uint8Array[] = new Array(sliceCount);
    let received = 0;

    for (let i = 0; i < sliceCount; i++) {
      file.getSliceAsync(i, result => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          file.closeAsync(() => { /* noop */ });
          reject(new Error(result.error?.message ?? 'getSliceAsync failed'));
          return;
        }
        const data = result.value.data as number[];
        slices[result.value.index] = new Uint8Array(data);
        received++;
        if (received === sliceCount) {
          file.closeAsync(() => { /* noop */ });
          const total = slices.reduce((sum, s) => sum + s.length, 0);
          const merged = new Uint8Array(total);
          let offset = 0;
          for (const s of slices) {
            merged.set(s, offset);
            offset += s.length;
          }
          resolve(merged);
        }
      });
    }
  });
}

/**
 * Detect a stray artifact paragraph introduced by insertFileFromBase64.
 * Returns 'trailing' if the last paragraph is empty, 'leading' if the first
 * paragraph is empty (and there is more than one paragraph), or null if no
 * artifact is found.
 *
 * Pure helper — no Office context required; designed to be unit-testable.
 */
export function detectArtifactParagraph(
  items: ReadonlyArray<{ text: string }>
): 'trailing' | 'leading' | null {
  if (items.length < 2) return null;
  if (items[items.length - 1]!.text === '') return 'trailing';
  if (items[0]!.text === '') return 'leading';
  return null;
}

/**
 * Map a getFileAsync failure into a human-readable Error. The raw `ApiNotFound`
 * code means the host (notably Word on the web) does not implement reading the
 * full compressed .docx — Document Repair needs the desktop app.
 */
export function toReadDocxError(code: string | number | undefined, message: string | undefined): Error {
  const haystack = `${code ?? ''} ${message ?? ''}`;
  if (/ApiNotFound/i.test(haystack)) {
    return new Error(
      'Document Repair needs the Word desktop app (Windows or Mac). Reading the ' +
        'full document isn’t supported in Word on the web — please reopen this ' +
        'document in desktop Word and scan again.',
    );
  }
  return new Error(message ?? 'getFileAsync failed');
}

export class OoxmlServiceImpl implements OoxmlService {
  async readDocx(opts: ReadOptions = {}): Promise<DocxBundle> {
    const sliceSize = opts.sliceSize ?? DEFAULT_SLICE_SIZE;
    const file = await new Promise<Office.File>((resolve, reject) => {
      try {
        Office.context.document.getFileAsync(
          Office.FileType.Compressed,
          { sliceSize },
          result => {
            if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
            else reject(toReadDocxError(result.error?.code, result.error?.message));
          }
        );
      } catch (e) {
        // Some hosts throw synchronously when the API is absent rather than
        // returning a failed AsyncResult.
        const code = (e as { code?: string })?.code;
        const message = e instanceof Error ? e.message : String(e);
        reject(toReadDocxError(code, message));
      }
    });
    const bytes = await readAllSlices(file);
    return new PizZipBundle(bytes);
  }

  async writeBack(bundle: DocxBundle, opts: WriteOptions = {}): Promise<void> {
    const mode = opts.mode ?? 'replace';
    if (!(bundle instanceof PizZipBundle)) {
      throw new Error('writeBack requires a PizZipBundle');
    }
    const base64 = bundle.toBase64();

    if (mode === 'newDocument') {
      await Word.run(async ctx => {
        const created = ctx.application.createDocument(base64);
        created.open();
        await ctx.sync();
      });
      return;
    }

    await Word.run(async ctx => {
      ctx.document.body.insertFileFromBase64(base64, Word.InsertLocation.replace);
      await ctx.sync();

      if (opts.onArtifactDetected !== undefined) {
        const paras = ctx.document.body.paragraphs;
        paras.load('items/text');
        await ctx.sync();

        const artifact = detectArtifactParagraph(paras.items);
        if (artifact !== null) {
          if (opts.onArtifactDetected === 'throw') {
            throw new Error('writeBack: stray artifact paragraph detected');
          }
          // 'fallbackToNewDoc': delete the stray paragraph in-place
          const strayIndex = artifact === 'trailing'
            ? paras.items.length - 1
            : 0;
          paras.items[strayIndex]!.delete();
          await ctx.sync();
        }
      }
    });
  }
}

/** Singleton OOXML service. */
export const ooxmlService: OoxmlService = new OoxmlServiceImpl();
