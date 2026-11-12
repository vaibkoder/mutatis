/**
 * Handing a batch's artefacts to the user (ADR-0037; spec §5.3).
 *
 * Two exits: one row's redline, or every completed row as a single archive.
 * The archive is assembled by `zip-store.ts` from the rows' own `Blob`s, so the
 * bytes never pass through the JS heap — which is the only reason "download
 * all" on a hundred-variant batch is safe at all.
 *
 * Host-free: the only browser APIs are `Blob`, `URL.createObjectURL` and an
 * anchor click, matching `host-web/web-host.ts`'s own download path. Nothing
 * leaves the machine (ADR-0001/0003).
 */
import type { ArtefactHandle, BatchRow } from './model.js';
import { writeStoreZip, type ZipEntry } from './zip-store.js';

/** Trigger a browser download of a Blob via a transient anchor. */
export function triggerDownload(blob: Blob, fileName: string): void {
  if (typeof document === 'undefined' || typeof URL.createObjectURL !== 'function') return;
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

/** The redline artefact of a row, if it has one retained. */
export function redlineOf(row: BatchRow): ArtefactHandle | undefined {
  return row.artefacts.find(a => a.kind === 'redlineDocx');
}

/**
 * Re-run one pair and produce its redline.
 *
 * Supplied by the driver, which is the only layer that still holds the `File`
 * handles. This is what makes the artefact ceiling a *degradation* rather than
 * a loss: past it, a row keeps no `Blob`, but the redline is still obtainable —
 * it costs one re-run instead of memory (spec §4.3 step 6).
 */
export type ArtefactProducer = (row: BatchRow) => Promise<ArtefactHandle | null>;

/**
 * Whether a row has a redline the user can obtain.
 *
 * A row marked `artefactOnDemand` compared successfully but kept no `Blob`,
 * because the batch had reached its artefact ceiling. It is still downloadable:
 * the bytes are produced by re-running that one pair. Reporting it as
 * undownloadable would turn a disclosed memory trade-off into a silently lost
 * result, which is the thing this feature is not allowed to do.
 */
export function isDownloadable(row: BatchRow): boolean {
  return row.status === 'done' && (redlineOf(row) !== undefined || row.artefactOnDemand);
}

/**
 * Resolve a row's redline: the retained `Blob` if it has one, otherwise a
 * re-run. Returns `null` when there is nothing to produce and no producer.
 */
async function resolveArtefact(
  row: BatchRow,
  produce?: ArtefactProducer,
): Promise<ArtefactHandle | null> {
  const retained = redlineOf(row);
  if (retained) return retained;
  if (row.status === 'done' && row.artefactOnDemand && produce) return produce(row);
  return null;
}

/** Download one row's redline, re-running the pair if it was not retained. */
export async function downloadArtefact(
  row: BatchRow,
  produce?: ArtefactProducer,
): Promise<void> {
  const artefact = await resolveArtefact(row, produce);
  if (!artefact) return;
  triggerDownload(artefact.blob, artefact.fileName);
}

/**
 * Download every downloadable row as one archive.
 *
 * Entries are fed to the writer lazily, one at a time, so peak memory stays at
 * a single artefact however many rows there are — and that laziness is also
 * what lets an on-demand row be re-run in its turn rather than all of them up
 * front, which would reintroduce the very memory ceiling that made them
 * on-demand in the first place.
 *
 * A row that yields nothing is skipped rather than written as an empty entry:
 * an empty `.docx` in the archive would read as a redline with no changes,
 * which is a different and wrong claim. The batch index is where its state is
 * reported.
 */
export async function downloadRowArchive(
  rows: readonly BatchRow[],
  templateLabel: string,
  produce?: ArtefactProducer,
): Promise<void> {
  const downloadable = rows.filter(isDownloadable);
  if (downloadable.length === 0) return;

  async function* entries(): AsyncGenerator<ZipEntry> {
    for (const row of downloadable) {
      const artefact = await resolveArtefact(row, produce);
      if (artefact) yield { name: artefact.fileName, blob: artefact.blob };
    }
  }

  const archive = await writeStoreZip(entries());
  const stem = templateLabel.replace(/\.docx$/i, '');
  triggerDownload(archive, `${stem} — redlines.zip`);
}
