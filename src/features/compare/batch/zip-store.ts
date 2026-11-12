/**
 * A minimal STORE-only ZIP writer for "download all" (ADR-0037; spec §5.3).
 *
 * Two reasons this is hand-rolled rather than delegated to PizZip, which the
 * project already ships:
 *
 *  1. **Memory.** PizZip assembles the whole archive in the JS heap. A batch's
 *     artefacts live in browser-managed `Blob`s precisely so they are NOT in the
 *     heap (spec §4.3); routing them through PizZip would undo that at the last
 *     step. Here each entry's bytes are read once to checksum it, then the
 *     ORIGINAL `Blob` is handed to the output `Blob` untouched, so peak heap is
 *     one entry — not the whole archive.
 *  2. **STORE is correct anyway.** A `.docx` is already a deflated ZIP and a
 *     rendered PDF is already compressed; deflating them a second time costs CPU
 *     for roughly nothing.
 *
 * Deliberately not implemented: Zip64. The batch's retention ceiling is well
 * under 4 GiB, and a silently-truncated 32-bit size field would produce an
 * archive that looks fine and is corrupt — so an oversize archive throws instead.
 *
 * Pure — no Office JS, no DOM beyond `Blob`, no network.
 */

/** One file to place in the archive. */
export interface ZipEntry {
  /** Entry name. Sanitised and de-duplicated by the writer if it needs to be. */
  name: string;
  blob: Blob;
}

/** Beyond this a ZIP needs Zip64, which this writer does not emit. */
const ZIP32_MAX = 0xffff_ffff;

/** Thrown rather than emit an archive whose 32-bit size fields have wrapped. */
export class ZipTooLargeError extends Error {
  constructor(readonly actual: number) {
    super(
      `The archive would exceed the 4 GiB single-archive limit (${actual} bytes). ` +
        `Download the redlines in smaller groups.`,
    );
    this.name = 'ZipTooLargeError';
  }
}

// --- CRC-32 (IEEE 802.3), table-driven ------------------------------------

const CRC_TABLE = ((): Uint32Array => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb8_8320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

/** CRC-32 of a byte array. Exported for tests. */
export function crc32(bytes: Uint8Array): number {
  let c = 0xffff_ffff;
  for (let i = 0; i < bytes.length; i += 1) {
    c = (CRC_TABLE[(c ^ (bytes[i] as number)) & 0xff] as number) ^ (c >>> 8);
  }
  return (c ^ 0xffff_ffff) >>> 0;
}

// --- little-endian scratch writer -----------------------------------------

class LeWriter {
  private readonly bytes: number[] = [];

  u16(v: number): this {
    this.bytes.push(v & 0xff, (v >>> 8) & 0xff);
    return this;
  }

  u32(v: number): this {
    this.bytes.push(v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff);
    return this;
  }

  raw(v: Uint8Array): this {
    for (const b of v) this.bytes.push(b);
    return this;
  }

  /**
   * `Uint8Array<ArrayBuffer>` rather than plain `Uint8Array`: the DOM's
   * `BlobPart` will not accept the `ArrayBufferLike` default, since that admits
   * a `SharedArrayBuffer`.
   */
  done(): Uint8Array<ArrayBuffer> {
    return new Uint8Array(this.bytes);
  }
}

/**
 * A fixed DOS timestamp (1 Jan 1980, 00:00) rather than the wall clock.
 *
 * The same batch downloaded twice should produce byte-identical archives — a
 * real timestamp would make every archive differ from every other and make the
 * writer untestable byte-for-byte. The entries carry no date of their own
 * either way, so nothing is lost.
 */
const DOS_TIME = 0;
const DOS_DATE = 0x0021; // 1980-01-01

const UTF8_FLAG = 0x0800;

/**
 * Characters that must not appear in an entry name.
 *
 * Path separators (a name containing one would place the entry in a
 * subdirectory, or escape the archive), the characters Windows rejects
 * outright, and the control range. Built from escapes rather than written as a
 * literal class so the source file carries no control bytes of its own.
 */
const UNSAFE_ENTRY_CHARS = new RegExp('[\\\\\\/:*?"<>|\\u0000-\\u001f\\u007f]', 'g');

/**
 * Make an entry name safe for a ZIP and for every desktop filesystem, and
 * de-duplicate it against the names already used.
 *
 * Deliberately conservative: spaces and hyphens are left alone.
 * `Project Falcon SPA-v7.docx` is a perfectly good filename, and mangling it
 * makes a downloaded redline harder to match back to the document it came
 * from — which is the one job the name has.
 *
 * Duplicate names are not a cosmetic problem: most extractors silently keep
 * only one entry of a duplicated name, so a batch containing two variants both
 * called `SPA.docx` would quietly lose a redline.
 */
export function safeEntryName(name: string, used: Set<string>): string {
  const cleaned =
    name.replace(UNSAFE_ENTRY_CHARS, '_').replace(/^\.+/, '_').trim() || 'document';

  if (!used.has(cleaned)) {
    used.add(cleaned);
    return cleaned;
  }

  const dot = cleaned.lastIndexOf('.');
  const stem = dot > 0 ? cleaned.slice(0, dot) : cleaned;
  const ext = dot > 0 ? cleaned.slice(dot) : '';
  for (let n = 2; ; n += 1) {
    const candidate = `${stem} (${n})${ext}`;
    if (!used.has(candidate)) {
      used.add(candidate);
      return candidate;
    }
  }
}

/**
 * Write a STORE-only ZIP.
 *
 * Entries are consumed one at a time — each is read to checksum it and then
 * released — so peak heap is one entry regardless of how many there are. The
 * source `Blob`s are placed into the result unchanged, which is what keeps the
 * archive's bytes out of the JS heap.
 *
 * Accepts an async iterable so a caller can produce an entry lazily: a row
 * whose artefact is on-demand re-runs its pair as the archive reaches it,
 * rather than everything being materialised up front.
 */
export async function writeStoreZip(
  entries: Iterable<ZipEntry> | AsyncIterable<ZipEntry>,
): Promise<Blob> {
  const parts: BlobPart[] = [];
  const central: Uint8Array<ArrayBuffer>[] = [];
  const used = new Set<string>();
  const encoder = new TextEncoder();
  let offset = 0;
  let count = 0;

  for await (const entry of entries as AsyncIterable<ZipEntry>) {
    const name: Uint8Array<ArrayBuffer> = encoder.encode(safeEntryName(entry.name, used));
    const bytes = new Uint8Array(await entry.blob.arrayBuffer());
    const size = bytes.length;
    const sum = crc32(bytes);

    const local = new LeWriter()
      .u32(0x0403_4b50)
      .u16(20)
      .u16(UTF8_FLAG)
      .u16(0) // STORE
      .u16(DOS_TIME)
      .u16(DOS_DATE)
      .u32(sum)
      .u32(size)
      .u32(size)
      .u16(name.length)
      .u16(0)
      .raw(name)
      .done();

    central.push(
      new LeWriter()
        .u32(0x0201_4b50)
        .u16(20)
        .u16(20)
        .u16(UTF8_FLAG)
        .u16(0)
        .u16(DOS_TIME)
        .u16(DOS_DATE)
        .u32(sum)
        .u32(size)
        .u32(size)
        .u16(name.length)
        .u16(0)
        .u16(0)
        .u16(0)
        .u16(0)
        .u32(0)
        .u32(offset)
        .raw(name)
        .done(),
    );

    parts.push(local, entry.blob);
    offset += local.length + size;
    count += 1;
    if (offset > ZIP32_MAX) throw new ZipTooLargeError(offset);
  }

  const centralSize = central.reduce((n, c) => n + c.length, 0);
  const eocd = new LeWriter()
    .u32(0x0605_4b50)
    .u16(0)
    .u16(0)
    .u16(count)
    .u16(count)
    .u32(centralSize)
    .u32(offset)
    .u16(0)
    .done();

  if (offset + centralSize + eocd.length > ZIP32_MAX) {
    throw new ZipTooLargeError(offset + centralSize + eocd.length);
  }

  return new Blob([...parts, ...central, eocd], { type: 'application/zip' });
}
