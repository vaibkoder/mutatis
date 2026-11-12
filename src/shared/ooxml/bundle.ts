/**
 * Host-free OOXML package container.
 *
 * `PizZipBundle` is the pure, client-side `.docx` (zip) read/modify/serialise
 * primitive — it touches PizZip only, never Office JS. It lives in its OWN module
 * (split out of `service.ts` in 12f) so the host-NEUTRAL consumers (Compare, the
 * web host adapter) can import it without dragging the Office-coupled
 * `OoxmlServiceImpl` (`Word.run` / `getFileAsync` / `insertFileFromBase64`) into
 * their bundle graph. This is what keeps the standalone web chunk Office-free
 * even with Compare in the import graph (confidentiality invariant; ADR-0001/0022).
 *
 * `service.ts` re-exports this for back-compat; new host-neutral code should
 * import `PizZipBundle` from here (`@shared/ooxml/bundle.js`) directly.
 */
import PizZip from 'pizzip';
import type { DocxBundle } from './types.js';

/**
 * DocxBundle backed by PizZip. Standardise on PizZip across the bundle for the
 * synchronous OOXML loop (spec/ARCHITECTURE.md §Technology Stack).
 */
export class PizZipBundle implements DocxBundle {
  private zip: PizZip;

  constructor(data?: Uint8Array | ArrayBuffer) {
    this.zip = data ? new PizZip(data) : new PizZip();
  }

  part(path: string): string | undefined {
    const file = this.zip.file(path);
    return file ? file.asText() : undefined;
  }

  /** Raw bytes of a part (binary-safe), e.g. an embedded `.odttf` font. */
  partBytes(path: string): Uint8Array | undefined {
    const file = this.zip.file(path);
    return file ? file.asUint8Array() : undefined;
  }

  setPart(path: string, xml: string): void {
    this.zip.file(path, xml);
  }

  /** Write a binary part (e.g. an image a copied part's relationships target). */
  setPartBytes(path: string, bytes: Uint8Array): void {
    this.zip.file(path, bytes);
  }

  hasPart(path: string): boolean {
    return this.zip.file(path) !== null;
  }

  listParts(): string[] {
    return Object.keys(this.zip.files).filter(name => !this.zip.files[name]?.dir);
  }

  removePart(path: string): void {
    this.zip.remove(path);
  }

  stripAttributes(partPath: string, attrNames: string[]): void {
    const xml = this.part(partPath);
    if (xml === undefined) return;
    let stripped = xml;
    for (const attr of attrNames) {
      // Targeted regex preserves byte fidelity better than full re-parse for simple strips.
      // Matches: w:rsid="..." or w:rsidR="..." etc. (attr may end with * wildcard)
      const escaped = attr.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace('\\*', '\\w*');
      const re = new RegExp(`\\s+${escaped}\\s*=\\s*"[^"]*"`, 'g');
      stripped = stripped.replace(re, '');
    }
    this.setPart(partPath, stripped);
  }

  /** Serialise the package to bytes. */
  toUint8Array(): Uint8Array {
    return this.zip.generate({ type: 'uint8array', compression: 'DEFLATE' });
  }

  /** Serialise to base64 for insertFileFromBase64. */
  toBase64(): string {
    return this.zip.generate({ type: 'base64', compression: 'DEFLATE' });
  }
}
