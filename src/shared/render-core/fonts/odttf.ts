/**
 * Embedded-font (.odttf) de-obfuscation — host-free, pure (brief §4.2).
 *
 * Word obfuscates embedded fonts by XOR-ing the FIRST 32 bytes of the otherwise
 * untouched font file with the 16-byte fontKey GUID, applied REVERSED and cycled
 * twice. The fontKey lives in `word/fontTable.xml`
 * (`<w:embedRegular r:id=… w:fontKey="{GUID}"/>`); the obfuscated bytes live in
 * `word/fonts/fontNNN.odttf`. Re-embedding a subset of the user's own document
 * font into the output PDF is the documented intent of the embedding bits.
 */

/** Parse a `{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}` GUID into its 16 raw bytes. */
export function parseFontKey(guid: string): Uint8Array {
  const hex = guid.replace(/[{}\-]/g, '');
  if (hex.length !== 32 || /[^0-9a-fA-F]/.test(hex)) {
    throw new Error(`odttf: invalid fontKey GUID "${guid}"`);
  }
  const out = new Uint8Array(16);
  for (let i = 0; i < 16; i++) out[i] = parseInt(hex.slice(i * 2, i * 2 + 2), 16);
  return out;
}

/**
 * De-obfuscate `.odttf` bytes in place-safe fashion (returns a new Uint8Array).
 * XORs the first 32 bytes with the reversed key (16 bytes), cycled twice.
 */
export function deobfuscateOdttf(odttf: Uint8Array, fontKey: string | Uint8Array): Uint8Array {
  const key = typeof fontKey === 'string' ? parseFontKey(fontKey) : fontKey;
  if (key.length !== 16) throw new Error('odttf: fontKey must be 16 bytes');
  const reversed = Uint8Array.from(key).reverse();
  const out = Uint8Array.from(odttf);
  const n = Math.min(32, out.length);
  for (let i = 0; i < n; i++) {
    // first 16 bytes use reversed[0..15], next 16 cycle again
    out[i] = (out[i] as number) ^ (reversed[i % 16] as number);
  }
  return out;
}
