/**
 * FNV-1a hashing for Compare paragraph records.
 * Spec: spec/features/08-compare.md §4.1 (ParaRecord.hash), §5.1.
 *
 * Pure: no side effects, no Date/Math.random.
 */

const FNV_OFFSET_BASIS_32 = 0x811c9dc5; // 2166136261
const FNV_PRIME_32 = 0x01000193; // 16777619

/**
 * FNV-1a 32-bit hash of a UTF-8 string. Returns an 8-char lowercase hex string.
 *
 * Hashing is performed over the UTF-8 byte sequence (not UTF-16 code units) so
 * the hash is stable across surrogate-pair handling and matches a canonical
 * FNV-1a of the encoded text.
 */
export function fnv1a(text: string): string {
  const bytes = utf8Bytes(text);
  let hash = FNV_OFFSET_BASIS_32;
  for (let i = 0; i < bytes.length; i++) {
    hash ^= bytes[i] as number;
    // Multiply by the FNV prime modulo 2^32 using Math.imul to stay in 32 bits.
    hash = Math.imul(hash, FNV_PRIME_32);
  }
  // Coerce to unsigned 32-bit, then to fixed-width hex.
  return (hash >>> 0).toString(16).padStart(8, '0');
}

/** Encode a string to its UTF-8 byte sequence. */
function utf8Bytes(text: string): number[] {
  const out: number[] = [];
  for (let i = 0; i < text.length; i++) {
    let codePoint = text.charCodeAt(i);

    // Combine surrogate pairs into a single code point.
    if (codePoint >= 0xd800 && codePoint <= 0xdbff && i + 1 < text.length) {
      const low = text.charCodeAt(i + 1);
      if (low >= 0xdc00 && low <= 0xdfff) {
        codePoint = 0x10000 + ((codePoint - 0xd800) << 10) + (low - 0xdc00);
        i++;
      }
    }

    if (codePoint < 0x80) {
      out.push(codePoint);
    } else if (codePoint < 0x800) {
      out.push(0xc0 | (codePoint >> 6), 0x80 | (codePoint & 0x3f));
    } else if (codePoint < 0x10000) {
      out.push(
        0xe0 | (codePoint >> 12),
        0x80 | ((codePoint >> 6) & 0x3f),
        0x80 | (codePoint & 0x3f),
      );
    } else {
      out.push(
        0xf0 | (codePoint >> 18),
        0x80 | ((codePoint >> 12) & 0x3f),
        0x80 | ((codePoint >> 6) & 0x3f),
        0x80 | (codePoint & 0x3f),
      );
    }
  }
  return out;
}
