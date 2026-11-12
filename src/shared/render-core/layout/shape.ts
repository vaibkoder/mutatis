/**
 * Shaping / advance-width measurement via HarfBuzz (harfbuzzjs, MIT).
 *
 * Word Latin defaults: kerning OFF, ligatures OFF, hyphenation OFF — so width
 * is the sum of `hmtx` advances, deterministic across engines (Spike A). The
 * layout core depends only on the `Shaper` INTERFACE, so the breaker/paginator
 * can be unit-tested with a fake measurer and the real HarfBuzz path is tested
 * in isolation.
 *
 * Wasm delivery (Phase 2, ADR-0024 §2.3 "vendored same-origin, offline"):
 * harfbuzzjs (v1.2.1) auto-initialises on import — `index.mjs` runs a top-level
 * `init(await createHarfBuzz())` with NO args, and neither `init` nor the
 * `moduleArg.wasmBinary` seam is reachable through the package's public API, so
 * the wasm cannot be byte-injected the way the OFL fonts are (`?bytes`). Instead
 * the emscripten factory resolves `new URL('harfbuzz.wasm', import.meta.url)`,
 * which Vite emits as a **same-origin bundled asset** and loads offline — the
 * SAME mechanism the project already ships for pdfjs (`pdfjsAssets()` +
 * `scripts/check-pdf-offline.mjs`, which gates off-origin/CDN refs, not `fetch`
 * per se). No external egress; no document content leaves the machine. In
 * Node/vitest the factory reads the wasm via `fs` (the engine tests exercise the
 * real shaper). The integrated browser-worker wasm load is verified at the build
 * + cross-engine steps (Phase 2 D1/D3).
 */
import type { ResolvedFace } from '../fonts/registry.js';

/** A measurer: advance width of `text` for `face` at `sizePt`, in points. */
export interface Shaper {
  advanceWidthPt(face: ResolvedFace, text: string, sizePt: number): number;
}

// Lazy HarfBuzz handle so importing this module doesn't force wasm init until
// the first real shape (keeps non-shaping unit tests fast).
type Hb = typeof import('harfbuzzjs');
let hbMod: Hb | null = null;
async function getHb(): Promise<Hb> {
  if (!hbMod) hbMod = await import('harfbuzzjs');
  return hbMod;
}

interface FaceHandle {
  font: unknown; // hb.Font (typed via the dynamic import; shaped through `as never`)
  upem: number;
}

/** Copy into a fresh ArrayBuffer (hb.Blob's typed ctor wants ArrayBuffer, not Uint8Array). */
function toArrayBuffer(u8: Uint8Array): ArrayBuffer {
  const ab = new ArrayBuffer(u8.byteLength);
  new Uint8Array(ab).set(u8);
  return ab;
}

/**
 * A HarfBuzz-backed shaper. Caches one hb.Font per face byte-buffer. Create with
 * `await createHarfBuzzShaper()`.
 */
export async function createHarfBuzzShaper(): Promise<Shaper> {
  const hb = await getHb();
  const off = ['-kern', '-liga', '-clig', '-rlig', '-calt']
    .map((s) => hb.Feature.fromString(s))
    .filter((f): f is NonNullable<typeof f> => f != null);

  const cache = new WeakMap<Uint8Array, FaceHandle>();
  function handle(face: ResolvedFace): FaceHandle {
    const cached = cache.get(face.bytes);
    if (cached) return cached;
    const blob = new hb.Blob(toArrayBuffer(face.bytes));
    const hbFace = new hb.Face(blob, 0);
    const font = new hb.Font(hbFace);
    const h: FaceHandle = { font, upem: hbFace.upem };
    cache.set(face.bytes, h);
    return h;
  }

  return {
    advanceWidthPt(face: ResolvedFace, text: string, sizePt: number): number {
      if (text.length === 0) return 0;
      const h = handle(face);
      const buf = new hb.Buffer();
      buf.addText(text);
      buf.guessSegmentProperties();
      hb.shape(h.font as never, buf, off);
      const glyphs = buf.getGlyphInfosAndPositions();
      let widthFu = 0;
      for (const g of glyphs) widthFu += g.xAdvance ?? 0;
      if (typeof (buf as unknown as { destroy?: () => void }).destroy === 'function') {
        (buf as unknown as { destroy: () => void }).destroy();
      }
      // Rounding policy (Spike A): sum in font units, convert to pt ONCE.
      return (widthFu / h.upem) * sizePt;
    },
  };
}
