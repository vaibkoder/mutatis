/**
 * Vendored OFL clone faces (RenderCore, feature 16 / ADR-0024, Phase 2 A4).
 *
 * Turns the 16 build-time-embedded clone fonts (via the `?bytes` plugin — zero
 * runtime egress) into the `FaceInput[]` the engine consumes, and registers the
 * clones as `FontFace`s for the preview's *display* (metrics still come from the
 * clone bytes via HarfBuzz, so positions stay valid regardless of which font the
 * browser paints with).
 *
 *   Times New Roman → Liberation Serif (Tinos bytes)
 *   Arial           → Liberation Sans
 *   Calibri         → Carlito
 *   Cambria         → Caladea
 *
 * The clone-family *names* (with spaces) match `registry.ts` CLONE_MAP; the
 * on-disk files use space-free prefixes for import specifiers. This module is
 * pulled only behind the lazy render path, so non-render flows never load the
 * ~1.6 MB of font bytes (acceptance 10).
 */
import type { FaceInput } from '../fonts/registry.js';

// Liberation Sans (Arial clone).
import sansR from '@fonts/LiberationSans-Regular.ttf?bytes';
import sansB from '@fonts/LiberationSans-Bold.ttf?bytes';
import sansI from '@fonts/LiberationSans-Italic.ttf?bytes';
import sansBI from '@fonts/LiberationSans-BoldItalic.ttf?bytes';
// Liberation Serif (Times New Roman clone — Tinos).
import serifR from '@fonts/LiberationSerif-Regular.ttf?bytes';
import serifB from '@fonts/LiberationSerif-Bold.ttf?bytes';
import serifI from '@fonts/LiberationSerif-Italic.ttf?bytes';
import serifBI from '@fonts/LiberationSerif-BoldItalic.ttf?bytes';
// Carlito (Calibri clone).
import carlitoR from '@fonts/Carlito-Regular.ttf?bytes';
import carlitoB from '@fonts/Carlito-Bold.ttf?bytes';
import carlitoI from '@fonts/Carlito-Italic.ttf?bytes';
import carlitoBI from '@fonts/Carlito-BoldItalic.ttf?bytes';
// Caladea (Cambria clone).
import caladeaR from '@fonts/Caladea-Regular.ttf?bytes';
import caladeaB from '@fonts/Caladea-Bold.ttf?bytes';
import caladeaI from '@fonts/Caladea-Italic.ttf?bytes';
import caladeaBI from '@fonts/Caladea-BoldItalic.ttf?bytes';

const FACES: readonly FaceInput[] = [
  { family: 'Liberation Sans', bold: false, italic: false, bytes: sansR },
  { family: 'Liberation Sans', bold: true, italic: false, bytes: sansB },
  { family: 'Liberation Sans', bold: false, italic: true, bytes: sansI },
  { family: 'Liberation Sans', bold: true, italic: true, bytes: sansBI },
  { family: 'Liberation Serif', bold: false, italic: false, bytes: serifR },
  { family: 'Liberation Serif', bold: true, italic: false, bytes: serifB },
  { family: 'Liberation Serif', bold: false, italic: true, bytes: serifI },
  { family: 'Liberation Serif', bold: true, italic: true, bytes: serifBI },
  { family: 'Carlito', bold: false, italic: false, bytes: carlitoR },
  { family: 'Carlito', bold: true, italic: false, bytes: carlitoB },
  { family: 'Carlito', bold: false, italic: true, bytes: carlitoI },
  { family: 'Carlito', bold: true, italic: true, bytes: carlitoBI },
  { family: 'Caladea', bold: false, italic: false, bytes: caladeaR },
  { family: 'Caladea', bold: true, italic: false, bytes: caladeaB },
  { family: 'Caladea', bold: false, italic: true, bytes: caladeaI },
  { family: 'Caladea', bold: true, italic: true, bytes: caladeaBI },
];

/** The vendored clone faces (stable byte references → resolver metrics caching). */
export function loadRenderFaces(): readonly FaceInput[] {
  return FACES;
}

/** The clone display families that `registerPreviewFonts` makes available. */
export const CLONE_DISPLAY_FAMILIES: readonly string[] = [
  'Liberation Sans', 'Liberation Serif', 'Carlito', 'Caladea',
];

let registered = false;

/**
 * Register the clone faces as `FontFace`s for the preview's display. Idempotent;
 * a no-op where `FontFace`/`document.fonts` is unavailable (jsdom, the worker).
 * Display-only — layout positions come from the clone metrics, not the browser.
 */
export function registerPreviewFonts(doc: Document = globalThis.document): void {
  if (registered) return;
  if (typeof FontFace === 'undefined' || !doc?.fonts) return;
  for (const f of FACES) {
    try {
      // BufferSource source — never a url() (no fetch). Family name is the raw
      // clone family; CSS consumers quote names with spaces (Safari-safe).
      const face = new FontFace(f.family, f.bytes as BufferSource, {
        weight: f.bold ? '700' : '400',
        style: f.italic ? 'italic' : 'normal',
      });
      doc.fonts.add(face);
      void face.load().catch(() => {});
    } catch {
      // Display-only; a single failed face never blocks layout/PDF.
    }
  }
  registered = true;
}
