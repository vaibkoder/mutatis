/** Shared fixtures for RenderCore unit tests (Phase 1). */
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { FaceInput } from '@shared/render-core/fonts/registry.js';

const root = resolve(__dirname, '../../..');

/** OFL Liberation Sans (committed under spikes/render-core/assets) — exact-metric Arial clone. */
export const libSansRegular = (): Uint8Array =>
  new Uint8Array(readFileSync(resolve(root, 'spikes/render-core/assets/LiberationSans-Regular.ttf')));
export const libSansBold = (): Uint8Array =>
  new Uint8Array(readFileSync(resolve(root, 'spikes/render-core/assets/LiberationSans-Bold.ttf')));

/**
 * A face set covering Serif+Sans (both backed by Liberation Sans bytes — fine for
 * layout-logic tests; only metrics/advances matter, not glyph shapes).
 */
export function testFaces(): FaceInput[] {
  const reg = libSansRegular();
  const bold = libSansBold();
  return [
    { family: 'Liberation Serif', bold: false, italic: false, bytes: reg },
    { family: 'Liberation Serif', bold: true, italic: false, bytes: bold },
    { family: 'Liberation Sans', bold: false, italic: false, bytes: reg },
    { family: 'Liberation Sans', bold: true, italic: false, bytes: bold },
  ];
}
