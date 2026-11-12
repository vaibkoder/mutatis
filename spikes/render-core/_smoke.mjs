// RenderCore spike harness smoke test (THROWAWAY — not bundled).
// Confirms harfbuzzjs + fontkit + @pdfme/pdf-lib load in Node and a font shapes.
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { hb, loadFont, widthPt } from './hb-util.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
const asset = (f) => join(HERE, 'assets', f);

console.log('— loading deps —');
console.log('harfbuzzjs HB version:', hb.versionString());
const fontkit = (await import('fontkit')).default ?? (await import('fontkit'));
console.log('fontkit loaded:', typeof fontkit.create === 'function');
const pdfLib = await import('@pdfme/pdf-lib');
console.log('@pdfme/pdf-lib PDFDocument:', typeof pdfLib.PDFDocument);

console.log('— shaping "Office" with Liberation Sans (kern/liga off) —');
const { font, upem } = loadFont(asset('LiberationSans-Regular.ttf'));
console.log('upem:', upem, 'width of "Office" @12pt:', widthPt(font, upem, 'Office', 12).toFixed(2), 'pt');

console.log('SMOKE OK');
