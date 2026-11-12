/**
 * Ambient declaration for `fontkit` (no bundled types). RenderCore only hands the
 * module straight to `pdf-lib`'s `registerFontkit`, which validates the shape at
 * runtime, so a loose declaration is sufficient and keeps the engine type-clean.
 */
declare module 'fontkit';
