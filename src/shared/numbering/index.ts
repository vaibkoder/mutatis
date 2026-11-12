/**
 * Shared numbering module — host-neutral OOXML list-number resolver.
 * Parses `word/numbering.xml` into a scheme and computes rendered labels
 * ("6.1", "(a)") via a document-order counter state machine. Consumed by the
 * web host (numbering driver) and by RenderCore (feature 16). No Office/Word.
 */
export * from './numbering-scheme.js';
export * from './style-numbering.js';
export * from './numbering-format.js';
export * from './render-list-numbers.js';
