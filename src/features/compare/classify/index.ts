/**
 * Feature 08 — Compare · classifier public surface (Tasks B1–B3).
 *
 * Rule-based, priority-ordered change classification, SG/UK forcing rules, and
 * renumbering-cascade collapse. CATEGORISES, never SUPPRESSES (ADR-0008).
 *
 * Spec: spec/features/08-compare.md §6–§7.
 */

export { classifyOp, classifyAll } from './classifier.js';
export { classifyClasses, PRIORITY_ORDER } from './taxonomy.js';
export { applySgUkRules } from './sguk-rules.js';
export { detectSpellingVariant } from './spelling-variant.js';
export { applySignatureBlockRule } from './signature-block.js';
export { computeCascades } from './cascade.js';
