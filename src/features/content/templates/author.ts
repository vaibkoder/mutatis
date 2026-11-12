/**
 * Template author logic — C2.
 *
 * Mark a selected text range as a tagged PlainText content control.
 * Persist the variable metadata to docSettings.templateSchema.
 *
 * Office JS constraints (ARCHITECTURE.md):
 *   - insertContentControl(PlainText) — available WordApi 1.1
 *   - .tag / .title / .appearance — settable on same sync
 *   - Tags must be UPPER_SNAKE and unique within the document
 *   - Batch: set all properties before ctx.sync() — no per-property syncs
 */

import { wordRun } from '@shared/office/index.js';
import { getStorageProvider } from '@shared/storage/index.js';
import type { TemplateVariable, TemplateSchema } from './types.js';
import { TEMPLATE_SCHEMA_KEY, SLOT_COUNT_DEFAULT, guardSlotCount } from './types.js';

// ---------------------------------------------------------------------------
// Tag validation helpers
// ---------------------------------------------------------------------------

/** UPPER_SNAKE pattern — letters, digits, underscore, must start with a letter. */
const UPPER_SNAKE_RE = /^[A-Z][A-Z0-9_]*$/;

/** Validate a tag string. Returns an error message or null if valid. */
export function validateTag(tag: string): string | null {
  if (!tag.trim()) return 'Tag must not be empty.';
  if (!UPPER_SNAKE_RE.test(tag)) {
    return `Tag "${tag}" must be UPPER_SNAKE case (letters, digits, underscore; start with a letter).`;
  }
  return null;
}

// ---------------------------------------------------------------------------
// Schema persistence helpers (docSettings)
// ---------------------------------------------------------------------------

/** Load the current template schema from docSettings (or return a blank schema). */
export async function loadSchema(): Promise<TemplateSchema> {
  const store = getStorageProvider().docSettings;
  const stored = await store.get(TEMPLATE_SCHEMA_KEY) as TemplateSchema | undefined;
  if (stored && typeof stored === 'object' && Array.isArray(stored.variables)) {
    return stored;
  }
  return {
    templateId: `tmpl_${Date.now()}`,
    jurisdictionProfile: 'EW',
    variables: [],
    conditionals: [],
    repeating: [],
  };
}

/** Persist the template schema to docSettings. */
export async function saveSchema(schema: TemplateSchema): Promise<void> {
  const store = getStorageProvider().docSettings;
  await store.set(TEMPLATE_SCHEMA_KEY, schema as unknown);
}

// ---------------------------------------------------------------------------
// Mark variable — C2 core
// ---------------------------------------------------------------------------

export interface MarkVariableRequest {
  /** UPPER_SNAKE tag — must be unique in the document. */
  tag: string;
  /** Human label shown in the questionnaire (== content-control .title). */
  label: string;
  variable: TemplateVariable;
}

export interface MarkVariableResult {
  ok: boolean;
  /** Set if the operation failed (duplicate tag, Office error, etc.). */
  error?: string;
}

/**
 * Insert a PlainText content control around the current selection,
 * tag it with the given UPPER_SNAKE identifier, and persist the variable
 * schema to docSettings.
 *
 * Returns ok:false with an error message if:
 *   - the tag is not valid UPPER_SNAKE
 *   - the tag already exists in the schema
 *
 * NOTE: uniqueness is checked against the *schema*, not the live document
 * content controls. A tag removed from the schema but still in the document
 * will be surfaced as a drift warning at fill time (C3).
 */
export async function markVariable(
  req: MarkVariableRequest,
): Promise<MarkVariableResult> {
  // 1. Validate the tag.
  const tagError = validateTag(req.tag);
  if (tagError) return { ok: false, error: tagError };

  // 2. Load schema and check for duplicate.
  const schema = await loadSchema();
  if (schema.variables.some(v => v.tag === req.tag)) {
    return {
      ok: false,
      error: `Tag "${req.tag}" already exists in the template schema. Each tag must be unique.`,
    };
  }

  // 3. Insert the content control in the document.
  try {
    await wordRun(async ctx => {
      const sel = ctx.document.getSelection();
      // Word.ContentControlType.plainText = 'PlainText'
      const cc = sel.insertContentControl('PlainText');
      cc.tag = req.tag;
      cc.title = req.label;
      // boundingBox = authoring appearance
      cc.appearance = 'BoundingBox';
      await ctx.sync();
    });
  } catch (err) {
    return {
      ok: false,
      error: `Failed to insert content control: ${(err as Error).message ?? String(err)}`,
    };
  }

  // 4. Persist the variable to the schema.
  schema.variables.push({ ...req.variable, tag: req.tag, label: req.label });
  await saveSchema(schema);

  return { ok: true };
}

// ---------------------------------------------------------------------------
// Mark conditional section — C2 (author mode)
// ---------------------------------------------------------------------------

export interface MarkConditionalRequest {
  tag: string;
  label: string;
  showWhen: TemplateSchema['conditionals'][number]['showWhen'];
  onHide: TemplateSchema['conditionals'][number]['onHide'];
}

/**
 * Wrap the current selection in a RichText content control tagged for
 * conditional show/hide, and persist the conditional to the schema.
 */
export async function markConditionalSection(
  req: MarkConditionalRequest,
): Promise<MarkVariableResult> {
  const tagError = validateTag(req.tag);
  if (tagError) return { ok: false, error: tagError };

  const schema = await loadSchema();
  if (schema.conditionals.some(c => c.tag === req.tag)) {
    return {
      ok: false,
      error: `Conditional tag "${req.tag}" already exists in the schema.`,
    };
  }

  try {
    await wordRun(async ctx => {
      const sel = ctx.document.getSelection();
      const cc = sel.insertContentControl('RichText');
      cc.tag = req.tag;
      cc.title = req.label;
      cc.appearance = 'BoundingBox';
      await ctx.sync();
    });
  } catch (err) {
    return {
      ok: false,
      error: `Failed to insert conditional content control: ${(err as Error).message ?? String(err)}`,
    };
  }

  schema.conditionals.push({
    tag: req.tag,
    label: req.label,
    showWhen: req.showWhen,
    onHide: req.onHide,
  });
  await saveSchema(schema);

  return { ok: true };
}

// ---------------------------------------------------------------------------
// Define repeating group — C2 (author mode)
// ---------------------------------------------------------------------------

export interface DefineRepeatingGroupRequest {
  group: string;
  slotCount?: number;
  slotTags: TemplateSchema['repeating'][number]['slotTags'];
  hideUnusedSlots?: boolean;
}

/**
 * Add a repeating group definition to the schema.
 * slotCount is clamped to 1–SLOT_COUNT_MAX (default SLOT_COUNT_DEFAULT).
 */
export async function defineRepeatingGroup(
  req: DefineRepeatingGroupRequest,
): Promise<MarkVariableResult> {
  const tagError = validateTag(req.group);
  if (tagError) return { ok: false, error: tagError };

  const schema = await loadSchema();
  if (schema.repeating.some(r => r.group === req.group)) {
    return {
      ok: false,
      error: `Repeating group "${req.group}" already exists in the schema.`,
    };
  }

  let slotCount = req.slotCount ?? SLOT_COUNT_DEFAULT;
  try {
    slotCount = guardSlotCount(slotCount);
  } catch (err) {
    return { ok: false, error: (err as Error).message };
  }

  schema.repeating.push({
    group: req.group,
    slotCount,
    slotTags: req.slotTags,
    hideUnusedSlots: req.hideUnusedSlots ?? true,
  });
  await saveSchema(schema);

  return { ok: true };
}

// ---------------------------------------------------------------------------
// Schema deletion helpers (for the schema list edit/delete UI)
// ---------------------------------------------------------------------------

/** Remove a variable from the schema by tag. */
export async function removeVariable(tag: string): Promise<void> {
  const schema = await loadSchema();
  schema.variables = schema.variables.filter(v => v.tag !== tag);
  await saveSchema(schema);
}

/** Remove a conditional section from the schema by tag. */
export async function removeConditional(tag: string): Promise<void> {
  const schema = await loadSchema();
  schema.conditionals = schema.conditionals.filter(c => c.tag !== tag);
  await saveSchema(schema);
}

/** Remove a repeating group from the schema by group identifier. */
export async function removeRepeatingGroup(group: string): Promise<void> {
  const schema = await loadSchema();
  schema.repeating = schema.repeating.filter(r => r.group !== group);
  await saveSchema(schema);
}
