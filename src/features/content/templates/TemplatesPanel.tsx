/**
 * Templates sub-panel — C9.
 *
 * Replaces the TemplatesPanel stub (task S1).  Provides three modes:
 *   Author    — mark variables, mark conditionals, define repeating groups, view schema list
 *   Fill      — questionnaire, conditional toggles, repeating editor, data-set push
 *   Data Sets — save current answers / load & push data sets
 *
 * All Office JS calls go through the driver; the panel is thin orchestration + UI.
 */

import { signal } from '@preact/signals';
import { getActiveProfile } from '@shared/config/index.js';
import type { TemplateSchema, AnswersMap, SharedDataSet } from './types.js';
import { SLOT_COUNT_MAX } from './types.js';
import { QuestionnaireForm, validateAnswers } from './QuestionnaireForm.js';
import { getExecutionBlocks, renderExecutionBlock } from './execution-blocks.js';
import {
  loadSchema,
  validateTag,
  markVariable,
  removeVariable,
  removeConditional,
  removeRepeatingGroup,
} from './author.js';
import {
  scanAndReconcile,
  fillQuestionnaire,
  buildInitialAnswers,
} from './fill.js';
import { applyConditionals, deleteConditionalDestructive } from './conditional.js';
import { listDataSets, loadDataSet, deleteDataSet, pushScalarDataSet, buildFuzzyPreview, pushRepeatingDataSet } from './datasets.js';
import type { ReconciliationResult } from './fill.js';

// ---------------------------------------------------------------------------
// Mode switcher
// ---------------------------------------------------------------------------

type TemplateMode = 'author' | 'fill' | 'datasets';

const _mode = signal<TemplateMode>('fill');

// ---------------------------------------------------------------------------
// Shared state
// ---------------------------------------------------------------------------

const _schema       = signal<TemplateSchema | null>(null);
const _loading      = signal(false);
const _error        = signal<string | null>(null);
const _statusMsg    = signal<string | null>(null);

/** Answers for the current fill session. */
const _answers      = signal<AnswersMap>({});
const _reconcile    = signal<ReconciliationResult | null>(null);
const _validErrors  = signal<Record<string, string>>({});

/** Fuzzy push preview waiting for user confirmation. */
const _fuzzyPreview = signal<{
  preview: ReturnType<typeof buildFuzzyPreview>;
  dataSet: SharedDataSet;
  confirmedKeys: Set<string>;
} | null>(null);

/** Author mode: which sub-form is open. */
type AuthorSubForm = 'none' | 'variable' | 'conditional' | 'repeating';
const _authorSubForm = signal<AuthorSubForm>('none');

// ---------------------------------------------------------------------------
// Helper to refresh schema + answers
// ---------------------------------------------------------------------------

async function refreshSchema(): Promise<void> {
  _loading.value = true;
  _error.value = null;
  try {
    const schema = await loadSchema();
    _schema.value = schema;
    // Also refresh the reconciled view if in fill mode.
    if (_mode.value === 'fill') {
      try {
        const rec = await scanAndReconcile();
        _reconcile.value = rec;
        _answers.value = buildInitialAnswers(schema, rec.documentControls);
      } catch {
        // Live document not available (e.g. unit tests) — leave answers as-is.
      }
    }
  } catch (err) {
    _error.value = `Failed to load schema: ${(err as Error).message ?? String(err)}`;
  } finally {
    _loading.value = false;
  }
}

// ---------------------------------------------------------------------------
// Author sub-panel
// ---------------------------------------------------------------------------

/** Mark Variable form */
function MarkVariableForm() {
  const tag      = signal('');
  const label    = signal('');
  const type     = signal('text');
  const required = signal(true);
  const err      = signal<string | null>(null);
  const busy     = signal(false);

  const submit = async () => {
    const tagErr = validateTag(tag.value.trim().toUpperCase());
    if (tagErr) { err.value = tagErr; return; }
    if (!label.value.trim()) { err.value = 'Label is required.'; return; }
    busy.value = true;
    err.value = null;
    const result = await markVariable({
      tag: tag.value.trim().toUpperCase(),
      label: label.value.trim(),
      variable: {
        tag: tag.value.trim().toUpperCase(),
        label: label.value.trim(),
        type: type.value as import('./types.js').VariableType,
        required: required.value,
      },
    });
    busy.value = false;
    if (!result.ok) { err.value = result.error ?? 'Unknown error.'; return; }
    _statusMsg.value = `Variable ${tag.value.trim().toUpperCase()} marked.`;
    _authorSubForm.value = 'none';
    await refreshSchema();
  };

  return (
    <div class="dp-tmpl-subform" lang="en-GB">
      <h4 class="dp-tmpl-subform__title">Mark as Variable</h4>
      <p class="dp-tmpl-hint">Select text in the document first, then complete this form.</p>
      <label class="dp-tmpl-label">
        Tag (UPPER_SNAKE)
        <input
          class="dp-tmpl-field"
          type="text"
          value={tag.value.toUpperCase()}
          onInput={(e) => { tag.value = (e.target as HTMLInputElement).value.toUpperCase(); }}
          placeholder="CLIENT_NAME"
        />
      </label>
      <label class="dp-tmpl-label">
        Label (shown to drafter)
        <input
          class="dp-tmpl-field"
          type="text"
          value={label.value}
          onInput={(e) => { label.value = (e.target as HTMLInputElement).value; }}
          placeholder="Client name"
        />
      </label>
      <label class="dp-tmpl-label">
        Type
        <select
          class="dp-tmpl-field"
          value={type.value}
          onChange={(e) => { type.value = (e.target as HTMLSelectElement).value; }}
        >
          <option value="text">Text</option>
          <option value="longtext">Long text</option>
          <option value="number">Number</option>
          <option value="date">Date (DD Month YYYY)</option>
          <option value="select">Select (dropdown)</option>
          <option value="entityType">Entity type (jurisdiction-aware)</option>
          <option value="regNumber">Registration number (UK_CH / SG_UEN)</option>
          <option value="address">Address (multi-line)</option>
          <option value="boolean">Boolean (checkbox)</option>
        </select>
      </label>
      <label class="dp-tmpl-checkbox">
        <input
          type="checkbox"
          checked={required.value}
          onChange={(e) => { required.value = (e.target as HTMLInputElement).checked; }}
        />
        {' '}Required field
      </label>
      {err.value && <p class="dp-tmpl-field-error">{err.value}</p>}
      <div class="dp-tmpl-actions">
        <button
          class="dp-tmpl-btn dp-tmpl-btn--primary"
          onClick={submit}
          disabled={busy.value}
        >
          {busy.value ? 'Marking…' : 'Mark Variable'}
        </button>
        <button
          class="dp-tmpl-btn"
          onClick={() => { _authorSubForm.value = 'none'; }}
        >
          Cancel
        </button>
      </div>
    </div>
  );
}

/** Author panel — schema list */
function AuthorPanel() {
  const schema = _schema.value;
  const subForm = _authorSubForm.value;

  const profile = getActiveProfile();
  const jurisdictionLabel = profile.id === 'SG' ? 'Singapore' : 'England & Wales';

  return (
    <div class="dp-tmpl-author" lang="en-GB">
      <p class="dp-tmpl-jurisdiction">Jurisdiction: <strong>{jurisdictionLabel}</strong></p>

      {subForm === 'none' && (
        <div class="dp-tmpl-actions dp-tmpl-actions--column">
          <button
            class="dp-tmpl-btn dp-tmpl-btn--primary"
            onClick={() => { _authorSubForm.value = 'variable'; }}
          >
            Mark as Variable
          </button>
          <button
            class="dp-tmpl-btn"
            onClick={() => { _statusMsg.value = 'Select text and use Mark Conditional (coming in extended UI).'; }}
          >
            Mark Conditional Section
          </button>
          <button
            class="dp-tmpl-btn"
            onClick={() => { _statusMsg.value = 'Define repeating group via author form (coming in extended UI).'; }}
          >
            Define Repeating Group
          </button>
        </div>
      )}

      {subForm === 'variable' && <MarkVariableForm />}

      {schema && (
        <section class="dp-tmpl-schema" aria-label="Defined variables">
          <h4 class="dp-tmpl-section-title">Variables ({schema.variables.length})</h4>
          {schema.variables.length === 0
            ? <p class="dp-tmpl-empty">No variables defined yet.</p>
            : (
              <ul class="dp-tmpl-schema-list">
                {schema.variables.map(v => (
                  <li key={v.tag} class="dp-tmpl-schema-item">
                    <span class="dp-tmpl-tag">{v.tag}</span>
                    <span class="dp-tmpl-label-text">{v.label}</span>
                    <span class="dp-tmpl-type-badge">{v.type}</span>
                    <button
                      class="dp-tmpl-btn dp-tmpl-btn--danger"
                      onClick={async () => {
                        await removeVariable(v.tag);
                        await refreshSchema();
                      }}
                      aria-label={`Remove variable ${v.tag}`}
                    >
                      ✕
                    </button>
                  </li>
                ))}
              </ul>
            )}

          <h4 class="dp-tmpl-section-title">Conditionals ({schema.conditionals.length})</h4>
          {schema.conditionals.length === 0
            ? <p class="dp-tmpl-empty">No conditional sections defined.</p>
            : (
              <ul class="dp-tmpl-schema-list">
                {schema.conditionals.map(c => (
                  <li key={c.tag} class="dp-tmpl-schema-item">
                    <span class="dp-tmpl-tag">{c.tag}</span>
                    <span class="dp-tmpl-label-text">{c.label}</span>
                    <span class="dp-tmpl-type-badge">{c.onHide}</span>
                    <button
                      class="dp-tmpl-btn dp-tmpl-btn--danger"
                      onClick={async () => {
                        await removeConditional(c.tag);
                        await refreshSchema();
                      }}
                      aria-label={`Remove conditional ${c.tag}`}
                    >
                      ✕
                    </button>
                  </li>
                ))}
              </ul>
            )}

          <h4 class="dp-tmpl-section-title">Repeating Groups ({schema.repeating.length})</h4>
          {schema.repeating.length === 0
            ? (
              <p class="dp-tmpl-empty">
                No repeating groups. Fixed-slot cap: up to {SLOT_COUNT_MAX} slots.
              </p>
            )
            : (
              <ul class="dp-tmpl-schema-list">
                {schema.repeating.map(r => (
                  <li key={r.group} class="dp-tmpl-schema-item">
                    <span class="dp-tmpl-tag">{r.group}</span>
                    <span class="dp-tmpl-label-text">Up to {r.slotCount} slots</span>
                    <button
                      class="dp-tmpl-btn dp-tmpl-btn--danger"
                      onClick={async () => {
                        await removeRepeatingGroup(r.group);
                        await refreshSchema();
                      }}
                      aria-label={`Remove group ${r.group}`}
                    >
                      ✕
                    </button>
                  </li>
                ))}
              </ul>
            )}
        </section>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Fill panel
// ---------------------------------------------------------------------------

function FillPanel() {
  const schema    = _schema.value;
  const answers   = _answers.value;
  const reconcile = _reconcile.value;
  const valErrs   = _validErrors.value;
  const busy      = signal(false);
  const msg       = signal<string | null>(null);

  if (!schema) {
    return <p class="dp-tmpl-empty">Loading schema…</p>;
  }

  const handleChange = (tag: string, value: string | boolean) => {
    _answers.value = { ..._answers.value, [tag]: value };
  };

  const handleSubmit = async () => {
    const errs = validateAnswers(schema.variables, _answers.value, schema.jurisdictionProfile);
    _validErrors.value = errs;
    if (Object.keys(errs).length > 0) return;

    busy.value = true;
    msg.value = null;
    try {
      // Build fill items.
      const items = schema.variables.map(v => ({
        tag: v.tag,
        type: v.type,
        rawValue: _answers.value[v.tag] ?? v.default ?? '',
      }));

      const fillResult = await fillQuestionnaire(items);

      // Apply conditionals.
      if (schema.conditionals.length > 0) {
        await applyConditionals(schema.conditionals, _answers.value);
      }

      const written = fillResult.written.length;
      const notFound = fillResult.notFound.length;
      msg.value = `Filled ${written} variable${written === 1 ? '' : 's'}.` +
        (notFound > 0 ? ` ${notFound} tag(s) not found in document — check drift warnings.` : '');
    } catch (err) {
      msg.value = `Error: ${(err as Error).message ?? String(err)}`;
    } finally {
      busy.value = false;
    }
  };

  return (
    <div class="dp-tmpl-fill" lang="en-GB">
      {reconcile && reconcile.inDocNotSchema.length > 0 && (
        <div class="dp-tmpl-warn" role="alert">
          <strong>Tags in document but not in schema:</strong>
          <ul>
            {reconcile.inDocNotSchema.map(t => (
              <li key={t.tag}>{t.tag} — "{t.currentText}"</li>
            ))}
          </ul>
        </div>
      )}
      {reconcile && reconcile.inSchemaNotDoc.length > 0 && (
        <div class="dp-tmpl-warn" role="alert">
          <strong>Tags in schema but not found as content controls:</strong>
          <ul>
            {reconcile.inSchemaNotDoc.map(t => (
              <li key={t.tag}>{t.tag} ({t.label})</li>
            ))}
          </ul>
        </div>
      )}

      <QuestionnaireForm
        variables={schema.variables}
        jurisdiction={schema.jurisdictionProfile}
        answers={answers}
        onChange={handleChange}
        validationErrors={valErrs}
      />

      {schema.conditionals.length > 0 && (
        <section class="dp-tmpl-conditionals" aria-label="Conditional sections">
          <h4 class="dp-tmpl-section-title">Conditional Sections</h4>
          {schema.conditionals.map(cond => (
            <div key={cond.tag} class="dp-tmpl-cond-item">
              <span class="dp-tmpl-label-text">{cond.label}</span>
              {cond.onHide === 'destructive-confirm' && (
                <button
                  class="dp-tmpl-btn dp-tmpl-btn--danger"
                  onClick={async () => {
                    if (window.confirm(
                      `Permanently remove "${cond.label}"?\n\nThis is irreversible and cannot be undone within this session.`,
                    )) {
                      await deleteConditionalDestructive(cond.tag);
                      _statusMsg.value = `Section "${cond.label}" permanently removed.`;
                    }
                  }}
                >
                  Remove (irreversible)
                </button>
              )}
            </div>
          ))}
        </section>
      )}

      {schema.repeating.length > 0 && (
        <section class="dp-tmpl-repeating" aria-label="Repeating groups">
          <h4 class="dp-tmpl-section-title">Repeating Groups</h4>
          {schema.repeating.map(grp => (
            <div key={grp.group} class="dp-tmpl-repeat-group">
              <strong>{grp.group}</strong>
              <span class="dp-tmpl-hint">
                {' '}— up to {grp.slotCount} {grp.slotCount === 1 ? 'slot' : 'slots'} (max {SLOT_COUNT_MAX})
              </span>
              <p class="dp-tmpl-hint">Use the Data Sets panel to push rows into this group.</p>
            </div>
          ))}
        </section>
      )}

      {msg.value && <p class="dp-tmpl-status">{msg.value}</p>}

      <div class="dp-tmpl-actions">
        <button
          class="dp-tmpl-btn dp-tmpl-btn--primary"
          onClick={handleSubmit}
          disabled={busy.value}
        >
          {busy.value ? 'Filling…' : 'Fill Template'}
        </button>
        <button
          class="dp-tmpl-btn"
          onClick={async () => {
            await refreshSchema();
            _statusMsg.value = 'Refreshed from document.';
          }}
        >
          Refresh
        </button>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Data Sets panel
// ---------------------------------------------------------------------------

function DataSetsPanel() {
  const schema        = _schema.value;
  const dsIds         = signal<string[]>([]);
  const loadingDs     = signal(false);
  const dsMsg         = signal<string | null>(null);
  const fuzzyPrev     = _fuzzyPreview.value;

  const loadList = async () => {
    loadingDs.value = true;
    const ids = await listDataSets('personal');
    dsIds.value = ids;
    loadingDs.value = false;
  };

  // Execution blocks.
  const execBlocks = schema
    ? getExecutionBlocks(schema.jurisdictionProfile)
    : [];

  const handlePushDataSet = async (id: string) => {
    if (!schema) return;
    const ds = await loadDataSet(id, 'personal');
    if (!ds) { dsMsg.value = 'Data set not found.'; return; }

    if (ds.kind === 'scalar') {
      const preview = buildFuzzyPreview(ds.values ?? {}, schema);
      if (preview.fuzzy.length > 0 || preview.unmatched.length > 0) {
        _fuzzyPreview.value = { preview, dataSet: ds, confirmedKeys: new Set() };
        return;
      }
      // All exact — apply directly.
      const result = await pushScalarDataSet({ dataSet: ds, schema, confirmedFuzzy: [] });
      dsMsg.value = `Wrote ${result.written.length} values. Skipped: ${result.skipped.join(', ') || 'none'}.`;
    } else {
      const result = await pushRepeatingDataSet(ds, schema);
      dsMsg.value = result.capError ?? `Wrote ${result.written.length} slot tags.`;
    }
  };

  return (
    <div class="dp-tmpl-datasets" lang="en-GB">
      {/* Backend residency indicator */}
      <p class="dp-tmpl-hint">
        Storage: <strong>local (personal)</strong> — shared backend not configured.
        SharePoint is the favoured backend for firm-shared data sets (SRA/PDPA data residency).
      </p>

      <button
        class="dp-tmpl-btn"
        onClick={loadList}
        disabled={loadingDs.value}
      >
        {loadingDs.value ? 'Loading…' : 'Refresh Data Sets'}
      </button>

      {dsIds.value.length > 0 && (
        <ul class="dp-tmpl-schema-list">
          {dsIds.value.map(id => (
            <li key={id} class="dp-tmpl-schema-item">
              <span class="dp-tmpl-tag">{id}</span>
              <button
                class="dp-tmpl-btn dp-tmpl-btn--primary"
                onClick={() => handlePushDataSet(id)}
              >
                Push
              </button>
              <button
                class="dp-tmpl-btn dp-tmpl-btn--danger"
                onClick={async () => {
                  await deleteDataSet(id, 'personal');
                  await loadList();
                }}
              >
                Delete
              </button>
            </li>
          ))}
        </ul>
      )}

      {/* Fuzzy match confirmation */}
      {fuzzyPrev && (
        <div class="dp-tmpl-fuzzy-warn" role="alert">
          <strong>Fuzzy matches found — please confirm before applying:</strong>
          <ul>
            {fuzzyPrev.preview.fuzzy.map(w => (
              <li key={w.dataSetKey}>
                <label class="dp-tmpl-checkbox">
                  <input
                    type="checkbox"
                    checked={fuzzyPrev.confirmedKeys.has(w.dataSetKey)}
                    onChange={(e) => {
                      const confirmed = new Set(fuzzyPrev.confirmedKeys);
                      if ((e.target as HTMLInputElement).checked) {
                        confirmed.add(w.dataSetKey);
                      } else {
                        confirmed.delete(w.dataSetKey);
                      }
                      _fuzzyPreview.value = { ...fuzzyPrev, confirmedKeys: confirmed };
                    }}
                  />
                  {' '}Map "{w.dataSetKey}" → {w.schemaTag} (score: {(w.score * 100).toFixed(0)}%)
                </label>
              </li>
            ))}
          </ul>
          {fuzzyPrev.preview.unmatched.length > 0 && (
            <p class="dp-tmpl-warn-text">
              Unmatched keys (will be skipped): {fuzzyPrev.preview.unmatched.join(', ')}
            </p>
          )}
          <div class="dp-tmpl-actions">
            <button
              class="dp-tmpl-btn dp-tmpl-btn--primary"
              onClick={async () => {
                if (!schema || !fuzzyPrev) return;
                const confirmed = Array.from(fuzzyPrev.confirmedKeys).map(k =>
                  fuzzyPrev.preview.fuzzy.find(w => w.dataSetKey === k)!,
                ).filter(Boolean);
                const result = await pushScalarDataSet({
                  dataSet: fuzzyPrev.dataSet,
                  schema,
                  confirmedFuzzy: confirmed,
                });
                _fuzzyPreview.value = null;
                dsMsg.value = `Wrote ${result.written.length} values. Skipped: ${result.skipped.join(', ') || 'none'}.`;
              }}
            >
              Apply
            </button>
            <button
              class="dp-tmpl-btn"
              onClick={() => { _fuzzyPreview.value = null; }}
            >
              Cancel
            </button>
          </div>
        </div>
      )}

      {dsMsg.value && <p class="dp-tmpl-status">{dsMsg.value}</p>}

      {/* Execution blocks */}
      <section class="dp-tmpl-exec-blocks" aria-label="Execution blocks">
        <h4 class="dp-tmpl-section-title">Execution / Attestation Blocks</h4>
        {execBlocks.length === 0
          ? <p class="dp-tmpl-empty">No execution blocks available for current jurisdiction.</p>
          : execBlocks.map(block => (
            <div key={block.id} class="dp-tmpl-exec-block">
              <strong class="dp-tmpl-exec-label">{block.label}</strong>
              <p class="dp-tmpl-hint">{block.basis}</p>
              <pre class="dp-tmpl-exec-preview">{renderExecutionBlock(block)}</pre>
              <button
                class="dp-tmpl-btn"
                onClick={() => {
                  // Copy block text to clipboard for insertion.
                  void navigator.clipboard.writeText(renderExecutionBlock(block));
                  dsMsg.value = `"${block.label}" copied to clipboard.`;
                }}
              >
                Copy to clipboard
              </button>
            </div>
          ))}
      </section>
    </div>
  );
}

// ---------------------------------------------------------------------------
// TemplatesPanel root
// ---------------------------------------------------------------------------

export function TemplatesPanel() {
  const mode = _mode.value;

  // Initialise schema on first mount.
  if (_schema.value === null && !_loading.value) {
    void refreshSchema();
  }

  return (
    <div class="dp-tmpl" lang="en-GB">
      <nav class="dp-tmpl-modenav" role="tablist" aria-label="Template tools">
        {(['author', 'fill', 'datasets'] as TemplateMode[]).map(m => (
          <button
            key={m}
            type="button"
            role="tab"
            aria-selected={mode === m}
            class={`dp-tmpl-modetab${mode === m ? ' dp-tmpl-modetab--active' : ''}`}
            onClick={() => {
              _mode.value = m;
              void refreshSchema();
            }}
          >
            {m === 'author' ? 'Author' : m === 'fill' ? 'Fill' : 'Data Sets'}
          </button>
        ))}
      </nav>

      {_error.value && (
        <p class="dp-tmpl-field-error" role="alert">{_error.value}</p>
      )}
      {_statusMsg.value && (
        <p class="dp-tmpl-status" role="status">{_statusMsg.value}</p>
      )}
      {_loading.value && (
        <p class="dp-tmpl-loading" aria-live="polite">Loading…</p>
      )}

      <div class="dp-tmpl-content" role="tabpanel">
        {mode === 'author' && <AuthorPanel />}
        {mode === 'fill'   && <FillPanel />}
        {mode === 'datasets' && <DataSetsPanel />}
      </div>
    </div>
  );
}
