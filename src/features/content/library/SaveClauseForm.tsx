/**
 * Precedent / Clause Library — Save Clause form component.
 *
 * Spec: spec/features/07-toc-precedent-templates.md §B.7 ("Save panel")
 *
 * Rendered inside LibraryPanel when saveFormOpen is true.
 * Required fields enforced inline; Save button disabled until all required fields valid.
 * Visibility toggle disabled when no shared backend is configured.
 */

import type { ClauseMetadata } from './types.js';
import {
  saveFormOpen,
  saveFormTitle,
  saveFormMetadata,
  saveFormErrors,
  saveFormWarnings,
  saveBusy,
  saveSuccess,
  saveFormValid,
  submitSaveForm,
} from './driver.js';
import { hasSharedBackend } from './library-store.js';

// ---------------------------------------------------------------------------
// Practice area options (firm-level config; hardcoded defaults per SG/UK scope)
// ---------------------------------------------------------------------------

const PRACTICE_AREAS: string[] = [
  'Banking & Finance',
  'Capital Markets',
  'Commercial',
  'Construction',
  'Corporate',
  'Dispute Resolution',
  'Employment',
  'Financial Regulatory',
  'Insolvency & Restructuring',
  'Intellectual Property',
  'M&A',
  'Private Client',
  'Projects',
  'Property',
  'Shipping & Transport',
  'Technology & Data',
  'Other',
];

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function fieldError(field: string): string | undefined {
  return saveFormErrors.value.find(e => e.field === field)?.message;
}

function fieldWarning(field: string): string | undefined {
  return saveFormWarnings.value.find(w => w.field === field)?.message;
}

function updateMeta<K extends keyof ClauseMetadata>(key: K, value: ClauseMetadata[K]): void {
  saveFormMetadata.value = { ...saveFormMetadata.value, [key]: value };
}

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

export function SaveClauseForm() {
  const meta = saveFormMetadata.value;
  const backendConfigured = hasSharedBackend();

  function handleSubmit(e: Event) {
    e.preventDefault();
    // Author identity: pull from Office context if available, else use a placeholder.
    let createdBy = 'Unknown';
    try {
      const profile = (typeof Office !== 'undefined' && Office.context?.mailbox?.userProfile);
      if (profile) createdBy = profile.displayName || profile.emailAddress || createdBy;
    } catch { /* ignore */ }

    void submitSaveForm({
      title: saveFormTitle.value,
      metadata: meta as ClauseMetadata,
      createdBy,
    });
  }

  const genError = fieldError('_');

  return (
    <form
      class="dp-library__save-form"
      onSubmit={handleSubmit}
      lang="en-GB"
      aria-label="Save clause to library"
    >
      <h3 class="dp-library__save-form-title">Save Clause to Library</h3>

      {/* Title */}
      <div class="dp-library__field">
        <label for="dp-lib-title" class="dp-library__label">
          Title <span class="dp-library__required" aria-label="required">*</span>
        </label>
        <input
          id="dp-lib-title"
          type="text"
          class={`dp-library__input${fieldError('title') ? ' dp-library__input--error' : ''}`}
          value={saveFormTitle.value}
          onInput={(e) => { saveFormTitle.value = (e.target as HTMLInputElement).value; }}
          disabled={saveBusy.value}
          aria-describedby={fieldError('title') ? 'dp-lib-title-error' : undefined}
          aria-required="true"
          maxLength={200}
          placeholder="e.g. Limitation of Liability (cap at fees)"
        />
        {fieldError('title') && (
          <span id="dp-lib-title-error" class="dp-library__error" role="alert">
            {fieldError('title')}
          </span>
        )}
      </div>

      {/* Jurisdiction */}
      <div class="dp-library__field">
        <label for="dp-lib-jurisdiction" class="dp-library__label">
          Jurisdiction <span class="dp-library__required" aria-label="required">*</span>
        </label>
        <select
          id="dp-lib-jurisdiction"
          class={`dp-library__select${fieldError('jurisdiction') ? ' dp-library__input--error' : ''}`}
          value={meta.jurisdiction ?? ''}
          onChange={(e) => {
            updateMeta('jurisdiction', (e.target as HTMLSelectElement).value as ClauseMetadata['jurisdiction']);
          }}
          disabled={saveBusy.value}
          aria-required="true"
        >
          <option value="">Select…</option>
          <option value="SG">Singapore</option>
          <option value="EW">England &amp; Wales</option>
          <option value="OTHER">Other (specify below)</option>
        </select>
        {fieldError('jurisdiction') && (
          <span class="dp-library__error" role="alert">{fieldError('jurisdiction')}</span>
        )}
      </div>

      {/* jurisdictionOther (only shown when OTHER) */}
      {meta.jurisdiction === 'OTHER' && (
        <div class="dp-library__field">
          <label for="dp-lib-jurisdiction-other" class="dp-library__label">
            Specify jurisdiction <span class="dp-library__required" aria-label="required">*</span>
          </label>
          <input
            id="dp-lib-jurisdiction-other"
            type="text"
            class={`dp-library__input${fieldError('jurisdictionOther') ? ' dp-library__input--error' : ''}`}
            value={meta.jurisdictionOther ?? ''}
            onInput={(e) => {
              updateMeta('jurisdictionOther', (e.target as HTMLInputElement).value);
            }}
            disabled={saveBusy.value}
            aria-required="true"
            placeholder="e.g. New York, Delaware"
          />
          {fieldError('jurisdictionOther') && (
            <span class="dp-library__error" role="alert">{fieldError('jurisdictionOther')}</span>
          )}
        </div>
      )}

      {/* Practice area */}
      <div class="dp-library__field">
        <label for="dp-lib-practice-area" class="dp-library__label">
          Practice area <span class="dp-library__required" aria-label="required">*</span>
        </label>
        <select
          id="dp-lib-practice-area"
          class={`dp-library__select${fieldError('practiceArea') ? ' dp-library__input--error' : ''}`}
          value={meta.practiceArea ?? ''}
          onChange={(e) => {
            updateMeta('practiceArea', (e.target as HTMLSelectElement).value);
          }}
          disabled={saveBusy.value}
          aria-required="true"
        >
          <option value="">Select…</option>
          {PRACTICE_AREAS.map(pa => (
            <option key={pa} value={pa}>{pa}</option>
          ))}
        </select>
        {fieldError('practiceArea') && (
          <span class="dp-library__error" role="alert">{fieldError('practiceArea')}</span>
        )}
      </div>

      {/* Last reviewed */}
      <div class="dp-library__field">
        <label for="dp-lib-last-reviewed" class="dp-library__label">
          Last reviewed <span class="dp-library__required" aria-label="required">*</span>
        </label>
        <input
          id="dp-lib-last-reviewed"
          type="date"
          class={`dp-library__input${fieldError('lastReviewed') ? ' dp-library__input--error' : ''}`}
          value={meta.lastReviewed ?? ''}
          onInput={(e) => {
            updateMeta('lastReviewed', (e.target as HTMLInputElement).value);
          }}
          disabled={saveBusy.value}
          aria-required="true"
          max={new Date().toISOString().slice(0, 10)}   // prevent future date in the picker
        />
        {fieldError('lastReviewed') && (
          <span class="dp-library__error" role="alert">{fieldError('lastReviewed')}</span>
        )}
        {!fieldError('lastReviewed') && fieldWarning('lastReviewed') && (
          <span class="dp-library__warning" role="note">{fieldWarning('lastReviewed')}</span>
        )}
      </div>

      {/* Approved by */}
      <div class="dp-library__field">
        <label for="dp-lib-approved-by" class="dp-library__label">
          Approved by <span class="dp-library__required" aria-label="required">*</span>
        </label>
        <input
          id="dp-lib-approved-by"
          type="text"
          class={`dp-library__input${fieldError('approvedBy') ? ' dp-library__input--error' : ''}`}
          value={meta.approvedBy ?? ''}
          onInput={(e) => {
            updateMeta('approvedBy', (e.target as HTMLInputElement).value);
          }}
          disabled={saveBusy.value}
          aria-required="true"
          placeholder="e.g. J. Tan (Partner)"
        />
        {fieldError('approvedBy') && (
          <span class="dp-library__error" role="alert">{fieldError('approvedBy')}</span>
        )}
      </div>

      {/* Visibility */}
      <div class="dp-library__field">
        <label class="dp-library__label">
          Visibility <span class="dp-library__required" aria-label="required">*</span>
        </label>
        <div class="dp-library__radio-group" role="radiogroup" aria-label="Visibility">
          <label class="dp-library__radio-label">
            <input
              type="radio"
              name="visibility"
              value="personal"
              checked={meta.visibility === 'personal' || !meta.visibility}
              onChange={() => updateMeta('visibility', 'personal')}
              disabled={saveBusy.value}
            />
            {' '}Personal (local only)
          </label>
          <label class={`dp-library__radio-label${!backendConfigured ? ' dp-library__radio-label--disabled' : ''}`}>
            <input
              type="radio"
              name="visibility"
              value="shared"
              checked={meta.visibility === 'shared'}
              onChange={() => updateMeta('visibility', 'shared')}
              disabled={saveBusy.value || !backendConfigured}
            />
            {' '}Shared
            {!backendConfigured && (
              <span class="dp-library__field-hint">
                {' '}(requires SharePoint or GitHub backend — not yet configured)
              </span>
            )}
          </label>
        </div>
        {fieldError('visibility') && (
          <span class="dp-library__error" role="alert">{fieldError('visibility')}</span>
        )}
      </div>

      {/* Optional: tags */}
      <div class="dp-library__field">
        <label for="dp-lib-tags" class="dp-library__label">Tags</label>
        <input
          id="dp-lib-tags"
          type="text"
          class="dp-library__input"
          value={(meta.tags ?? []).join(', ')}
          onInput={(e) => {
            const raw = (e.target as HTMLInputElement).value;
            const tags = raw.split(',').map(t => t.trim()).filter(t => t.length > 0);
            updateMeta('tags', tags);
          }}
          disabled={saveBusy.value}
          placeholder="e.g. liability, cap, indemnity"
        />
        <span class="dp-library__field-hint">Comma-separated tags for search.</span>
      </div>

      {/* Optional: notes */}
      <div class="dp-library__field">
        <label for="dp-lib-notes" class="dp-library__label">Drafting notes</label>
        <textarea
          id="dp-lib-notes"
          class="dp-library__textarea"
          value={meta.notes ?? ''}
          onInput={(e) => {
            updateMeta('notes', (e.target as HTMLTextAreaElement).value);
          }}
          disabled={saveBusy.value}
          rows={3}
          placeholder="Optional guidance for practitioners using this clause."
        />
      </div>

      {/* General error */}
      {genError && (
        <div class="dp-library__error dp-library__error--general" role="alert">
          {genError}
        </div>
      )}

      {/* Buttons */}
      <div class="dp-library__form-actions">
        <button
          type="submit"
          class="dp-library__btn dp-library__btn--primary"
          disabled={saveBusy.value || !saveFormValid.value}
          aria-busy={saveBusy.value}
        >
          {saveBusy.value ? 'Saving…' : 'Save Clause'}
        </button>
        <button
          type="button"
          class="dp-library__btn dp-library__btn--secondary"
          onClick={() => { saveFormOpen.value = false; }}
          disabled={saveBusy.value}
        >
          Cancel
        </button>
      </div>

      {saveSuccess.value && (
        <div class="dp-library__success" role="status" aria-live="polite">
          {saveSuccess.value}
        </div>
      )}
    </form>
  );
}
