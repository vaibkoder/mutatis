/**
 * QuestionnaireForm — C3/C9.
 *
 * Renders one widget per variable type.  Validates required fields and
 * reg-number masks.  Day-first date picker (DD Month YYYY).  No US-isms.
 */

import type { TemplateVariable } from './types.js';
import type { AnswersMap } from './types.js';
import { validateRegNumber, validateFieldValue, getEntityTypeOptions } from './validators.js';
import { validateDate, formatDateDayFirst, getRegNumberPlaceholder, getRegNumberLabel } from './format.js';
import type { TemplateJurisdiction } from './types.js';

// ---------------------------------------------------------------------------
// Props
// ---------------------------------------------------------------------------

interface QuestionnaireFormProps {
  variables: TemplateVariable[];
  jurisdiction: TemplateJurisdiction;
  answers: AnswersMap;
  onChange: (tag: string, value: string | boolean) => void;
  /** If set, shows a validation error summary at the top. */
  validationErrors?: Record<string, string>;
}

// ---------------------------------------------------------------------------
// Widget helper components
// ---------------------------------------------------------------------------

interface FieldProps {
  variable: TemplateVariable;
  value: string | boolean | number;
  jurisdiction: TemplateJurisdiction;
  onChange: (tag: string, value: string | boolean) => void;
  error: string | undefined;
}

function FieldWidget({ variable, value, jurisdiction, onChange, error }: FieldProps) {
  const strVal = typeof value === 'boolean' ? String(value) : String(value ?? '');

  const handleChange = (e: Event) => {
    const target = e.target as HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement;
    if (variable.type === 'boolean') {
      onChange(variable.tag, (target as HTMLInputElement).checked);
    } else {
      onChange(variable.tag, target.value);
    }
  };

  const inputClass = `dp-tmpl-field${error ? ' dp-tmpl-field--error' : ''}`;

  switch (variable.type) {
    case 'boolean':
      return (
        <label class="dp-tmpl-checkbox">
          <input
            type="checkbox"
            checked={strVal === 'true'}
            onChange={handleChange}
          />
          {' '}{variable.label}
          {variable.required && <span class="dp-tmpl-required" aria-hidden="true"> *</span>}
        </label>
      );

    case 'longtext':
    case 'address':
      return (
        <textarea
          class={inputClass}
          rows={4}
          value={strVal}
          onInput={handleChange}
          aria-label={variable.label}
          required={variable.required}
        />
      );

    case 'select':
    case 'entityType': {
      const options =
        variable.type === 'entityType'
          ? getEntityTypeOptions(jurisdiction)
          : (variable.options ?? []);
      return (
        <select
          class={inputClass}
          value={strVal}
          onChange={handleChange}
          aria-label={variable.label}
          required={variable.required}
        >
          <option value="">— select —</option>
          {options.map(opt => (
            <option key={opt} value={opt}>{opt}</option>
          ))}
        </select>
      );
    }

    case 'date':
      return (
        <input
          type="date"
          class={inputClass}
          value={strVal}
          onInput={handleChange}
          aria-label={variable.label}
          required={variable.required}
          placeholder="YYYY-MM-DD (displayed as DD Month YYYY)"
        />
      );

    case 'regNumber': {
      const scheme = variable.regScheme ?? (jurisdiction === 'SG' ? 'SG_UEN' : 'UK_CH');
      return (
        <input
          type="text"
          class={inputClass}
          value={strVal}
          onInput={handleChange}
          aria-label={getRegNumberLabel(scheme)}
          placeholder={getRegNumberPlaceholder(scheme)}
          required={variable.required}
          maxLength={variable.validation?.maxLen}
        />
      );
    }

    case 'number':
      return (
        <input
          type="number"
          class={inputClass}
          value={strVal}
          onInput={handleChange}
          aria-label={variable.label}
          required={variable.required}
          min={variable.validation?.min}
          max={variable.validation?.max}
        />
      );

    default: // text
      return (
        <input
          type="text"
          class={inputClass}
          value={strVal}
          onInput={handleChange}
          aria-label={variable.label}
          required={variable.required}
          maxLength={variable.validation?.maxLen}
        />
      );
  }
}

// ---------------------------------------------------------------------------
// Exported validation helper (used by the panel before submitting)
// ---------------------------------------------------------------------------

/**
 * Validate all fields in the answers map against the variable definitions.
 * Returns a record of tag → error message; empty record = all valid.
 */
export function validateAnswers(
  variables: TemplateVariable[],
  answers: AnswersMap,
  jurisdiction: TemplateJurisdiction,
): Record<string, string> {
  const errors: Record<string, string> = {};

  for (const v of variables) {
    const rawValue = answers[v.tag];
    const strValue = typeof rawValue === 'boolean' ? String(rawValue) : String(rawValue ?? '');

    // Required check.
    const fieldErr = validateFieldValue(
      strValue,
      {
        required: v.required,
        ...(v.validation?.maxLen !== undefined ? { maxLen: v.validation.maxLen } : {}),
        ...(v.validation?.pattern !== undefined ? { pattern: v.validation.pattern } : {}),
        ...(v.validation?.min !== undefined ? { min: v.validation.min } : {}),
        ...(v.validation?.max !== undefined ? { max: v.validation.max } : {}),
      },
      v.label,
    );
    if (fieldErr) { errors[v.tag] = fieldErr; continue; }

    // Date validation.
    if (v.type === 'date' && strValue.trim()) {
      const dateErr = validateDate(strValue, v.label);
      if (dateErr) { errors[v.tag] = dateErr; continue; }
    }

    // Reg-number validation.
    if (v.type === 'regNumber' && strValue.trim()) {
      const scheme = v.regScheme ?? (jurisdiction === 'SG' ? 'SG_UEN' : 'UK_CH');
      const result = validateRegNumber(strValue, scheme, jurisdiction);
      if (!result.valid && result.error) {
        errors[v.tag] = result.error;
      }
    }
  }

  return errors;
}

// ---------------------------------------------------------------------------
// QuestionnaireForm component
// ---------------------------------------------------------------------------

export function QuestionnaireForm({
  variables,
  jurisdiction,
  answers,
  onChange,
  validationErrors,
}: QuestionnaireFormProps) {
  if (variables.length === 0) {
    return <p class="dp-tmpl-empty">No variables are defined in this template.</p>;
  }

  return (
    <div class="dp-tmpl-form" lang="en-GB">
      {validationErrors && Object.keys(validationErrors).length > 0 && (
        <div class="dp-tmpl-error-summary" role="alert">
          <strong>Please correct the following:</strong>
          <ul>
            {Object.values(validationErrors).map((err, i) => (
              <li key={i}>{err}</li>
            ))}
          </ul>
        </div>
      )}

      {variables.map(v => {
        const error = validationErrors?.[v.tag];
        return (
          <div key={v.tag} class={`dp-tmpl-field-group${error ? ' dp-tmpl-field-group--error' : ''}`}>
            {v.type !== 'boolean' && (
              <label class="dp-tmpl-label" htmlFor={`dp-tmpl-field-${v.tag}`}>
                {v.label}
                {v.required && (
                  <span class="dp-tmpl-required" aria-label="required"> *</span>
                )}
                {v.type === 'date' && (
                  <span class="dp-tmpl-hint"> (day first — DD Month YYYY)</span>
                )}
                {v.type === 'regNumber' && v.regScheme && (
                  <span class="dp-tmpl-hint"> ({getRegNumberLabel(v.regScheme)})</span>
                )}
              </label>
            )}

            <div id={`dp-tmpl-field-${v.tag}`}>
              <FieldWidget
                variable={v}
                value={answers[v.tag] ?? v.default ?? ''}
                jurisdiction={jurisdiction}
                onChange={onChange}
                error={error}
              />
            </div>

            {/* Show parsed day-first date as a preview for date fields */}
            {v.type === 'date' && answers[v.tag] && (() => {
              const formatted = formatDateDayFirst(String(answers[v.tag]));
              return formatted ? (
                <span class="dp-tmpl-date-preview">→ {formatted}</span>
              ) : null;
            })()}

            {error && (
              <span class="dp-tmpl-field-error" role="alert">{error}</span>
            )}
          </div>
        );
      })}
    </div>
  );
}
