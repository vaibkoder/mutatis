/**
 * Feature 06 — Numbering: task-pane panel (M4).
 *
 * Spec §7 layout:
 *   1. Scheme library (list + New/Duplicate/Import/Export buttons).
 *   2. Level editor (level rows with style/format/indent/restart + live preview).
 *   3. Apply scope + Apply button.
 *   4. Promote / Demote.
 *   5. Cross-reference review panel (View / Clear).
 *   6. Non-blocking banners (§7.6).
 */

import { useEffect, useRef } from 'preact/hooks';
import { validateScheme, previewLevel, duplicateScheme, BUILTIN_IDS } from './schema.js';
import type { NumberingLevel, ListNumberingStyle } from './types.js';
import {
  allSchemes,
  selectedSchemeId,
  selectedScheme,
  applyScope,
  lastApplyReport,
  operationRunning,
  reviewState,
  promoteDemoteState,
  banners,
  operationError,
  editingScheme,
  editingDirty,
  activateNumbering,
  runApply,
  runPromote,
  runDemote,
  runClearReview,
  selectScheme,
  saveCurrentEdit,
  runExport,
  runImport,
  updateBoundaryState,
} from './driver.js';
import { UpgradePanel } from './upgrade/UpgradePanel.js';

const LIST_NUMBERING_OPTIONS: { value: ListNumberingStyle; label: string }[] = [
  { value: 'arabic',      label: 'Arabic (1, 2, 3)' },
  { value: 'lowerLetter', label: 'Lower letter (a, b, c)' },
  { value: 'upperLetter', label: 'Upper letter (A, B, C)' },
  { value: 'lowerRoman',  label: 'Lower roman (i, ii, iii)' },
  { value: 'upperRoman',  label: 'Upper roman (I, II, III)' },
  { value: 'none',        label: 'None (bullet / unnumbered)' },
];

// ---------------------------------------------------------------------------
// SchemeLibrary
// ---------------------------------------------------------------------------

function SchemeLibrary() {
  const schemes = allSchemes.value;
  const selectedId = selectedSchemeId.value;

  // US preset hidden by default for SG/UK profiles (spec §4.6); never show suppressForSgUk.
  const visible = schemes.filter(s => !s.suppressForSgUk);

  function handleDuplicate() {
    const base = selectedScheme.value;
    if (!base) return;
    const copy = duplicateScheme(base, `user-${Date.now()}`);
    editingScheme.value = copy;
    editingDirty.value = true;
  }

  const fileInputRef = useRef<HTMLInputElement>(null);

  function handleImportClick() {
    fileInputRef.current?.click();
  }

  async function handleFileChange(e: Event) {
    const input = e.currentTarget as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;
    const errors = await runImport(file as Blob);
    if (errors.length > 0) {
      alert(`Import failed:\n${errors.join('\n')}`);
    }
    input.value = '';
  }

  return (
    <section class="dp-num-library" aria-label="Numbering scheme library">
      <header class="dp-num-library__header">
        <h2 class="dp-num-library__title">Schemes</h2>
        <div class="dp-num-library__actions">
          <button type="button" class="dp-num-btn dp-num-btn--small" onClick={handleDuplicate}
            title="Duplicate the selected scheme to create an editable copy">
            Duplicate
          </button>
          <button type="button" class="dp-num-btn dp-num-btn--small" onClick={handleImportClick}>
            Import
          </button>
          <button type="button" class="dp-num-btn dp-num-btn--small"
            onClick={() => void runExport()}>
            Export
          </button>
          <input ref={fileInputRef} type="file" accept=".json" class="dp-num-hidden"
            onChange={handleFileChange} aria-label="Import scheme file" />
        </div>
      </header>

      <ul class="dp-num-library__list" role="listbox" aria-label="Numbering schemes">
        {visible.map(s => (
          <li
            key={s.id}
            role="option"
            aria-selected={s.id === selectedId}
            class={`dp-num-library__item${s.id === selectedId ? ' dp-num-library__item--selected' : ''}`}
            onClick={() => void selectScheme(s.id)}
          >
            <span class="dp-num-library__item-name">{s.name}</span>
            {s.builtIn && <span class="dp-num-library__badge dp-num-library__badge--builtin">Built-in</span>}
            {s.jurisdiction === 'SG-UK' && <span class="dp-num-library__badge dp-num-library__badge--default">SG/UK</span>}
          </li>
        ))}
      </ul>
    </section>
  );
}

// ---------------------------------------------------------------------------
// LevelEditor
// ---------------------------------------------------------------------------

function LevelEditor() {
  const scheme = editingScheme.value ?? selectedScheme.value;
  if (!scheme) return <p class="dp-num-empty">Select a scheme to view its levels.</p>;

  const isEditable = editingScheme.value !== null && !editingScheme.value.builtIn;
  const isBuiltIn = !!scheme.builtIn || BUILTIN_IDS.has(scheme.id);

  const validation = validateScheme(scheme);

  function updateLevel(levelIndex: number, patch: Partial<NumberingLevel>) {
    if (!editingScheme.value) return;
    const newLevels = editingScheme.value.levels.map(l =>
      l.levelIndex === levelIndex ? { ...l, ...patch } : l,
    );
    editingScheme.value = { ...editingScheme.value, levels: newLevels };
    editingDirty.value = true;
  }

  return (
    <section class="dp-num-editor" aria-label="Level editor">
      <header class="dp-num-editor__header">
        <h2 class="dp-num-editor__title">
          {isEditable ? (
            <input
              class="dp-num-editor__name-input"
              value={editingScheme.value?.name ?? scheme.name}
              onInput={e => {
                if (!editingScheme.value) return;
                editingScheme.value = { ...editingScheme.value, name: (e.currentTarget as HTMLInputElement).value };
                editingDirty.value = true;
              }}
              aria-label="Scheme name"
            />
          ) : (
            <span>{scheme.name}</span>
          )}
        </h2>
        {isBuiltIn && !isEditable && (
          <p class="dp-num-editor__builtin-note">
            Built-in scheme is read-only. <button type="button" class="dp-num-link" onClick={() => {
              const copy = duplicateScheme(scheme, `user-${Date.now()}`);
              editingScheme.value = copy;
              editingDirty.value = true;
            }}>Duplicate to edit</button>
          </p>
        )}
        {isEditable && (
          <div class="dp-num-editor__save-row">
            <button
              type="button"
              class="dp-num-btn dp-num-btn--primary"
              disabled={!editingDirty.value || !validation.valid}
              onClick={() => void saveCurrentEdit()}
            >
              Save scheme
            </button>
            {!validation.valid && (
              <ul class="dp-num-editor__errors" role="alert">
                {validation.errors.map((e, i) => <li key={i}>{e}</li>)}
              </ul>
            )}
          </div>
        )}
      </header>

      <table class="dp-num-levels" aria-label="Numbering levels">
        <thead>
          <tr>
            <th>Level</th>
            <th>Style</th>
            <th>Format</th>
            <th>Prefix</th>
            <th>Suffix</th>
            <th title="Text indent (points)">Indent (pt)</th>
            <th title="Number indent (points)">Num (pt)</th>
            <th title="Restart after level (-1 = never restart)">Restart</th>
            <th>Preview</th>
          </tr>
        </thead>
        <tbody>
          {[...scheme.levels].sort((a, b) => a.levelIndex - b.levelIndex).map(level => (
            <tr key={level.levelIndex} class="dp-num-levels__row">
              <td class="dp-num-levels__idx">L{level.levelIndex}</td>
              <td class="dp-num-levels__style">
                {isEditable ? (
                  <select
                    value={level.listNumbering}
                    onChange={e => updateLevel(level.levelIndex, {
                      listNumbering: (e.currentTarget as HTMLSelectElement).value as ListNumberingStyle,
                    })}
                    aria-label={`Level ${level.levelIndex} numbering style`}
                  >
                    {LIST_NUMBERING_OPTIONS.map(o => (
                      <option key={o.value} value={o.value}>{o.label}</option>
                    ))}
                  </select>
                ) : (
                  <span>{level.listNumbering}</span>
                )}
              </td>
              <td class="dp-num-levels__format">
                {isEditable ? (
                  <input
                    class="dp-num-levels__format-input"
                    value={level.formatString}
                    onInput={e => updateLevel(level.levelIndex, {
                      formatString: (e.currentTarget as HTMLInputElement).value,
                    })}
                    aria-label={`Level ${level.levelIndex} format string — use %1…%9 tokens`}
                    title="Format string: use %N tokens where N is the 1-based level number (e.g. %1.%2. for compound decimal, (%3) for paren-letter)"
                  />
                ) : (
                  <code>{level.formatString}</code>
                )}
              </td>
              <td class="dp-num-levels__affix">
                {isEditable ? (
                  <input class="dp-num-levels__affix-input" value={level.prefix}
                    onInput={e => updateLevel(level.levelIndex, { prefix: (e.currentTarget as HTMLInputElement).value })}
                    aria-label={`Level ${level.levelIndex} prefix`} />
                ) : (
                  <code>{level.prefix || '—'}</code>
                )}
              </td>
              <td class="dp-num-levels__affix">
                {isEditable ? (
                  <input class="dp-num-levels__affix-input" value={level.suffix}
                    onInput={e => updateLevel(level.levelIndex, { suffix: (e.currentTarget as HTMLInputElement).value })}
                    aria-label={`Level ${level.levelIndex} suffix`} />
                ) : (
                  <code>{level.suffix || '—'}</code>
                )}
              </td>
              <td class="dp-num-levels__indent">
                {isEditable ? (
                  <input type="number" class="dp-num-levels__indent-input"
                    value={level.indent.textIndentPt} min={0} max={576} step={6}
                    onInput={e => updateLevel(level.levelIndex, {
                      indent: { ...level.indent, textIndentPt: Number((e.currentTarget as HTMLInputElement).value) },
                    })}
                    aria-label={`Level ${level.levelIndex} text indent in points`} />
                ) : (
                  <span>{level.indent.textIndentPt}</span>
                )}
              </td>
              <td class="dp-num-levels__indent">
                {isEditable ? (
                  <input type="number" class="dp-num-levels__indent-input"
                    value={level.indent.numberIndentPt} min={0} max={576} step={6}
                    onInput={e => updateLevel(level.levelIndex, {
                      indent: { ...level.indent, numberIndentPt: Number((e.currentTarget as HTMLInputElement).value) },
                    })}
                    aria-label={`Level ${level.levelIndex} number indent in points`} />
                ) : (
                  <span>{level.indent.numberIndentPt}</span>
                )}
              </td>
              <td class="dp-num-levels__restart">
                {isEditable ? (
                  <select
                    value={level.restartAfterLevel ?? -2}
                    onChange={e => {
                      const v = Number((e.currentTarget as HTMLSelectElement).value);
                      updateLevel(level.levelIndex, { restartAfterLevel: v === -2 ? null : v });
                    }}
                    aria-label={`Level ${level.levelIndex} restart after level`}
                  >
                    <option value={-2}>Default</option>
                    <option value={-1}>Never (continuous)</option>
                    {Array.from({ length: level.levelIndex }, (_, i) => (
                      <option key={i} value={i}>After level {i}</option>
                    ))}
                  </select>
                ) : (
                  <span>{level.restartAfterLevel === -1 ? '∞' : level.restartAfterLevel === null ? 'default' : `after L${level.restartAfterLevel}`}</span>
                )}
              </td>
              <td class="dp-num-levels__preview">
                <code>{previewLevel(scheme, level.levelIndex)}</code>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </section>
  );
}

// ---------------------------------------------------------------------------
// ApplyControls
// ---------------------------------------------------------------------------

function ApplyControls() {
  const running = operationRunning.value;
  const report = lastApplyReport.value;
  const scope = applyScope.value;
  const err = operationError.value;

  return (
    <section class="dp-num-apply" aria-label="Apply scheme">
      <h2 class="dp-num-apply__title">Apply</h2>
      <fieldset class="dp-num-apply__scope">
        <legend class="dp-sr-only">Apply scope</legend>
        <label>
          <input type="radio" name="applyScope" value="selection" checked={scope === 'selection'}
            onChange={() => { applyScope.value = 'selection'; }} />
          Selection
        </label>
        <label>
          <input type="radio" name="applyScope" value="document" checked={scope === 'document'}
            onChange={() => { applyScope.value = 'document'; }} />
          Whole document
        </label>
      </fieldset>
      <button
        type="button"
        class="dp-num-btn dp-num-btn--primary"
        disabled={running}
        aria-busy={running ? 'true' : 'false'}
        onClick={() => void runApply()}
      >
        {running ? 'Applying…' : 'Apply scheme'}
      </button>
      {err && <p class="dp-num-error" role="alert">{err}</p>}
      {report && (
        <p class="dp-num-apply__report" role="status">
          Applied to {report.listsAffected} list(s).
          {report.staticParasSkipped > 0 && ` ${report.staticParasSkipped} static paragraph(s) skipped.`}
          {report.tableParasSkipped > 0 && ` ${report.tableParasSkipped} table paragraph(s) excluded.`}
          {report.degradedLevels.length > 0 && ` Levels ${report.degradedLevels.join(', ')} partially applied.`}
        </p>
      )}
    </section>
  );
}

// ---------------------------------------------------------------------------
// PromoteDemoteControls
// ---------------------------------------------------------------------------

function PromoteDemoteControls() {
  const { canPromote, canDemote } = promoteDemoteState.value;
  const running = operationRunning.value;

  return (
    <section class="dp-num-levels-ctrl" aria-label="Promote and demote">
      <h2 class="dp-num-levels-ctrl__title">Level</h2>
      <div class="dp-num-levels-ctrl__buttons">
        <button
          type="button"
          class="dp-num-btn"
          disabled={!canPromote || running}
          onClick={() => void runPromote()}
          aria-label="Promote selection one level"
        >
          ◄ Promote
        </button>
        <button
          type="button"
          class="dp-num-btn"
          disabled={!canDemote || running}
          onClick={() => void runDemote()}
          aria-label="Demote selection one level"
        >
          Demote ►
        </button>
      </div>
    </section>
  );
}

// ---------------------------------------------------------------------------
// ReviewPanel
// ---------------------------------------------------------------------------

function ReviewPanel() {
  const review = reviewState.value;
  if (!review.active) return null;

  return (
    <section class="dp-num-review" aria-label="Cross-reference review">
      <header class="dp-num-review__header">
        <h2 class="dp-num-review__title">
          References to review
          <span class="dp-num-review__count">{review.count}</span>
        </h2>
        <button type="button" class="dp-num-btn dp-num-btn--small"
          onClick={() => void runClearReview()}
          aria-label="Clear review highlights">
          Clear highlights
        </button>
      </header>
      {review.items.length > 0 && (
        <ul class="dp-num-review__list" aria-label="Highlighted references">
          {review.items.slice(0, 50).map((item, i) => (
            <li key={i} class="dp-num-review__item">
              <code class="dp-num-review__ref">{item.text}</code>
              {item.context && item.context !== item.text && (
                <span class="dp-num-review__context">…{item.context}…</span>
              )}
            </li>
          ))}
          {review.items.length > 50 && (
            <li class="dp-num-review__overflow">
              + {review.items.length - 50} more references
            </li>
          )}
        </ul>
      )}
      <p class="dp-num-review__note">
        Numbering changed — verify highlighted references manually. No text was rewritten.
      </p>
    </section>
  );
}

// ---------------------------------------------------------------------------
// Banners
// ---------------------------------------------------------------------------

function Banners() {
  const msgs = banners.value;
  if (msgs.length === 0) return null;
  return (
    <aside class="dp-num-banners" role="note">
      {msgs.map((msg, i) => (
        <p key={i} class="dp-num-banner">{msg}</p>
      ))}
    </aside>
  );
}

// ---------------------------------------------------------------------------
// NumberingPanel (root)
// ---------------------------------------------------------------------------

export function NumberingPanel() {
  useEffect(() => {
    void activateNumbering();
    updateBoundaryState();
  }, []);

  return (
    <div class="dp-num-panel" lang="en-GB">
      <Banners />
      <SchemeLibrary />
      <LevelEditor />
      <ApplyControls />
      <PromoteDemoteControls />
      <UpgradePanel />
      <ReviewPanel />
    </div>
  );
}

/** Skeleton shown while the panel loads. */
export function NumberingPanelSkeleton() {
  return (
    <div class="dp-num-panel dp-num-panel--loading" aria-busy="true">
      <p class="dp-sr-only">Loading numbering schemes…</p>
    </div>
  );
}
