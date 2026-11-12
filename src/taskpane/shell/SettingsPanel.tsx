import { signal } from '@preact/signals';
import { getStoredJurisdiction, setStoredJurisdiction } from '@shared/config/index.js';
import { getCapabilities } from '@shared/office/index.js';
import type { JurisdictionId } from '@shared/config/index.js';

const _jurisdiction = signal<JurisdictionId>(getStoredJurisdiction());

export function SettingsPanel() {
  const jurisdiction = _jurisdiction.value;
  const caps = getCapabilities();

  function handleJurisdictionChange(e: Event) {
    const val = (e.target as HTMLSelectElement).value as JurisdictionId;
    setStoredJurisdiction(val);
    _jurisdiction.value = val;
  }

  return (
    <section
      class="dp-settings-panel"
      role="tabpanel"
      id="dp-panel-settings"
      aria-labelledby="dp-tab-settings"
    >
      <h2 class="dp-settings-panel__heading">Settings</h2>

      <fieldset class="dp-settings-section">
        <legend class="dp-settings-section__legend">Jurisdiction</legend>
        <label class="dp-settings-field" for="dp-jurisdiction-select">
          Active jurisdiction
        </label>
        <select
          id="dp-jurisdiction-select"
          class="dp-settings-field__control"
          value={jurisdiction}
          onChange={handleJurisdictionChange}
          aria-describedby="dp-jurisdiction-desc"
        >
          <option value="UK">England &amp; Wales</option>
          <option value="SG">Singapore</option>
        </select>
        <p id="dp-jurisdiction-desc" class="dp-settings-field__hint">
          Sets default terminology, currency, and date conventions for all checkers.
        </p>
      </fieldset>

      <fieldset class="dp-settings-section">
        <legend class="dp-settings-section__legend">Host capabilities</legend>
        <dl class="dp-capabilities-list">
          <dt>WordApi 1.5 (footnotes, TOC)</dt>
          <dd>{caps.wordApi15 ? '✓ Available' : '✕ Not available'}</dd>
          <dt>WordApi 1.6 (tracked changes)</dt>
          <dd>{caps.wordApi16 ? '✓ Available' : '✕ Not available'}</dd>
          <dt>OOXML loop (Repair, Compare)</dt>
          <dd>{caps.ooxmlLoop ? '✓ Available' : '✕ Not available'}</dd>
        </dl>
      </fieldset>
    </section>
  );
}
