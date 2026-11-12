/**
 * Feature 08 — Compare: the three filter checkboxes (§8.2).
 *
 * ADR-0008: toggling any of these changes ONLY the navigation-list visibility.
 * They never mutate driver.ops and never reach the output/author layer. The
 * authored redline document is already produced and is immutable here.
 */
import type { FilterSignals } from '../ui-types.js';

interface FilterTogglesProps {
  /** Any state carrying the three display-only filter signals (either compare mode). */
  driver: FilterSignals;
}

export function FilterToggles({ driver }: FilterTogglesProps) {
  // ADR-0008: filters affect display only — driver.ops is never mutated by a filter toggle.
  return (
    <fieldset class="dp-compare__filters">
      <legend class="dp-compare__filters-legend">Filters (list view only)</legend>

      <label class="dp-compare__filter">
        <input
          type="checkbox"
          checked={driver.hideFormattingOnly.value}
          onChange={e => {
            driver.hideFormattingOnly.value = (e.target as HTMLInputElement).checked;
          }}
        />
        Hide formatting-only
      </label>

      <label class="dp-compare__filter">
        <input
          type="checkbox"
          checked={driver.collapseRenumbering.value}
          onChange={e => {
            driver.collapseRenumbering.value = (e.target as HTMLInputElement).checked;
          }}
        />
        Collapse renumbering cascades
      </label>

      <p class="dp-compare__filters-note">
        Filters change only what is shown in the list. The redline document still
        contains every change.
      </p>
    </fieldset>
  );
}
