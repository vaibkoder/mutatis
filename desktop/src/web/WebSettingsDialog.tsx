import { Dialog } from '@base-ui/react/dialog';
import { Settings as SettingsIcon, X } from 'lucide-react';
import { useTheme, type ThemePreference } from '../renderer/src/theme/ThemeProvider.js';
import { useSettings, type Jurisdiction } from './WebSettingsProvider.js';

/**
 * The settings dialog for the single-file build: Jurisdiction and Appearance.
 * Model, API keys and Updates are Electron-only (no key store, no updater here).
 * Jurisdiction still changes what the redline reports, so its warning stays.
 */
export function WebSettingsDialog() {
  const { preference, setPreference } = useTheme();
  const { jurisdiction, setJurisdiction } = useSettings();

  return (
    <Dialog.Root>
      <Dialog.Trigger
        aria-label="Settings"
        title="Settings"
        className="no-drag flex h-7 w-7 items-center justify-center rounded text-text-muted transition-colors hover:bg-border hover:text-text"
      >
        <SettingsIcon size={14} aria-hidden />
      </Dialog.Trigger>

      <Dialog.Portal>
        <Dialog.Backdrop className="fixed inset-0 bg-black/40 transition-opacity duration-150 data-ending-style:opacity-0 data-starting-style:opacity-0" />
        <Dialog.Popup className="fixed top-1/2 left-1/2 flex max-h-[80vh] w-[30rem] max-w-[calc(100vw-3rem)] -translate-x-1/2 -translate-y-1/2 flex-col overflow-hidden rounded-lg border border-border bg-surface-raised shadow-xl transition-[scale,opacity] duration-100 ease-out data-ending-style:scale-[0.98] data-ending-style:opacity-0 data-starting-style:scale-[0.98] data-starting-style:opacity-0">
          <div className="flex shrink-0 items-center gap-2 border-b border-border px-4 py-2.5">
            <Dialog.Title className="text-[13px] font-semibold text-text">Settings</Dialog.Title>
            <Dialog.Close
              aria-label="Close settings"
              className="ml-auto flex h-6 w-6 items-center justify-center rounded text-text-muted transition-colors hover:bg-border hover:text-text"
            >
              <X size={13} aria-hidden />
            </Dialog.Close>
          </div>

          <div className="base-ui-disable-scrollbar flex flex-col gap-5 overflow-y-auto px-4 py-4">
            <Section
              title="Jurisdiction"
              hint="Selects the comparison profile the engine runs under. It changes what the redline reports, so switch it before you compare, not after."
            >
              <Choice
                value={jurisdiction}
                onChange={(value) => setJurisdiction(value as Jurisdiction)}
                options={[
                  { value: 'UK', label: 'United Kingdom' },
                  { value: 'SG', label: 'Singapore' },
                ]}
                label="Jurisdiction profile"
              />
            </Section>

            <Section title="Appearance">
              <Choice
                value={preference}
                onChange={(value) => setPreference(value as ThemePreference)}
                options={[
                  { value: 'light', label: 'Light' },
                  { value: 'dark', label: 'Dark' },
                  { value: 'system', label: 'Follow system' },
                ]}
                label="Theme"
              />
            </Section>
          </div>
        </Dialog.Popup>
      </Dialog.Portal>
    </Dialog.Root>
  );
}

function Section({
  title,
  hint,
  children,
}: {
  title: string;
  hint?: string;
  children: React.ReactNode;
}) {
  return (
    <section className="flex flex-col gap-1.5">
      <h3 className="text-[11px] font-semibold tracking-wide uppercase text-text-muted">
        {title}
      </h3>
      {hint && <p className="text-[11px] leading-relaxed text-text-faint">{hint}</p>}
      {children}
    </section>
  );
}

function Choice({
  value,
  onChange,
  options,
  label,
}: {
  value: string;
  onChange: (value: string) => void;
  options: { value: string; label: string }[];
  label: string;
}) {
  return (
    <div role="radiogroup" aria-label={label} className="flex gap-1.5">
      {options.map((option) => {
        const selected = option.value === value;
        return (
          <button
            key={option.value}
            type="button"
            role="radio"
            aria-checked={selected}
            onClick={() => onChange(option.value)}
            className={
              selected
                ? 'rounded border border-accent bg-accent-soft px-2.5 py-1 text-[11px] font-medium text-text'
                : 'rounded border border-border px-2.5 py-1 text-[11px] text-text-muted transition-colors hover:border-border-strong hover:text-text'
            }
          >
            {option.label}
          </button>
        );
      })}
    </div>
  );
}
