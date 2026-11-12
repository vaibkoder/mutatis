import { Dialog } from '@base-ui/react/dialog';
import { Settings as SettingsIcon, X } from 'lucide-react';
import type { ModelChoice } from '@common/ipc.js';
import { useTheme, type ThemePreference } from '../theme/ThemeProvider.js';
import { ProviderKeyRow } from './ProviderKeyRow.js';
import { useSettings, type Jurisdiction } from './SettingsProvider.js';
import { UpdatesSection } from './UpdatesSection.js';

/**
 * Settings.
 *
 * Four things live here, and they are genuinely different kinds of setting:
 * which model answers (a preference), which jurisdiction profile the COMPARE
 * ENGINE runs under (a correctness setting — it changes the redline), the API
 * keys (which the renderer can write and never read), and updates (which are
 * manual, and say so).
 *
 * The jurisdiction control carries its consequence in the text beside it. A
 * silent profile switch that changes what the engine reports as a change would
 * be the worst kind of hidden setting in a tool like this.
 */
export function SettingsDialog() {
  const { preference, setPreference } = useTheme();
  const {
    jurisdiction,
    setJurisdiction,
    model,
    setModel,
    snapshot,
    keys,
    loadFailed,
    saveKey,
  } = useSettings();

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

            <Section
              title="Model"
              hint="Which model answers in the chat panel. Providers without a key are listed but cannot be selected."
            >
              {snapshot ? (
                <select
                  aria-label="Assistant model"
                  value={`${model.provider}/${model.model}`}
                  onChange={(event) => {
                    const [provider, ...rest] = event.target.value.split('/');
                    if (provider === 'deepseek' || provider === 'alibaba') {
                      setModel({ provider, model: rest.join('/') } satisfies ModelChoice);
                    }
                  }}
                  className="w-full rounded border border-border bg-surface px-2 py-1.5 text-[12px] text-text outline-none focus:border-focus"
                >
                  {snapshot.providers.map((provider) => {
                    const configured = keys?.[provider.id].configured ?? false;
                    return (
                      <optgroup
                        key={provider.id}
                        label={configured ? provider.label : `${provider.label} — no key`}
                      >
                        {provider.models.map((entry) => (
                          <option
                            key={entry.id}
                            value={`${provider.id}/${entry.id}`}
                            disabled={!configured}
                          >
                            {entry.label}
                          </option>
                        ))}
                      </optgroup>
                    );
                  })}
                </select>
              ) : (
                <p className="text-[11px] text-text-muted">
                  {loadFailed ? 'The provider list could not be loaded.' : 'Loading…'}
                </p>
              )}
            </Section>

            <Section
              title="API keys"
              hint="Stored encrypted by the system keychain, in this app's main process. They are never readable from this window and are sent only to the provider you chose."
            >
              {snapshot?.encryptionAvailable === false && (
                <p className="rounded border border-danger/40 bg-danger/10 px-2 py-1.5 text-[11px] leading-relaxed text-danger">
                  The system keychain is unavailable, so no key can be stored. Mutatis will
                  not write one to disk in plain text.
                </p>
              )}
              <div className="flex flex-col gap-2">
                {snapshot?.providers.map((provider) => (
                  <ProviderKeyRow
                    key={provider.id}
                    provider={provider}
                    configured={keys?.[provider.id].configured ?? false}
                    mask={keys?.[provider.id].mask ?? ''}
                    encryptionAvailable={snapshot.encryptionAvailable}
                    onSave={saveKey}
                  />
                ))}
              </div>
            </Section>

            <Section
              title="Updates"
              hint="Mutatis never checks for an update on its own. It contacts nothing until you press the button below, and downloads nothing until you press a second one."
            >
              <UpdatesSection />
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
