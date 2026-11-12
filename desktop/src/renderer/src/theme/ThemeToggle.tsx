import { Monitor, Moon, Sun } from 'lucide-react';
import { useTheme, type ThemePreference } from './ThemeProvider.js';

const OPTIONS: ReadonlyArray<{ value: ThemePreference; label: string; Icon: typeof Sun }> = [
  { value: 'light', label: 'Light', Icon: Sun },
  { value: 'dark', label: 'Dark', Icon: Moon },
  { value: 'system', label: 'Follow system', Icon: Monitor },
];

export function ThemeToggle() {
  const { preference, setPreference } = useTheme();

  return (
    // no-drag is mandatory: this sits inside the titlebar's drag region, which
    // swallows every pointer event from anything that does not opt out.
    <div
      className="no-drag flex items-center gap-0.5 rounded-md border border-border bg-surface-sunken p-0.5"
      role="radiogroup"
      aria-label="Colour theme"
    >
      {OPTIONS.map(({ value, label, Icon }) => {
        const active = preference === value;
        return (
          <button
            key={value}
            type="button"
            role="radio"
            aria-checked={active}
            aria-label={label}
            title={label}
            onClick={() => setPreference(value)}
            className={
              'flex h-6 w-6 items-center justify-center rounded transition-colors ' +
              (active
                ? 'bg-accent text-accent-text'
                : 'text-text-muted hover:bg-border hover:text-text')
            }
          >
            <Icon size={13} strokeWidth={2} aria-hidden />
          </button>
        );
      })}
    </div>
  );
}
