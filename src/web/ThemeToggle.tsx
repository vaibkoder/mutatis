/**
 * apps/web — the light / dark / follow-system control (ADR-0044).
 *
 * The task pane has no equivalent and must not grow one: it follows
 * `Office.context.officeTheme`, which is the user's own Word setting, and
 * ADR-0012 settled that. A browser tab has no such setting to follow, which is
 * the whole reason this exists here and nowhere else.
 *
 * THE ONE SUBTLETY, and it is the reason this file is not three lines:
 * "follow system" REMOVES the `data-theme` attribute rather than writing the
 * resolved value. web.css's pre-paint block is scoped to
 * `:root:not([data-theme])`, and that block is what makes the very first paint
 * correct for everyone who has not overridden anything — the module script is
 * deferred, so it cannot run before that paint. Writing `light`/`dark` for a
 * "system" user would take them out of that selector's reach and hand them a
 * frame of the wrong theme on every load.
 *
 * `colorScheme` is still written on every render, because it must carry the
 * RESOLVED value: it is what tells the browser to paint form controls,
 * scrollbars and the canvas behind the page in the right mode.
 */
import { useCallback, useEffect, useState } from 'preact/hooks';
import { IconMoon, IconSun, IconSystem } from './icons.js';

export type ThemePreference = 'light' | 'dark' | 'system';
type ResolvedTheme = 'light' | 'dark';

/** Namespaced so it cannot collide with anything else on a shared origin. */
export const THEME_STORAGE_KEY = 'mutatis.web.theme';

const DARK_QUERY = '(prefers-color-scheme: dark)';

const OPTIONS: ReadonlyArray<{
  value: ThemePreference;
  label: string;
  Icon: typeof IconSun;
}> = [
  { value: 'light', label: 'Light', Icon: IconSun },
  { value: 'dark', label: 'Dark', Icon: IconMoon },
  { value: 'system', label: 'Follow system', Icon: IconSystem },
];

/**
 * localStorage can THROW rather than return null — a hardened browser profile
 * or a blocked-site-data setting makes the accessor itself raise, so this is a
 * try/catch and not a null check.
 */
export function readStoredPreference(): ThemePreference {
  try {
    const raw = localStorage.getItem(THEME_STORAGE_KEY);
    if (raw === 'light' || raw === 'dark' || raw === 'system') return raw;
  } catch {
    /* fall through to the default */
  }
  return 'system';
}

/** `matchMedia` is absent in some test environments; absent means light. */
function systemTheme(): ResolvedTheme {
  try {
    return window.matchMedia?.(DARK_QUERY).matches ? 'dark' : 'light';
  } catch {
    return 'light';
  }
}

/**
 * Apply a preference to `<html>`. Exported so a caller can set the attribute
 * without mounting the control, and so the test can assert it directly.
 */
export function applyTheme(preference: ThemePreference, resolved: ResolvedTheme): void {
  const root = document.documentElement;
  if (preference === 'system') {
    root.removeAttribute('data-theme');
  } else {
    root.setAttribute('data-theme', preference);
  }
  root.style.colorScheme = resolved;
}

export function ThemeToggle() {
  const [preference, setPreference] = useState<ThemePreference>(readStoredPreference);
  const [systemResolved, setSystemResolved] = useState<ResolvedTheme>(systemTheme);

  // Track the OS setting so "system" stays live rather than sampling once at
  // mount — otherwise a user who flips their OS theme keeps our old one until
  // they reload.
  useEffect(() => {
    let mql: MediaQueryList | undefined;
    try {
      mql = window.matchMedia?.(DARK_QUERY);
    } catch {
      return;
    }
    if (!mql?.addEventListener) return;
    const onChange = () => setSystemResolved(mql.matches ? 'dark' : 'light');
    mql.addEventListener('change', onChange);
    return () => mql.removeEventListener('change', onChange);
  }, []);

  const resolved: ResolvedTheme = preference === 'system' ? systemResolved : preference;

  useEffect(() => {
    applyTheme(preference, resolved);
  }, [preference, resolved]);

  const choose = useCallback((next: ThemePreference) => {
    setPreference(next);
    try {
      localStorage.setItem(THEME_STORAGE_KEY, next);
    } catch {
      // The preference is lost on reload, which is a papercut, not a failure.
      // It must never break the control.
    }
  }, []);

  return (
    <div class="dp-web__theme" role="radiogroup" aria-label="Colour theme">
      {OPTIONS.map(({ value, label, Icon }) => (
        <button
          key={value}
          type="button"
          role="radio"
          aria-checked={preference === value ? 'true' : 'false'}
          aria-label={label}
          title={label}
          class="dp-web__theme-btn"
          onClick={() => choose(value)}
        >
          <Icon />
        </button>
      ))}
    </div>
  );
}
