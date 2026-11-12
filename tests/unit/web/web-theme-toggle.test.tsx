/**
 * apps/web — the light / dark / follow-system control (ADR-0044).
 *
 * The behaviour worth testing here is not "a button sets an attribute". It is
 * the one asymmetry that makes the pre-paint CSS block work: "follow system"
 * must REMOVE `data-theme`, because web.css's `@media (prefers-color-scheme)`
 * block is scoped to `:root:not([data-theme])` and that block is what makes the
 * first paint correct before the deferred module script runs. Writing the
 * resolved value there looks identical in every screenshot and costs every
 * system-preference user a frame of the wrong theme on every load.
 *
 * Plus the two failure modes that are easy to ship: a localStorage that THROWS
 * (hardened profile, blocked site data) must not break the control, and the
 * resolved `colorScheme` must be a concrete value even when the preference is
 * not.
 */
import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { render } from 'preact';
import { act } from 'preact/test-utils';
import {
  ThemeToggle,
  THEME_STORAGE_KEY,
  applyTheme,
  readStoredPreference,
} from '../../../src/web/ThemeToggle.js';

let container: HTMLDivElement;

/** jsdom has no matchMedia; install one reporting the OS as `dark`. */
function stubMatchMedia(prefersDark: boolean): void {
  Object.defineProperty(window, 'matchMedia', {
    writable: true,
    configurable: true,
    value: (query: string) => ({
      matches: prefersDark && query.includes('dark'),
      media: query,
      addEventListener: () => {},
      removeEventListener: () => {},
    }),
  });
}

beforeEach(() => {
  container = document.createElement('div');
  document.body.appendChild(container);
  document.documentElement.removeAttribute('data-theme');
  document.documentElement.style.colorScheme = '';
  localStorage.clear();
  stubMatchMedia(false);
});

afterEach(() => {
  render(null, container);
  container.remove();
  document.documentElement.removeAttribute('data-theme');
  vi.restoreAllMocks();
});

function buttons(): HTMLButtonElement[] {
  return [...container.querySelectorAll<HTMLButtonElement>('.dp-web__theme-btn')];
}

function click(label: string): void {
  const btn = buttons().find(b => b.getAttribute('aria-label') === label);
  if (!btn) throw new Error(`no theme button labelled ${label}`);
  act(() => {
    btn.click();
  });
}

describe('ThemeToggle', () => {
  it('offers light, dark and follow-system as one radiogroup', () => {
    act(() => render(<ThemeToggle />, container));
    const group = container.querySelector('[role="radiogroup"]');
    expect(group?.getAttribute('aria-label')).toBe('Colour theme');
    expect(buttons().map(b => b.getAttribute('aria-label'))).toEqual([
      'Light',
      'Dark',
      'Follow system',
    ]);
    // Every control carries its own name: the icons are decorative.
    for (const b of buttons()) {
      expect(b.querySelector('svg')?.getAttribute('aria-hidden')).toBe('true');
    }
  });

  it('defaults to follow-system and leaves data-theme OFF', () => {
    act(() => render(<ThemeToggle />, container));
    expect(document.documentElement.hasAttribute('data-theme')).toBe(false);
    const system = buttons().find(b => b.getAttribute('aria-label') === 'Follow system');
    expect(system?.getAttribute('aria-checked')).toBe('true');
  });

  it('an explicit choice writes the attribute, the colour scheme and storage', () => {
    act(() => render(<ThemeToggle />, container));
    click('Dark');
    expect(document.documentElement.getAttribute('data-theme')).toBe('dark');
    expect(document.documentElement.style.colorScheme).toBe('dark');
    expect(localStorage.getItem(THEME_STORAGE_KEY)).toBe('dark');

    click('Light');
    expect(document.documentElement.getAttribute('data-theme')).toBe('light');
    expect(document.documentElement.style.colorScheme).toBe('light');
    expect(localStorage.getItem(THEME_STORAGE_KEY)).toBe('light');
  });

  it('RETURNING to follow-system removes the attribute rather than writing the resolved value', () => {
    // The load-bearing case. web.css's pre-paint block is scoped to
    // :root:not([data-theme]); writing 'dark' here would type-check, look
    // right on screen, and silently cost every system user a wrong first
    // paint on every load.
    stubMatchMedia(true);
    act(() => render(<ThemeToggle />, container));
    click('Light');
    expect(document.documentElement.getAttribute('data-theme')).toBe('light');

    click('Follow system');
    expect(document.documentElement.hasAttribute('data-theme')).toBe(false);
    // ...while colorScheme still carries the RESOLVED value, so the browser
    // paints form controls and scrollbars in the right mode.
    expect(document.documentElement.style.colorScheme).toBe('dark');
    expect(localStorage.getItem(THEME_STORAGE_KEY)).toBe('system');
  });

  it('restores a stored preference on mount', () => {
    localStorage.setItem(THEME_STORAGE_KEY, 'dark');
    act(() => render(<ThemeToggle />, container));
    expect(document.documentElement.getAttribute('data-theme')).toBe('dark');
    const dark = buttons().find(b => b.getAttribute('aria-label') === 'Dark');
    expect(dark?.getAttribute('aria-checked')).toBe('true');
  });

  it('ignores a stored value that is not one of the three', () => {
    localStorage.setItem(THEME_STORAGE_KEY, 'solarized');
    expect(readStoredPreference()).toBe('system');
  });

  it('still works when localStorage THROWS on read and on write', () => {
    // A hardened profile or blocked site data makes the accessor itself raise;
    // this is why the implementation is a try/catch and not a null check.
    const getItem = vi.spyOn(Storage.prototype, 'getItem').mockImplementation(() => {
      throw new Error('blocked');
    });
    const setItem = vi.spyOn(Storage.prototype, 'setItem').mockImplementation(() => {
      throw new Error('blocked');
    });

    expect(() => act(() => render(<ThemeToggle />, container))).not.toThrow();
    expect(() => click('Dark')).not.toThrow();
    // The preference is lost on reload — a papercut — but the control works.
    expect(document.documentElement.getAttribute('data-theme')).toBe('dark');

    getItem.mockRestore();
    setItem.mockRestore();
  });

  it('survives an environment with no matchMedia at all', () => {
    Object.defineProperty(window, 'matchMedia', {
      writable: true,
      configurable: true,
      value: undefined,
    });
    expect(() => act(() => render(<ThemeToggle />, container))).not.toThrow();
    expect(document.documentElement.style.colorScheme).toBe('light');
  });
});

describe('applyTheme', () => {
  it('is the single place the attribute rule lives, and the rule holds off-component', () => {
    applyTheme('dark', 'dark');
    expect(document.documentElement.getAttribute('data-theme')).toBe('dark');

    applyTheme('system', 'light');
    expect(document.documentElement.hasAttribute('data-theme')).toBe(false);
    expect(document.documentElement.style.colorScheme).toBe('light');
  });
});
