/**
 * E — AwarenessNotice (feature 11 — spec §6.3, ADR-0021 principle 2).
 *
 * The one-time, non-blocking toast. It must appear on the first copy/download,
 * never re-appear for the same browser profile (localStorage gate), auto-dismiss
 * after 8 s, and close early on ✕. `noticeVisible` is a module-level signal that
 * outlives `localStorage.clear()`, so we reset it via `dismissAwarenessNotice()`
 * in `beforeEach` (the harness clears localStorage for us).
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { render } from 'preact';
import { act } from 'preact/test-utils';
import { vi } from 'vitest';

import { AwarenessNotice } from '../../../../src/features/ai-export/ui/AwarenessNotice.js';
import {
  maybeShowAwarenessNotice, dismissAwarenessNotice,
} from '../../../../src/features/ai-export/ui/prefs.js';

let container: HTMLDivElement;

beforeEach(() => {
  dismissAwarenessNotice(); // reset the module-level signal (localStorage cleared by setup.ts)
  container = document.createElement('div');
  document.body.appendChild(container);
});

afterEach(() => {
  render(null, container);
  container.remove();
});

const notice = () => container.querySelector('.dp-ai-export__notice');

describe('E — AwarenessNotice', () => {
  it('renders nothing until an export action fires', () => {
    act(() => { render(<AwarenessNotice />, container); });
    expect(notice()).toBeNull();
  });

  it('appears on the first copy/download and shows the local-only message', () => {
    act(() => { render(<AwarenessNotice />, container); });
    act(() => { maybeShowAwarenessNotice(); });
    const el = notice();
    expect(el).not.toBeNull();
    expect(el?.getAttribute('role')).toBe('status');
    expect(el?.textContent).toContain('clipboard');
    expect(el?.textContent).toContain('no network calls');
  });

  it('does not reappear on a second action (localStorage gate)', () => {
    act(() => { render(<AwarenessNotice />, container); });
    act(() => { maybeShowAwarenessNotice(); }); // first
    expect(notice()).not.toBeNull();
    act(() => { dismissAwarenessNotice(); });   // user/auto dismiss
    expect(notice()).toBeNull();
    act(() => { maybeShowAwarenessNotice(); });  // second action — gated, no-op
    expect(notice()).toBeNull();
  });

  it('auto-dismisses after 8 seconds', () => {
    vi.useFakeTimers();
    try {
      act(() => { render(<AwarenessNotice />, container); });
      act(() => { maybeShowAwarenessNotice(); });
      expect(notice()).not.toBeNull();
      act(() => { vi.advanceTimersByTime(8000); });
      expect(notice()).toBeNull();
    } finally {
      vi.useRealTimers();
    }
  });

  it('closes early when the ✕ button is pressed', () => {
    act(() => { render(<AwarenessNotice />, container); });
    act(() => { maybeShowAwarenessNotice(); });
    const close = container.querySelector('.dp-ai-export__notice-close') as HTMLButtonElement;
    expect(close).not.toBeNull();
    act(() => { close.click(); });
    expect(notice()).toBeNull();
  });
});
