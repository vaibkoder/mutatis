import type { LayoutStorage } from 'react-resizable-panels';

/**
 * localStorage, but it can never break the shell.
 *
 * The accessor itself throws in some contexts (a hardened profile, blocked site
 * data), not merely return null — so both halves are guarded. A lost panel
 * layout is a papercut; a shell that fails to mount is not.
 *
 * This works at all only because we serve the renderer from app:// rather than
 * file://: localStorage on a file:// origin is unreliable.
 */
export const layoutStorage: LayoutStorage = {
  getItem(key) {
    try {
      return localStorage.getItem(key);
    } catch {
      return null;
    }
  },
  setItem(key, value) {
    try {
      localStorage.setItem(key, value);
    } catch {
      /* layout simply will not persist */
    }
  },
};
