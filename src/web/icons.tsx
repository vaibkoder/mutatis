/**
 * apps/web — the small icon set the web frame uses (ADR-0044).
 *
 * Hand-authored rather than pulled from an icon package: the standalone target
 * ships as a static, backendless bundle and every dependency is weight we carry
 * forever for five glyphs. They follow the same drawing conventions as the
 * desktop app's icons — a 24-unit grid, 2-unit strokes, round caps and joins,
 * `currentColor` so they inherit the token the surrounding text uses — so the
 * two products read as one family without either importing the other's assets.
 *
 * Every icon is decorative: it repeats a visible label or sits inside a control
 * that carries its own `aria-label`. They are therefore `aria-hidden` without
 * exception, and NONE of them may become the only carrier of meaning.
 */
import type { JSX } from 'preact';

interface IconProps {
  /** Rendered size in px. 13 is the frame's default, matching the label row. */
  size?: number;
}

function svg(size: number, children: JSX.Element): JSX.Element {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      stroke-width="2"
      stroke-linecap="round"
      stroke-linejoin="round"
      aria-hidden="true"
      focusable="false"
    >
      {children}
    </svg>
  );
}

/** A page with lines of text — the imported document pane. */
export function IconDocument({ size = 13 }: IconProps): JSX.Element {
  return svg(
    size,
    <>
      <path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8Z" />
      <path d="M14 3v5h5" />
      <path d="M9 13h6" />
      <path d="M9 17h6" />
    </>,
  );
}

/** A checked list — the findings pane. */
export function IconChecklist({ size = 13 }: IconProps): JSX.Element {
  return svg(
    size,
    <>
      <path d="m3 7 2 2 4-4" />
      <path d="m3 17 2 2 4-4" />
      <path d="M13 6h8" />
      <path d="M13 18h8" />
    </>,
  );
}

/** Sun — the explicit light preference. */
export function IconSun({ size = 13 }: IconProps): JSX.Element {
  return svg(
    size,
    <>
      <circle cx="12" cy="12" r="4" />
      <path d="M12 2v2" />
      <path d="M12 20v2" />
      <path d="M2 12h2" />
      <path d="M20 12h2" />
      <path d="m4.9 4.9 1.5 1.5" />
      <path d="m17.6 17.6 1.5 1.5" />
      <path d="m19.1 4.9-1.5 1.5" />
      <path d="m6.4 17.6-1.5 1.5" />
    </>,
  );
}

/** Crescent — the explicit dark preference. */
export function IconMoon({ size = 13 }: IconProps): JSX.Element {
  return svg(size, <path d="M20 14.5A8.5 8.5 0 0 1 9.5 4a8.5 8.5 0 1 0 10.5 10.5Z" />);
}

/** A display — "follow the operating system". */
export function IconSystem({ size = 13 }: IconProps): JSX.Element {
  return svg(
    size,
    <>
      <rect x="2" y="4" width="20" height="13" rx="2" />
      <path d="M8 21h8" />
      <path d="M12 17v4" />
    </>,
  );
}
