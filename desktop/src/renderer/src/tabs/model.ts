import type { DocumentSet } from '../state/documents.js';

/**
 * The work pane's tab model.
 *
 * The landing page is itself a tab rather than a separate mode, so "start a new
 * comparison" never means "throw away the redline you were reading". It is
 * pinned and cannot be closed.
 */

export type TabId = string;

export interface LandingTab {
  readonly kind: 'landing';
  readonly id: TabId;
  readonly title: string;
}

export type ComparisonStatus =
  | { readonly state: 'running'; readonly progress: number; readonly step: string }
  | { readonly state: 'ready' }
  | { readonly state: 'failed'; readonly message: string }
  | { readonly state: 'cancelled' };

export interface RedlineTab {
  readonly kind: 'redline';
  readonly id: TabId;
  readonly title: string;
  readonly originalName: string;
  readonly revisedName: string;
  /** How the documents were read, when it matters (a PDF is rebuilt from layout). */
  readonly notes?: readonly string[];
  readonly status: ComparisonStatus;
  /**
   * Populated once the pipeline finishes. Typed as unknown here so the tab
   * model stays free of engine types; the redline view narrows it.
   */
  readonly result: unknown;
}

/**
 * The many-to-many confirmation step, as a tab of its own.
 *
 * It is a tab rather than a modal because confirming a pairing over fifty
 * documents is real work — the user must be able to leave it, look at something
 * else, and come back without losing it.
 */
export interface PairingTab {
  readonly kind: 'pairing';
  readonly id: TabId;
  readonly title: string;
  readonly status: 'building' | 'ready' | 'failed';
  /** A `PairingResult`. Typed loosely so the tab model stays engine-free. */
  readonly result: unknown;
  readonly message?: string;
}

/**
 * The deviation matrix, as a tab of its own.
 *
 * It lives on the RIGHT, beside the redlines, because it is engine output: every
 * number in it is counted, not generated. The assistant's narration of it lives
 * in the chat column, to the side. Putting the matrix inside a chat message
 * would put a computed artefact behind whatever the model chose to say about it.
 */
export interface MatrixTab {
  readonly kind: 'matrix';
  readonly id: TabId;
  readonly title: string;
  /** A `DeviationMatrix`. Typed loosely so the tab model stays engine-free. */
  readonly result: unknown;
}

/**
 * The assistant's own changes, redlined against the document it edited.
 *
 * A separate kind from `RedlineTab` because it answers a different question —
 * "what did the assistant just do to my document?" rather than "how do these two
 * documents differ?" — and because it carries the edited file, which the user can
 * save, and the list of edits the engine REFUSED to apply.
 */
export interface SelfRedlineTab {
  readonly kind: 'self-redline';
  readonly id: TabId;
  readonly title: string;
  readonly targetName: string;
  readonly status: ComparisonStatus;
  /** A `SelfRedlineOutcome`. Typed loosely so the tab model stays engine-free. */
  readonly result: unknown;
}

export type Tab = LandingTab | RedlineTab | PairingTab | MatrixTab | SelfRedlineTab;

export const LANDING_TAB: LandingTab = {
  kind: 'landing',
  id: 'landing',
  title: 'New comparison',
};

export interface TabState {
  readonly tabs: readonly Tab[];
  readonly activeId: TabId;
}

export const initialTabState: TabState = {
  tabs: [LANDING_TAB],
  activeId: LANDING_TAB.id,
};

let seq = 0;
export function nextTabId(): TabId {
  return `redline-${++seq}`;
}

export function nextPairingTabId(): TabId {
  return `pairing-${++seq}`;
}

export function nextMatrixTabId(): TabId {
  return `matrix-${++seq}`;
}

export function nextSelfRedlineTabId(): TabId {
  return `assistant-${++seq}`;
}

/** A concise label for a pair, trimming the extension and any shared prefix. */
export function pairTitle(originalName: string, revisedName: string): string {
  const strip = (n: string) => n.replace(/\.[^.]+$/, '');
  const a = strip(originalName);
  const b = strip(revisedName);
  return a === b ? a : `${a} → ${b}`;
}

export function addTab(state: TabState, tab: Tab): TabState {
  return { tabs: [...state.tabs, tab], activeId: tab.id };
}

export function updateTab(state: TabState, id: TabId, patch: Partial<RedlineTab>): TabState {
  return {
    ...state,
    tabs: state.tabs.map((t) =>
      t.id === id && t.kind === 'redline' ? ({ ...t, ...patch } as RedlineTab) : t,
    ),
  };
}

export function updateSelfRedlineTab(
  state: TabState,
  id: TabId,
  patch: Partial<SelfRedlineTab>,
): TabState {
  return {
    ...state,
    tabs: state.tabs.map((t) =>
      t.id === id && t.kind === 'self-redline' ? ({ ...t, ...patch } as SelfRedlineTab) : t,
    ),
  };
}

export function updatePairingTab(
  state: TabState,
  id: TabId,
  patch: Partial<PairingTab>,
): TabState {
  return {
    ...state,
    tabs: state.tabs.map((t) =>
      t.id === id && t.kind === 'pairing' ? ({ ...t, ...patch } as PairingTab) : t,
    ),
  };
}

export function closeTab(state: TabState, id: TabId): TabState {
  if (id === LANDING_TAB.id) return state; // pinned
  const index = state.tabs.findIndex((t) => t.id === id);
  if (index === -1) return state;

  const tabs = state.tabs.filter((t) => t.id !== id);
  if (state.activeId !== id) return { ...state, activeId: state.activeId };

  // Activate the neighbour to the left, which is where the eye already is.
  const fallback = tabs[Math.max(0, index - 1)] ?? LANDING_TAB;
  return { tabs, activeId: fallback.id };
}

export function selectTab(state: TabState, id: TabId): TabState {
  return state.tabs.some((t) => t.id === id) ? { ...state, activeId: id } : state;
}

export function activeTab(state: TabState): Tab {
  return state.tabs.find((t) => t.id === state.activeId) ?? LANDING_TAB;
}

/** Pairs to compare for the current selection, in the order they will run. */
export function pairsFor(documents: DocumentSet): ReadonlyArray<{
  original: DocumentSet['original'][number];
  revised: DocumentSet['modified'][number];
}> {
  const { original, modified } = documents;
  if (original.length === 0 || modified.length === 0) return [];

  // 1:1 and 1:N fan out from the single side. N:M is proposed, never assumed —
  // phase 3 replaces this branch with the pairing grid.
  if (original.length === 1) {
    return modified.map((revised) => ({ original: original[0]!, revised }));
  }
  if (modified.length === 1) {
    return original.map((orig) => ({ original: orig, revised: modified[0]! }));
  }
  return [];
}
