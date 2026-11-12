/**
 * Driver for Interactive Reference Lookup (feature 04).
 *
 * Orchestrates the full lookup pipeline (M0/M3/M5, spec §5.1–5.4):
 *  1. Build Section Index + Terms Dictionary once at task-pane open.
 *  2. Register DocumentSelectionChanged via selection.ts.
 *  3. On each settled selection: classify → resolve → publish to lookupResult signal.
 *  4. Subscribe DocModel.onInvalidate → set dirty + debounced dictionary rebuild (~1 s).
 *     A lookup during rebuild uses the stale dict but wraps the result in LookupUpdating.
 *
 * The Section Index is also rebuilt on invalidation (cheap — pure in-memory scan).
 */

import { signal } from '@preact/signals';
import { DocModel } from '@shared/doc-model/index.js';
import { buildSectionIndex } from '@shared/section-index/index.js';
import type { SectionIndex } from '@shared/section-index/types.js';
import { getActiveProfile, getLookupConfig } from '@shared/config/index.js';
import { detectCapabilities } from '@shared/office/index.js';
import type { LookupResult } from './types.js';
import { buildTermsDictionary } from './dictionary.js';
import { classify } from './classify.js';
import { resolve } from './resolve.js';
import {
  registerSelectionHandler,
  readSelectionWindow,
  nextSeq,
  isSuperseded,
  clearDebounce,
} from './selection.js';
import type { TermsDictionary } from './types.js';
import { promoteDefinition } from './definition-list.js';
import type { DefinitionListEntry } from './definition-list.js';

// ── Published state ───────────────────────────────────────────────────────────

/** Current lookup result. Null before the feature activates. */
export const lookupResult = signal<LookupResult | null>(null);

/** True when the feature is disabled (registration failed or host incompatible). */
export const lookupDisabled = signal(false);

/** Non-empty when the feature has a pane-level status message. */
export const lookupStatusMessage = signal<string | null>(null);

/** One-time banner when WordApi 1.3 is unavailable (auto-numbering degraded). */
export const lookupWordApi13Banner = signal<string | null>(null);

/**
 * Accumulating list of looked-up definitions (feature 04b, desktop only).
 * Most-recent first; deduped by termKey. Empty on Online and before first lookup.
 */
export const definitionList = signal<DefinitionListEntry[]>([]);

/** termKey of the active (expanded) definition card, or null. */
export const activeDefinitionKey = signal<string | null>(null);

/** True on a desktop host — gates the accumulating-list UI (feature 04b). */
export const isDesktopHost = signal(false);

/** Make a definition card the active (expanded) one. Pure UI — no Office calls. */
export function setActiveDefinition(termKey: string): void {
  activeDefinitionKey.value = termKey;
}

// ── Internal state ────────────────────────────────────────────────────────────

let _index: SectionIndex | null = null;
let _dict: TermsDictionary | null = null;
let _rebuildTimer: ReturnType<typeof setTimeout> | null = null;
let _disposeInvalidate: (() => void) | null = null;

// ── Dictionary rebuild ────────────────────────────────────────────────────────

function scheduleRebuild(): void {
  if (_dict) _dict.dirty = true;

  if (_rebuildTimer !== null) clearTimeout(_rebuildTimer);
  const { dictRebuildDebounceMs } = getLookupConfig(getActiveProfile());
  _rebuildTimer = setTimeout(() => {
    _rebuildTimer = null;
    rebuildNow();
  }, dictRebuildDebounceMs);
}

function rebuildNow(): void {
  const profile = getActiveProfile();
  const bodyParas = DocModel.paragraphs.value;
  const footnoteParas = DocModel.footnotes.value;
  const paras = DocModel.allParagraphs.value;

  _dict = buildTermsDictionary(bodyParas, footnoteParas, profile);
  _index = buildSectionIndex(paras, profile);
}

// ── Lookup pipeline ───────────────────────────────────────────────────────────

async function runLookup(): Promise<void> {
  if (lookupDisabled.value || !_index || !_dict) return;

  const mySeq = nextSeq();

  let windowText: string;
  let cursorOffset: number | undefined;
  let cursorParaId: string | undefined;

  try {
    await Word.run(async (ctx) => {
      const win = await readSelectionWindow(ctx);
      if (isSuperseded(mySeq)) return;

      windowText = win.text;
      cursorOffset = win.cursorOffset;

      // Get the cursor paragraph's text to derive the DocModel paragraph ID
      // for the proximity heuristic.  DocModel IDs are positional ("body-para-N"),
      // so we match by text against the in-memory snapshot.
      const sel = ctx.document.getSelection();
      const paras = sel.paragraphs;
      paras.load('items');
      await ctx.sync();
      const p = paras.items[0];
      if (p) {
        p.load('text');
        await ctx.sync();
        const paraText = (p.text ?? '').trim();
        if (paraText) {
          const allParas = DocModel.allParagraphs.value;
          const matched = allParas.find(n => n.text.trim() === paraText);
          cursorParaId = matched?.id;
        }
      }
    });
  } catch {
    if (isSuperseded(mySeq)) return;
    return; // Office error — silently skip
  }

  if (isSuperseded(mySeq)) return;

  const text = windowText!;
  if (!text) {
    if (!isSuperseded(mySeq)) lookupResult.value = { type: 'idle' };
    return;
  }

  const profile = getActiveProfile();
  const cls = classify(text, _dict!, cursorOffset);

  if (isSuperseded(mySeq)) return;

  const allParas = DocModel.allParagraphs.value;
  let result = resolve(cls, _index!, _dict!, allParas, profile, cursorParaId);

  // Wrap in Updating if dictionary was dirty when lookup ran.
  if (_dict!.dirty) {
    result = { type: 'updating', stale: result as Exclude<LookupResult, { type: 'updating' }> };
  }

  if (isSuperseded(mySeq)) return;
  lookupResult.value = result;

  // Feature 04b: on desktop, accumulate definition lookups into the side-panel list
  // and make the just-clicked term the active (expanded) card.
  if (isDesktopHost.value) {
    const shown = result.type === 'updating' ? result.stale : result;
    if (shown.type === 'definition' || shown.type === 'def-by-xref') {
      const term = shown.term;
      const entry: DefinitionListEntry = {
        termKey: term.termKey,
        display: term.display,
        result: shown,
      };
      definitionList.value = promoteDefinition(definitionList.value, entry);
      activeDefinitionKey.value = term.termKey;
    }
  }
}

// ── Activation ────────────────────────────────────────────────────────────────

/**
 * Activate the Interactive Lookup feature.
 * Call once after Office.onReady and after DocModel.load() completes.
 */
export async function activateLookup(): Promise<void> {
  // Check WordApi 1.3 availability; show degradation banner if absent (spec §8.8).
  const caps = detectCapabilities();
  isDesktopHost.value = caps.isDesktop;
  if (!caps.wordApi13) {
    lookupWordApi13Banner.value =
      'Some auto-numbered references may not resolve in this version of Word. ' +
      'Manually typed references and defined terms still work.';
  }

  // Build initial index + dictionary.
  rebuildNow();

  // Subscribe to DocModel invalidation for subsequent rebuilds.
  _disposeInvalidate = DocModel.onInvalidate(() => {
    scheduleRebuild();
  });

  // Register the selection handler.
  const status = await registerSelectionHandler(() => {
    runLookup().catch(() => { /* silently ignore errors in the lookup pipeline */ });
  });

  if (status === 'failed') {
    lookupDisabled.value = true;
    lookupStatusMessage.value =
      'Reference preview is unavailable: could not register with this version of Word.';
  }
}

/** Deactivate and clean up. */
export function deactivateLookup(): void {
  clearDebounce();
  if (_rebuildTimer !== null) { clearTimeout(_rebuildTimer); _rebuildTimer = null; }
  _disposeInvalidate?.();
  _disposeInvalidate = null;
  _index = null;
  _dict = null;
  lookupResult.value = null;
  lookupDisabled.value = false;
  lookupStatusMessage.value = null;
  definitionList.value = [];
  activeDefinitionKey.value = null;
}
