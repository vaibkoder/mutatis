/**
 * Feature 08 — Compare: the scrollable navigation list of changes (§8.3).
 *
 * ADR-0008: this list applies the filter toggles to DISPLAY ONLY. The `ops`
 * array passed in is never mutated; `applyFilters` returns a NEW filtered array.
 * Hidden changes are always one toggle away and always remain in the redline
 * document — nothing is suppressed.
 *
 * Lightweight rendering: a plain array render inside an overflow-y:auto
 * container. No heavy virtualisation library at this stage (§9.10 leaves that
 * to a later pass); the render must simply not crash on ~500 rows.
 */
import { useState } from 'preact/hooks';
import type { CascadeAnnotation, ChangeClass, DiffOp, PropertyChange } from '../model.js';
import { objectEdits } from '../diff/object-edits.js';
import { firstCellChange, tableChangeCount, tableShapeSummary } from '../diff/table-summary.js';

/** What a collapsed cascade header says, and what it calls its members. */
export interface CascadeLabel {
  /** The header sentence, e.g. "Clause references renumbered (+1)". */
  text: string;
  /** Singular noun for one member, e.g. "reference" / "paragraph". */
  memberNoun: string;
}

/**
 * Build a cascade-id → header label map for the collapsed cascade rows.
 * Shared by all three compare panels (Word-vs-Word, Text-vs-text, PDF).
 *
 * Two sentences, because they are two facts (ADR-0042). A `crossReference`
 * cascade is the document POINTING AT a clause whose number moved — its members
 * are references. A `numbering` cascade is the clause numbers themselves moving;
 * its members contain no reference to anything, so calling them references would
 * claim text the document does not have. `SummaryPanel` has drawn that
 * distinction since ADR-0042 landed and this map did not, so the navigation list
 * — the surface a reader actually works from — read "Clause numbering (level 1)
 * renumbered (+1) — 20 references" for twenty paragraphs containing none.
 */
export function cascadeLabelMap(cascades: CascadeAnnotation[]): Map<string, CascadeLabel> {
  const map = new Map<string, CascadeLabel>();
  for (const c of cascades) {
    const sign = c.shift > 0 ? `+${c.shift}` : `${c.shift}`;
    map.set(
      c.id,
      c.source === 'numbering'
        ? { text: `${c.tokenKind} shifted (${sign})`, memberNoun: 'paragraph' }
        : { text: `${c.tokenKind} references renumbered (${sign})`, memberNoun: 'reference' },
    );
  }
  return map;
}

// ---------------------------------------------------------------------------
// Filter model (pure — exported for tests; ADR-0008 guard)
// ---------------------------------------------------------------------------

export interface FilterFlags {
  hideFormattingOnly: boolean;
  collapseRenumbering: boolean;
}

/** Classes that the §7 always-substantive rules force to stay visible. */
function isAlwaysSubstantive(op: DiffOp): boolean {
  // Signature-block / table-cell / spelling-variant items are never filtered.
  if (op.spellingVariant) return true;
  // Table-cell ops carry a partId prefix of 'tableCell:'.
  if (op.partId.startsWith('tableCell:')) return true;
  // A changed Word table (ADR-0046) — the badge and the filter say the same.
  if (op.tableDiff !== undefined && op.type === 'replace') return true;
  return false;
}

/**
 * Apply the display filters to an op list. Returns a NEW array — the input is
 * never mutated (ADR-0008). Renumbering-collapse is handled by the caller's
 * grouping; this function only removes formatting-only / whitespace-punct rows
 * and (when collapse is on) the individual cascade members.
 *
 * @returns the ops that should be shown as top-level rows.
 */
export function applyFilters(ops: DiffOp[], flags: FilterFlags): DiffOp[] {
  // ADR-0008: filters affect display only — the input ops array is never mutated.
  return ops.filter(op => {
    if (isAlwaysSubstantive(op)) return true;

    if (flags.hideFormattingOnly && op.classification === 'formatting') return false;

    // When collapse is ON, individual cascade members are hidden from the
    // top-level list (they remain reachable by expanding the cascade row).
    if (flags.collapseRenumbering && op.collapsedInto !== undefined) return false;

    return true;
  });
}

/** Group cascade members (collapsedInto) by their cascade id, for expansion. */
export function groupCascadeMembers(ops: DiffOp[]): Map<string, DiffOp[]> {
  const map = new Map<string, DiffOp[]>();
  for (const op of ops) {
    if (op.collapsedInto !== undefined) {
      const list = map.get(op.collapsedInto) ?? [];
      list.push(op);
      map.set(op.collapsedInto, list);
    }
  }
  return map;
}

// ---------------------------------------------------------------------------
// Presentation helpers
// ---------------------------------------------------------------------------

const CLASS_TAG: Record<ChangeClass, string> = {
  structural: 'STRUCT',
  obligationLanguage: 'OBLIG',
  numericalValue: 'NUM',
  definedTerm: 'TERM',
  crossRefRenumber: 'XREF',
  substantiveText: 'TEXT',
  whitespacePunct: 'WS',
  formatting: 'FMT',
};

/** First 80 chars of the change's ins/del text, for the row preview. */
export function previewText(op: DiffOp): string {
  // A Word table op (ADR-0046) carries its changes in its cells: show the first
  // changed cell's own preview, or the table's shape when only rows or columns
  // moved — and say how many more there are, so one cell never reads as the
  // whole of the change.
  if (op.tableDiff !== undefined) {
    const inner = firstCellChange(op);
    const head = inner === undefined ? tableShapeSummary(op.tableDiff) : previewText(inner);
    const more = tableChangeCount(op.tableDiff) - 1;
    return inner === undefined || more <= 0 ? head : `${head} (+${more} more in this table)`;
  }
  if (op.wordDiffs && op.wordDiffs.length > 0) {
    const changed = op.wordDiffs
      .filter(w => w.op !== 'eq')
      .map(w => `${w.op === 'ins' ? '+' : '-'}${w.text}`)
      .join(' ');
    return truncate(changed, 80);
  }
  // A property op (ADR-0042) has no word diffs — its two paragraphs are
  // word-for-word identical — so it would otherwise render as a bare category
  // chip with no text, and a list of them would be indistinguishable rows.
  // Say what actually moved: the number first, since that is what the reader
  // sees on the page, then the style.
  const pc = op.propertyChange;
  if (pc) {
    const parts: string[] = [];
    // ADR-0045 § Amendment — a footnote, picture or object added, removed or
    // moved, words untouched: named first, since it is a change to the content.
    if (pc.objects === true) {
      const { removed, added } = objectEdits(pc);
      parts.push([...removed.map(x => `-[${x.sort}]`), ...added.map(x => `+[${x.sort}]`)].join(' '));
    }
    if (pc.numberLabel || pc.numberingIdentity) {
      // Fall back to the LEVEL when neither side resolved a label — otherwise a
      // real re-levelling renders as "— → —", which is the blank row again. And
      // a WORD for a label that resolved to the empty string: a level at
      // `numFmt="none"`, or with an empty `lvlText`, renders nothing, which is a
      // different fact from "no level resolved" and must not print as a gap.
      const side = (v: PropertyChange['from']): string => {
        if (v.numberLabel === null) return v.ilvl === null ? '—' : `level ${v.ilvl + 1}`;
        return v.numberLabel === '' ? '(no visible number)' : v.numberLabel;
      };
      let before = side(pc.from);
      let after = side(pc.to);
      // The identity can change while the rendered label stands still, and then
      // both sides read the same — "1.1 → 1.1", which is the blank-row artefact
      // again. Two different things cause it and they must not be conflated:
      //
      //  - the paragraph was RE-LEVELLED to a level that happens to carry the
      //    same label (an `ilvl` change with `%1.`-style templates), or
      //  - the list DEFINITION under it changed while its number did not
      //    (ADR-0042 § "Where the key still exceeds the label").
      //
      // `ilvl` is carried on both sides precisely so the first can be told from
      // the second; saying "list definition changed" for a re-levelling would be
      // a false statement about the document, not merely a vague one.
      if (before === after && pc.from.ilvl !== pc.to.ilvl) {
        const lvl = (v: PropertyChange['from']): string =>
          v.ilvl === null ? 'no level' : `level ${v.ilvl + 1}`;
        before = `${before} (${lvl(pc.from)})`;
        after = `${after} (${lvl(pc.to)})`;
      }
      parts.push(before === after ? `${before} (list definition changed)` : `${before} → ${after}`);
    }
    if (pc.styleId) parts.push(`style ${pc.from.styleId ?? '—'} → ${pc.to.styleId ?? '—'}`);
    return truncate(parts.join('  ·  '), 80);
  }
  return '';
}

function truncate(text: string, max: number): string {
  const collapsed = text.replace(/\s+/g, ' ').trim();
  return collapsed.length > max ? `${collapsed.slice(0, max)}…` : collapsed;
}

function badgeFor(op: DiffOp): string | null {
  // DiffOp carries spellingVariant and partId; signature-block provenance lives
  // on the underlying ParaRecord (not on DiffOp), so it surfaces via the
  // classifier forcing 'substantiveText' rather than via a derivable flag here.
  if (op.spellingVariant) return 'house-style';
  // PDF path (feature 12): a fuzzy-anchored match or approximate table-region
  // cell is low-confidence. ADR-0008 — this is a BADGE, never a suppression: the
  // row still renders; the badge just tells the reviewer to verify it.
  if (op.lowConfidence === true || op.fuzzyAnchored === true) return 'low confidence';
  if (op.partId.startsWith('tableCell:')) return 'always substantive';
  // A Word table op (ADR-0046) whose cells' words or rows changed — §7.3 forced
  // every text change inside a cell to substantive, so it says so like the
  // Markdown path's cells. (A 'property' table op changed nothing the reader
  // reads as text — a style, a number — and is badged like any property op.)
  if (op.tableDiff !== undefined && op.type === 'replace') return 'always substantive';
  return null;
}

// ---------------------------------------------------------------------------
// Components
// ---------------------------------------------------------------------------

/** Where a row's change is, when it is not the main text (ADR-0047). */
export interface RowLocation {
  /** The badge, e.g. "Footer" or "Footnote 3 (added)". */
  label: string;
  /** The badge's tooltip. */
  hint: string;
}

interface ChangeListProps {
  /** The full op list (never mutated). */
  ops: DiffOp[];
  filters: FilterFlags;
  /** Cascade id → header label (from the summary), for the collapsed row. */
  cascadeLabels?: Map<string, CascadeLabel>;
  /**
   * Part id → location, for the ops of a header, footer or note (ADR-0047).
   * An op whose `partId` is absent here is in the main text and has no badge.
   */
  locations?: ReadonlyMap<string, { location: RowLocation }>;
  onGoToChange: (opId: string) => void;
}

export function ChangeList({ ops, filters, cascadeLabels, locations, onGoToChange }: ChangeListProps) {
  const visible = applyFilters(ops, filters);
  const cascadeMembers = groupCascadeMembers(ops);

  return (
    <div
      class="dp-compare__change-list"
      role="list"
      style="overflow-y: auto; max-height: 60vh;"
    >
      {/* Collapsed cascade headers (only when collapse is ON and members exist). */}
      {filters.collapseRenumbering &&
        [...cascadeMembers.entries()].map(([cascadeId, members]) => (
          <CascadeRow
            key={cascadeId}
            cascadeId={cascadeId}
            label={
              cascadeLabels?.get(cascadeId) ?? {
                text: `Renumbering cascade (${members.length})`,
                memberNoun: 'change',
              }
            }
            members={members}
            {...(locations ? { locations } : {})}
            onGoToChange={onGoToChange}
          />
        ))}

      {visible.length === 0 && (
        <p class="dp-compare__change-list-empty">No changes match the current filters.</p>
      )}

      {visible.map(op => (
        <ChangeRow
          key={op.id} op={op} onGoToChange={onGoToChange}
          {...(locations?.get(op.partId) ? { location: locations.get(op.partId)!.location } : {})}
        />
      ))}
    </div>
  );
}

interface ChangeRowProps {
  op: DiffOp;
  location?: RowLocation;
  onGoToChange: (opId: string) => void;
}

function ChangeRow({ op, location, onGoToChange }: ChangeRowProps) {
  const badge = badgeFor(op);
  const preview = previewText(op);
  const lowConfidence = op.lowConfidence === true || op.fuzzyAnchored === true;
  return (
    <div
      class={
        `dp-compare__change-row dp-compare__change-row--${op.classification}` +
        (lowConfidence ? ' dp-compare__change-row--low-confidence' : '')
      }
      role="listitem"
      tabIndex={0}
      onClick={() => onGoToChange(op.id)}
      onKeyDown={e => {
        if (e.key === 'Enter' || e.key === ' ') onGoToChange(op.id);
      }}
    >
      <span class="dp-compare__change-tag">{CLASS_TAG[op.classification]}</span>
      {location && <span class="dp-compare__change-location" title={location.hint}>{location.label}</span>}
      {op.type === 'move' && <span class="dp-compare__change-move">moved</span>}
      {preview && <span class="dp-compare__change-preview">{preview}</span>}
      {badge && <span class="dp-compare__change-badge">{badge}</span>}
    </div>
  );
}

interface CascadeRowProps {
  cascadeId: string;
  label: CascadeLabel;
  members: DiffOp[];
  locations?: ReadonlyMap<string, { location: RowLocation }>;
  onGoToChange: (opId: string) => void;
}

function CascadeRow({ label, members, locations, onGoToChange }: CascadeRowProps) {
  const [expanded, setExpanded] = useState(false);
  return (
    <div class="dp-compare__cascade">
      <button
        type="button"
        class="dp-compare__cascade-header"
        aria-expanded={expanded ? 'true' : 'false'}
        onClick={() => setExpanded(v => !v)}
      >
        {expanded ? '▾' : '▸'} {label.text} — {members.length} {label.memberNoun}
        {members.length === 1 ? '' : 's'}
      </button>
      {expanded && (
        <div class="dp-compare__cascade-members" role="list">
          {members.map(op => (
            <ChangeRow
              key={op.id} op={op} onGoToChange={onGoToChange}
              {...(locations?.get(op.partId) ? { location: locations.get(op.partId)!.location } : {})}
            />
          ))}
        </div>
      )}
    </div>
  );
}
