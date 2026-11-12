import { AlertTriangle, ArrowRight, Ban, Check, Info } from 'lucide-react';
import type {
  PairingDocument,
  PairingProposal,
  PairingRow,
} from '@features/compare/batch/pairing/model.js';
import type { PairScore } from '@features/compare/batch/pairing/score.js';

/**
 * The pairing confirmation grid.
 *
 * The rule this whole screen exists to enforce: **a pairing is proposed, never
 * silent.** Every document from both sides appears exactly once. A confidence
 * band is a badge — it changes how a row is presented, never whether it is
 * shown. Rivals are always visible on anything not auto-paired, because a
 * proposal whose alternatives you cannot see is one you can only rubber-stamp.
 *
 * A pair that cannot actually run still gets its row, with the reason named.
 * Dropping it would leave the user with a complete-looking batch quietly
 * missing a document.
 */

interface PairingGridProps {
  proposal: PairingProposal;
  onRepoint: (rowId: string, b: string | null) => void;
  onConfirm: () => void;
  busy: boolean;
}

export function PairingGrid({ proposal, onRepoint, onConfirm, busy }: PairingGridProps) {
  const { totals, rows, documents } = proposal;
  const nameOf = (docId: string | null): string =>
    docId === null ? '' : (documents.find((d) => d.docId === docId)?.fileName ?? docId);

  const runnablePairs = rows.filter((r) => r.band !== 'unmatched' && r.runnable).length;

  return (
    <div className="flex h-full flex-col">
      <header className="shrink-0 border-b border-border bg-surface-raised px-4 py-3">
        <h2 className="text-[13px] font-semibold">Confirm how these documents pair up</h2>
        <p className="mt-0.5 text-[11px] leading-relaxed text-text-muted">
          Nothing runs until you confirm. Every document appears below exactly once —
          change any pairing you disagree with.
        </p>

        <div className="mt-2 flex flex-wrap items-center gap-x-3 gap-y-1 text-[11px]">
          <Stat label="pairs" value={totals.pairs} />
          <Stat label="auto-paired" value={totals.autoPair} tone="ok" />
          <Stat label="need review" value={totals.needsReview} tone="warn" />
          <Stat label="unmatched" value={totals.unmatchedA + totals.unmatchedB} />
          {totals.unrunnable > 0 ? (
            <Stat label="cannot run" value={totals.unrunnable} tone="bad" />
          ) : null}
        </div>

        {proposal.disclosures.length > 0 ? (
          <ul className="mt-2 flex flex-col gap-0.5">
            {proposal.disclosures.map((d) => (
              <li key={d} className="flex items-start gap-1.5 text-[11px] text-text-muted">
                <Info size={12} className="mt-0.5 shrink-0" aria-hidden />
                <span>{d}</span>
              </li>
            ))}
          </ul>
        ) : null}
      </header>

      <div className="min-h-0 flex-1 overflow-y-auto">
        <ul className="divide-y divide-border">
          {rows.map((row) => (
            <Row
              key={row.id}
              row={row}
              documents={documents}
              nameOf={nameOf}
              onRepoint={onRepoint}
            />
          ))}
        </ul>
      </div>

      <footer className="flex h-12 shrink-0 items-center gap-3 border-t border-border bg-surface-raised px-4">
        <span className="text-[12px] text-text-muted">
          {runnablePairs} {runnablePairs === 1 ? 'pair' : 'pairs'} will be compared
          {totals.unrunnable > 0 ? `, ${totals.unrunnable} cannot run` : ''}
        </span>
        <button
          type="button"
          onClick={onConfirm}
          disabled={busy || runnablePairs === 0}
          className={
            'ml-auto flex items-center gap-1.5 rounded-md px-3 py-1.5 text-[12px] font-medium transition-colors ' +
            (busy || runnablePairs === 0
              ? 'cursor-not-allowed bg-border text-text-faint'
              : 'bg-accent text-accent-text hover:opacity-90')
          }
        >
          Confirm and compare
          <ArrowRight size={13} aria-hidden />
        </button>
      </footer>
    </div>
  );
}

function Stat({
  label,
  value,
  tone,
}: {
  label: string;
  value: number;
  tone?: 'ok' | 'warn' | 'bad';
}) {
  const colour =
    tone === 'ok'
      ? 'text-redline-ins'
      : tone === 'warn'
        ? 'text-[var(--redline-move)]'
        : tone === 'bad'
          ? 'text-danger'
          : 'text-text';
  return (
    <span className="text-text-muted">
      <span className={`tabular-nums font-medium ${colour}`}>{value}</span> {label}
    </span>
  );
}

function Row({
  row,
  documents,
  nameOf,
  onRepoint,
}: {
  row: PairingRow;
  documents: readonly PairingDocument[];
  nameOf: (docId: string | null) => string;
  onRepoint: (rowId: string, b: string | null) => void;
}) {
  // An unreadable document is not a pairing candidate — offering it in the
  // re-point control would let a user destroy a correct pairing by choosing
  // something the solver can never honour.
  const candidates = documents.filter((d) => d.side === 'b' && d.unreadable === undefined);
  const canRepoint = row.a !== null;

  return (
    <li className="px-4 py-2.5">
      <div className="flex items-start gap-3">
        <BandBadge row={row} />

        <div className="min-w-0 flex-1">
          <div className="flex flex-wrap items-baseline gap-x-2 text-[12px]">
            <span className="font-medium">{row.a !== null ? nameOf(row.a) : '—'}</span>
            <ArrowRight size={11} className="text-text-faint" aria-hidden />
            <span className={row.b !== null ? 'font-medium' : 'italic text-text-faint'}>
              {row.b !== null ? nameOf(row.b) : 'no counterpart'}
            </span>
            {row.score ? <ScoreChip score={row.score} /> : null}
          </div>

          {!row.runnable && row.reason ? (
            <p className="mt-1 flex items-start gap-1.5 text-[11px] leading-snug text-danger">
              <Ban size={12} className="mt-px shrink-0" aria-hidden />
              <span>{row.reason}</span>
            </p>
          ) : null}

          {/* Rivals are shown on anything that is not auto-paired. This is the
              difference between confirming a proposal and rubber-stamping it. */}
          {row.band !== 'auto-pair' && row.rivals.length > 0 ? (
            <div className="mt-1 flex flex-wrap items-baseline gap-x-2 gap-y-0.5 text-[11px] text-text-muted">
              <span className="text-text-faint">Also considered:</span>
              {row.rivals.map((rival) => (
                <button
                  key={rival.b}
                  type="button"
                  onClick={() => onRepoint(row.id, rival.b)}
                  className="rounded border border-border px-1.5 py-0.5 transition-colors hover:border-border-strong hover:text-text"
                  title={`Re-point to ${nameOf(rival.b)}`}
                >
                  {nameOf(rival.b)}{' '}
                  <span className="tabular-nums text-text-faint">
                    {(rival.score * 100).toFixed(0)}%
                  </span>
                </button>
              ))}
            </div>
          ) : null}
        </div>

        {canRepoint ? (
          <label className="shrink-0 text-[11px]">
            <span className="sr-only">Pair {nameOf(row.a)} with</span>
            <select
              value={row.b ?? ''}
              onChange={(e) => onRepoint(row.id, e.target.value === '' ? null : e.target.value)}
              className="rounded border border-border bg-surface px-1.5 py-1 text-[11px] text-text"
            >
              <option value="">Leave unmatched</option>
              {candidates.map((doc) => (
                <option key={doc.docId} value={doc.docId}>
                  {doc.fileName}
                </option>
              ))}
            </select>
          </label>
        ) : null}
      </div>
    </li>
  );
}

function BandBadge({ row }: { row: PairingRow }) {
  if (row.band === 'auto-pair') {
    return (
      <span
        title="High score with a clear lead over the next candidate. Still overridable."
        className="mt-0.5 flex h-5 w-5 shrink-0 items-center justify-center rounded-full bg-[var(--redline-ins-bg)] text-redline-ins"
      >
        <Check size={12} aria-hidden />
      </span>
    );
  }
  if (row.band === 'propose-with-rivals') {
    return (
      <span
        title="A pairing was found, but it is not strong or a rival is close behind. Worth a look."
        className="mt-0.5 flex h-5 w-5 shrink-0 items-center justify-center rounded-full bg-[var(--redline-move-bg)] text-[var(--redline-move)]"
      >
        <AlertTriangle size={11} aria-hidden />
      </span>
    );
  }
  return (
    <span
      title="No counterpart worth pairing with — an added or removed document. That is a real answer, not a failure."
      className="mt-0.5 flex h-5 w-5 shrink-0 items-center justify-center rounded-full bg-border text-text-faint"
    >
      <span aria-hidden className="text-[13px] leading-none">·</span>
    </span>
  );
}

function ScoreChip({ score }: { score: PairScore }) {
  // contentAvailable false means the content signal could not be measured at
  // all, not that the documents differ. Saying which is the whole point of the
  // engine carrying the flag.
  const title = score.contentAvailable
    ? `content ${(score.content * 100).toFixed(0)}% · filename ${(score.filename * 100).toFixed(0)}%`
    : `filename only — neither document had comparable text (filename ${(score.filename * 100).toFixed(0)}%)`;

  return (
    <span
      title={title}
      className="rounded bg-border/60 px-1.5 py-0.5 text-[10px] tabular-nums text-text-muted"
    >
      {(score.score * 100).toFixed(0)}%
      {!score.contentAvailable ? ' filename only' : ''}
    </span>
  );
}
