/**
 * The two blocks both batch surfaces render below their inputs (spec §5.4,
 * §5.5, §8 case 7, §10).
 *
 * Shared rather than written twice. They WERE written twice — the many-to-many
 * surface hand-rolled its own row line — and the copy drifted on exactly the
 * points that matter: it lost the classification chips, printed the raw status
 * enum, never said a batch had been cancelled, and never stated the flatten
 * choice its redlines were produced under. None of those omissions were
 * decisions; they were a second implementation of the same list, quietly
 * falling behind the first. Reuse-before-build, applied to
 * the place a review found it broken.
 */
import type { FlattenPolicy } from '@shared/config/index.js';
import type { ChangeClass } from '../../model.js';
import type { BatchIndex, BatchRow, BatchRowStatus } from '../model.js';
import { isDownloadable } from '../artefacts.js';

const STATUS_LABEL: Record<BatchRowStatus, string> = {
  queued: 'Queued',
  running: 'Comparing…',
  done: 'Compared',
  failed: 'Could not compare',
  cancelled: 'Cancelled',
  unrunnable: 'Not compared',
};

/** How each flatten choice reads to someone deciding whether to trust a redline. */
const FLATTEN_LABEL: Record<FlattenPolicy, string> = {
  acceptAll: 'Tracked changes accepted first',
  rejectAll: 'Tracked changes rejected first',
};

/**
 * Spec §10 — stated in-product.
 *
 * The first line is the load-bearing one. A grid of a hundred comparisons
 * invites the reading "so what do they have in common?", and this feature
 * deliberately does not answer that (the deviation matrix is cut — ADR-0037).
 * Leaving the invitation unanswered on screen is how a reader ends up inventing
 * an aggregate that was never computed.
 */
const BATCH_LIMITATIONS: readonly string[] = [
  'This is a list of comparisons, not an analysis across them: each row is one redline of one document against one other.',
  'Every pair is compared independently, so a clause moved between two versions is not tracked across them.',
  'Word-against-PDF pairs are identified but not compared.',
];

/**
 * The one §10 limitation that exists BECAUSE of many-to-many.
 *
 * It belongs on the surface that pairs, and nowhere else: on a one-to-many run
 * the user chose both sides of every comparison themselves, so there is no
 * pairing for this sentence to be about.
 */
const PAIRING_LIMITATION =
  'Pairing is lexical and structural: two documents that mean the same thing '
  + 'but share no wording will not pair. That is why the grid is confirmable.';

/**
 * The classification chips (spec §5.4).
 *
 * Ordered by how much a reader needs to see the class, not alphabetically:
 * a changed obligation or a changed number is what someone scanning a hundred
 * redlines is looking for, and renumbering is what they are looking past. Every
 * non-zero class is rendered — the ordering is emphasis, never a filter
 * (ADR-0008), and the chips always sum to the row's own change count.
 */
const CHIP: Record<ChangeClass, { rank: number; label: string }> = {
  obligationLanguage: { rank: 0, label: 'obligation' },
  numericalValue: { rank: 1, label: 'numbers' },
  definedTerm: { rank: 2, label: 'defined terms' },
  substantiveText: { rank: 3, label: 'text' },
  structural: { rank: 4, label: 'structure' },
  crossRefRenumber: { rank: 5, label: 'renumbering' },
  whitespacePunct: { rank: 6, label: 'whitespace' },
  formatting: { rank: 7, label: 'formatting' },
};

/**
 * Derived from the record above rather than written out as a second list.
 *
 * A `ChangeClass` added to the union fails typecheck on the record, so it
 * cannot be silently missing here — whereas a hand-written array would have
 * dropped it from the chips while `changeCount` went on counting it, quietly
 * breaking the invariant this file leans on.
 */
const CHIP_ORDER: readonly ChangeClass[] = (Object.keys(CHIP) as ChangeClass[])
  .sort((a, b) => (CHIP[a].rank - CHIP[b].rank));

export function BatchRowItem({
  row,
  onOpen,
  onDownload,
}: {
  row: BatchRow;
  onOpen?: (id: string) => void;
  onDownload: (id: string) => void;
}) {
  return (
    <li class={`dp-batch__row dp-batch__row--${row.status}`}>
      <span class="dp-batch__row-label" title={row.label}>{row.label}</span>
      <span class="dp-batch__row-status">{STATUS_LABEL[row.status]}</span>

      {row.status === 'done' && (
        <span class="dp-batch__row-count">
          {row.changeCount === 0
            ? 'No changes'
            : `${row.changeCount} change${row.changeCount === 1 ? '' : 's'}`}
        </span>
      )}

      {row.status === 'done' && row.changeCount > 0 && (
        <span class="dp-batch__row-chips">
          {CHIP_ORDER.filter(c => (row.byClassification[c] ?? 0) > 0).map(c => (
            <span key={c} class={`dp-batch__chip dp-batch__chip--${c}`}>
              {row.byClassification[c]} {CHIP[c].label}
            </span>
          ))}
        </span>
      )}

      {/* A reason is required for failed and unrunnable — showing the state
          without the reason is the part that would be useless. */}
      {row.reason !== undefined && (
        <span class="dp-batch__row-reason">{row.reason}</span>
      )}

      {row.artefactOnDemand && (
        <span class="dp-batch__row-reason">
          Redline not held in memory — this pair is compared again when you download it.
        </span>
      )}

      <span class="dp-batch__row-actions">
        {row.status === 'done' && onOpen !== undefined && (
          <button
            type="button"
            class="dp-batch__link dp-batch__open"
            onClick={() => onOpen(row.id)}
          >
            Open changes
          </button>
        )}
        {isDownloadable(row) && (
          <button type="button" class="dp-batch__link" onClick={() => onDownload(row.id)}>
            Download redline
          </button>
        )}
      </span>
    </li>
  );
}

/**
 * What a finished batch has to say about itself, under the rows.
 *
 * Three separate spec clauses, in one block because they are all statements
 * about the run rather than about any pair: the cancellation notice (§5.5), the
 * flatten choice every redline was produced under (§8 case 7), and the
 * limitations §10 requires in-product.
 */
export function BatchRunFootnotes({ index }: { index: BatchIndex }) {
  return (
    <div class="dp-batch__footnotes">
      {/* Spec §5.5: a cancelled batch must say so. The rows' own Cancelled
          chips are easy to miss under a bar that reads as complete. */}
      {index.totals.cancelled > 0 && (
        <p class="dp-batch__disclosure" role="status">
          Batch cancelled after {index.totals.done} of {index.totals.pairs}.
          The redlines already produced are kept.
        </p>
      )}

      {/* Spec §8 case 7 — stated, because it changes what every redline says.
          The two sides are named for the mode: "the template" and "each
          version" are the one-to-many shape, and calling a bundle's side A a
          template would be its own small untruth. */}
      <p class="dp-batch__note">
        {FLATTEN_LABEL[index.flatten.template]} on{' '}
        {index.mode === 'many-to-many' ? 'every side-A document' : 'the template'};{' '}
        {FLATTEN_LABEL[index.flatten.variant]} on{' '}
        {index.mode === 'many-to-many' ? 'every side-B document' : 'each version'}.
      </p>

      {/* Spec §10 — the limitations are stated in-product, not only in the
          spec. The first is the one a reader is most likely to assume away:
          this is a list of comparisons, not an analysis across them. */}
      <ul class="dp-batch__limits">
        {BATCH_LIMITATIONS.map(text => <li key={text}>{text}</li>)}
        {index.mode === 'many-to-many' && <li>{PAIRING_LIMITATION}</li>}
      </ul>
    </div>
  );
}
