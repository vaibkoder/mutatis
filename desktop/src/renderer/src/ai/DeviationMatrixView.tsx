import { Info, Table2 } from 'lucide-react';
import type { DeviationMatrix } from './deviation-matrix.js';

/**
 * The matrix, rendered from the computed data.
 *
 * Nothing on this screen came from a language model. The narration of it, if the
 * user asked for one, is in the chat column beside this tab — which is the whole
 * point of putting the matrix here rather than inside a chat bubble.
 */
export function DeviationMatrixView({ matrix }: { matrix: DeviationMatrix }) {
  return (
    <div className="base-ui-disable-scrollbar h-full overflow-auto px-5 py-4">
      <header className="mb-3 flex items-center gap-2">
        <Table2 size={15} className="text-accent" aria-hidden />
        <h2 className="text-[13px] font-semibold text-text">Deviation matrix</h2>
        <span className="text-[11px] text-text-faint">
          {matrix.columns.length} comparison{matrix.columns.length === 1 ? '' : 's'}
        </span>
      </header>

      <div className="overflow-x-auto rounded border border-border">
        <table className="w-full border-collapse text-[11px]">
          <thead>
            <tr className="bg-surface-sunken">
              <th className="sticky left-0 z-10 bg-surface-sunken px-2.5 py-1.5 text-left font-semibold text-text-muted">
                Classification
              </th>
              {matrix.columns.map((column) => (
                <th
                  key={column.title}
                  title={`${column.baseline} → ${column.revised}`}
                  className="px-2.5 py-1.5 text-right font-semibold text-text-muted"
                >
                  {column.title}
                </th>
              ))}
              <th className="px-2.5 py-1.5 text-right font-semibold text-text">Total</th>
              <th className="px-2.5 py-1.5 text-right font-semibold text-text-muted">In</th>
            </tr>
          </thead>
          <tbody>
            {matrix.rows.length === 0 ? (
              <tr>
                <td
                  colSpan={matrix.columns.length + 3}
                  className="px-2.5 py-4 text-center text-text-muted"
                >
                  No changes were found in any of these comparisons.
                </td>
              </tr>
            ) : (
              matrix.rows.map((row) => (
                <tr key={row.classification} className="border-t border-border">
                  <td className="sticky left-0 z-10 bg-surface px-2.5 py-1.5 text-text">
                    {row.classification}
                  </td>
                  {row.cells.map((cell, index) => (
                    <td
                      key={matrix.columns[index]?.title ?? index}
                      className={
                        cell.count === 0
                          ? 'px-2.5 py-1.5 text-right text-text-faint'
                          : 'px-2.5 py-1.5 text-right text-text'
                      }
                    >
                      {cell.count === 0 ? '—' : cell.count}
                    </td>
                  ))}
                  <td className="px-2.5 py-1.5 text-right font-medium text-text">{row.total}</td>
                  <td className="px-2.5 py-1.5 text-right text-text-muted">
                    {row.documents}/{matrix.columns.length}
                  </td>
                </tr>
              ))
            )}
          </tbody>
          <tfoot>
            <tr className="border-t border-border-strong bg-surface-sunken">
              <td className="sticky left-0 z-10 bg-surface-sunken px-2.5 py-1.5 font-semibold text-text">
                Changes
              </td>
              {matrix.columns.map((column) => (
                <td key={column.title} className="px-2.5 py-1.5 text-right font-medium text-text">
                  {column.changes}
                </td>
              ))}
              <td className="px-2.5 py-1.5 text-right font-semibold text-text">
                {matrix.columns.reduce((sum, column) => sum + column.changes, 0)}
              </td>
              <td />
            </tr>
          </tfoot>
        </table>
      </div>

      <div className="mt-3 grid gap-2 sm:grid-cols-2">
        <Panel title="In every comparison" items={matrix.universal} />
        <Panel title="In only one comparison" items={matrix.unique} />
      </div>

      <ul className="mt-3 flex flex-col gap-1.5">
        {matrix.notes.map((note) => (
          <li key={note} className="flex items-start gap-1.5 text-[11px] leading-relaxed text-text-muted">
            <Info size={11} className="mt-0.5 shrink-0 text-text-faint" aria-hidden />
            {note}
          </li>
        ))}
      </ul>
    </div>
  );
}

function Panel({ title, items }: { title: string; items: readonly string[] }) {
  return (
    <div className="rounded border border-border p-2.5">
      <h3 className="mb-1 text-[11px] font-semibold tracking-wide uppercase text-text-muted">
        {title}
      </h3>
      {items.length === 0 ? (
        <p className="text-[11px] text-text-faint">None.</p>
      ) : (
        <p className="text-[11px] leading-relaxed text-text">{items.join(', ')}</p>
      )}
    </div>
  );
}
