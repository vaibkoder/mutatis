import { memo } from 'react';
import type {
  RcBlock,
  RcCell,
  RcParagraph,
  RcRun,
  RcTable,
  RenderModel,
} from '@shared/render-core/types.js';
import { paragraphStyle, runStyle, verticalTag } from './css-map.js';

/**
 * The flowed-prose redline.
 *
 * Built on the engine's `RenderModel`, which is pure XML→model: no fonts, no
 * HarfBuzz, no pagination. The browser does the line breaking, so the result is
 * proportional, reflowable prose with token-level marks INSIDE sentences —
 * which is what legal redlining needs and what a code-diff grid cannot give.
 *
 * NOTHING IS EVER HIDDEN HERE. Unchanged paragraphs render alongside changed
 * ones, because the engine's output is the full merged document sequence and
 * suppressing any of it would make a complete-looking redline that is missing
 * content.
 */

export function RedlineDocument({ model }: { model: RenderModel }) {
  return (
    <article className="redline mx-auto max-w-[52rem] px-10 py-8">
      {model.blocks.map((block, i) => (
        <Block key={i} block={block} />
      ))}
    </article>
  );
}

const Block = memo(function Block({ block }: { block: RcBlock }) {
  return block.kind === 'paragraph' ? (
    <Paragraph paragraph={block.paragraph} />
  ) : (
    <Table table={block.table} />
  );
});

function changedIn(runs: readonly RcRun[]): boolean {
  return runs.some((r) => r.revision !== undefined);
}

function Paragraph({ paragraph }: { paragraph: RcParagraph }) {
  const changed = changedIn(paragraph.runs);
  const Tag = paragraph.isHeading ? 'h2' : 'p';

  // An empty paragraph still occupies a line in the document, and dropping it
  // would silently reflow everything after it. The paragraph mark's own rPr
  // carries the height.
  const isEmpty = paragraph.runs.length === 0;

  return (
    <div className="relative">
      {/* Change bar: the margin rule that lets a reviewer scan for changes
          without reading. Same convention as the add-in. */}
      {changed ? (
        <span
          aria-hidden
          className="absolute top-0 bottom-0 -left-5 w-0.5 bg-[var(--redline-bar)]"
        />
      ) : null}

      <Tag
        style={paragraphStyle(paragraph.ppr)}
        className={paragraph.isHeading ? 'font-semibold' : undefined}
        data-changed={changed ? 'true' : undefined}
      >
        {paragraph.numberLabel ? (
          <span className="redline-number">{paragraph.numberLabel}&nbsp;</span>
        ) : null}
        {isEmpty ? (
          <span
            aria-hidden
            style={{ fontSize: `${paragraph.markRpr.sizePt}pt` }}
          >
            {' '}
          </span>
        ) : (
          paragraph.runs.map((run, i) => <Run key={i} run={run} />)
        )}
      </Tag>
    </div>
  );
}

function Run({ run }: { run: RcRun }) {
  if (run.imagePlaceholder) {
    return (
      <span className="redline-image" title="Image">
        [image]
      </span>
    );
  }

  const style = runStyle(run.rpr);
  const vertical = verticalTag(run.rpr);

  // Underline is applied here, not in runStyle, so it can be suppressed when a
  // revision decoration already owns the text-decoration slot.
  const revision = run.revision;
  const wantsOwnUnderline = run.rpr.underline && revision?.type !== 'ins';

  let content: React.ReactNode = run.text;
  if (vertical === 'sup') content = <sup>{content}</sup>;
  else if (vertical === 'sub') content = <sub>{content}</sub>;
  if (wantsOwnUnderline) content = <u>{content}</u>;

  if (!revision) {
    return (
      <span style={style} data-footnote={run.footnoteId}>
        {content}
      </span>
    );
  }

  const classes = run.changeClasses?.join(' ');
  const title = revisionTitle(run);

  switch (revision.type) {
    case 'ins':
      return (
        <ins style={style} className="redline-ins" title={title} data-classes={classes}>
          {content}
        </ins>
      );
    case 'del':
      return (
        <del style={style} className="redline-del" title={title} data-classes={classes}>
          {content}
        </del>
      );
    case 'moveFrom':
      return (
        <del style={style} className="redline-move-from" title={title} data-classes={classes}>
          {content}
        </del>
      );
    case 'moveTo':
      return (
        <ins style={style} className="redline-move-to" title={title} data-classes={classes}>
          {content}
        </ins>
      );
  }
}

function revisionTitle(run: RcRun): string {
  const r = run.revision;
  if (!r) return '';
  const kind =
    r.type === 'ins'
      ? 'Inserted'
      : r.type === 'del'
        ? 'Deleted'
        : r.type === 'moveFrom'
          ? 'Moved from here'
          : 'Moved to here';
  const who = r.author ? ` by ${r.author}` : '';
  const classes = run.changeClasses?.length ? ` — ${run.changeClasses.join(', ')}` : '';
  return `${kind}${who}${classes}`;
}

function Table({ table }: { table: RcTable }) {
  return (
    <div className="my-3 overflow-x-auto">
      <table className="redline-table w-full border-collapse">
        <tbody>
          {table.rows.map((row, r) => (
            <tr key={r}>
              {row.cells.map((cell, c) => (
                <Cell key={c} cell={cell} />
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function Cell({ cell }: { cell: RcCell }) {
  return (
    <td
      colSpan={cell.gridSpan > 1 ? cell.gridSpan : undefined}
      className="border border-border px-2 py-1 align-top"
    >
      {cell.blocks.map((block, i) => (
        <Block key={i} block={block} />
      ))}
    </td>
  );
}
