import type { CSSProperties } from 'react';
import type { Block, Para, Run } from './model.js';

/**
 * Renders blocks as React elements. This is how pasted formatting is SHOWN: the
 * clipboard HTML itself is never put in the page.
 */

const HEADING_CLASS = [
  'text-[17px] font-bold',
  'text-[15px] font-bold',
  'text-[14px] font-bold',
  'text-[13px] font-bold',
  'text-[13px] font-semibold',
  'text-[13px] font-semibold',
];

function RunView({ run }: { run: Run }) {
  if (run.br) return <br />;
  if (run.tab) return <span className="inline-block w-8" />;
  const deco = [run.underline ? 'underline' : '', run.strike ? 'line-through' : ''].filter(Boolean).join(' ');
  const className = [run.bold ? 'font-bold' : '', run.italic ? 'italic' : '', deco].filter(Boolean).join(' ');
  const inner = run.text;
  if (run.sup) return <sup className={className}>{inner}</sup>;
  if (run.sub) return <sub className={className}>{inner}</sub>;
  return className ? <span className={className}>{inner}</span> : <>{inner}</>;
}

function ParaView({ p }: { p: Para }) {
  const style: CSSProperties = {
    ...(p.indent ? { paddingLeft: `${p.indent / 20}pt` } : {}),
    ...(p.align ? { textAlign: p.align === 'both' ? 'justify' : p.align } : {}),
  };
  const className =
    'min-h-[1.4em] whitespace-pre-wrap break-words ' +
    (p.heading !== undefined ? `${HEADING_CLASS[p.heading - 1]} mt-2` : '');
  return (
    <p className={className} style={style}>
      {p.runs.map((r, i) => (
        <RunView key={i} run={r} />
      ))}
    </p>
  );
}

export function BlocksPreview({ blocks, className = '' }: { blocks: readonly Block[]; className?: string }) {
  return (
    <div className={`flex flex-col gap-1 text-[13px] leading-snug text-text ${className}`}>
      {blocks.map((b, i) =>
        b.kind === 'p' ? (
          <ParaView key={i} p={b} />
        ) : (
          <table key={i} className="my-1 w-full border-collapse text-[12px]">
            <tbody>
              {b.rows.map((row, ri) => (
                <tr key={ri}>
                  {row.map((cell, ci) => (
                    <td key={ci} className="border border-border-strong px-1.5 py-1 align-top">
                      {cell.map((p, pi) => (
                        <ParaView key={pi} p={p} />
                      ))}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        ),
      )}
    </div>
  );
}
