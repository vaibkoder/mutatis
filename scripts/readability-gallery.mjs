/**
 * Build the redline specimen gallery from the run snapshots in .redline-poc/runs/.
 *
 *   node scripts/readability-gallery.mjs [outfile]
 *
 * Every snapshot written by scripts/readability-snapshot.ts becomes a column of
 * specimens, so a reader can see how each case's rendering moved run by run and
 * judge it against what they asked for. Development tool only.
 */
import { readdirSync, readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { join, dirname } from 'node:path';

const DIR = '.redline-poc/runs';
const OUT = process.argv[2] ?? '.redline-poc/redline-gallery.html';

const runs = readdirSync(DIR)
  .filter((f) => f.endsWith('.json'))
  .sort()
  .map((f) => ({ file: f, ...JSON.parse(readFileSync(join(DIR, f), 'utf8')) }));

if (runs.length === 0) throw new Error(`no snapshots in ${DIR}`);

const first = runs[0];
const last = runs[runs.length - 1];

const esc = (s) =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** Turn the bracket notation into the marks a lawyer reads: strike-through, underline, move. */
function render(line) {
  const out = [];
  const re = /(\{\+|\[-|\{>|\[<)([\s\S]*?)(\+\}|-\]|>\}|<\])/g;
  let at = 0;
  let m;
  while ((m = re.exec(line)) !== null) {
    if (m.index > at) out.push(esc(line.slice(at, m.index)));
    const body = esc(m[2]);
    if (m[1] === '{+') out.push(`<ins>${body}</ins>`);
    else if (m[1] === '[-') out.push(`<del>${body}</del>`);
    else if (m[1] === '{>') out.push(`<span class="mv mv-to">${body}</span>`);
    else out.push(`<span class="mv mv-from">${body}</span>`);
    at = re.lastIndex;
  }
  if (at < line.length) out.push(esc(line.slice(at)));
  return out.join('');
}

const byId = (run, id) => run.cases.find((c) => c.id === id);
const sameLines = (a, b) =>
  a && b && a.lines.length === b.lines.length && a.lines.every((l, i) => l === b.lines[i]);

const GRADE_LABEL = { good: 'reads well', partial: 'reads partly', bad: 'reads badly', unreviewed: 'awaiting verdict' };

const caseCards = last.cases
  .map((c0) => {
    const base = byId(first, c0.id) ?? c0;
    const cur = c0;
    const moved = runs.length > 1 && !sameLines(base, cur);
    const solved = moved && cur.metrics.redundancy <= 0.11 && base.metrics.redundancy > 0.11;

    const specimens = runs
      .map((run, i) => {
        const c = byId(run, base.id);
        if (!c) return '';
        const prev = i > 0 ? byId(runs[i - 1], base.id) : null;
        const unchanged = prev && sameLines(prev, c);
        if (unchanged && i > 0 && i < runs.length - 1) return '';
        const d = c.metrics;
        const delta =
          prev && prev.metrics.redundancy !== d.redundancy
            ? `<span class="delta ${d.redundancy < prev.metrics.redundancy ? 'better' : 'worse'}">${prev.metrics.redundancy} → ${d.redundancy}</span>`
            : '';
        return `<div class="specimen${i === runs.length - 1 ? ' latest' : ''}">
        <div class="runlabel">${esc(run.label)}${unchanged ? ' <em>(unchanged)</em>' : ''}</div>
        <div class="doc">${c.lines.map((l) => `<p>${render(l)}</p>`).join('')}</div>
        <div class="metrics"><span>marks <b>${d.marks}</b></span><span>redundancy <b>${d.redundancy}</b></span>${delta}</div>
      </div>`;
      })
      .join('');

    const wanted = cur.wanted
      ? `<div class="specimen wanted">
        <div class="runlabel">what the reviewer asked for</div>
        <div class="doc"><p>${render(cur.wanted)}</p></div>
      </div>`
      : '';

    return `<article class="case" data-grade="${base.grade}" data-moved="${moved}" data-solved="${solved}">
    <header>
      <span class="cid">${base.id}</span>
      <h2>${esc(base.title)}</h2>
      <span class="chip g-${base.grade}">${GRADE_LABEL[base.grade]}</span>
      ${moved ? `<span class="chip changed">${solved ? 'resolved' : 'changed'}</span>` : ''}
    </header>
    <p class="act">${esc(base.act)}${base.cites.length ? ` · <span class="cites">${base.cites.join(' ')}</span>` : ''}</p>
    ${specimens}${wanted}
  </article>`;
  })
  .join('\n');

const runSummary = runs
  .map(
    (r, i) => `<div class="runstat${i === runs.length - 1 ? ' latest' : ''}">
    <div class="rl">${esc(r.label)}</div>
    <div class="rv">${r.totals.redundancy}</div>
    <div class="rk">total redundancy</div>
  </div>`,
  )
  .join('<div class="arrow">→</div>');

const changedCount = last.cases.filter((c) => !sameLines(byId(first, c.id), c)).length;

const html = `<title>Redline Specimens</title>
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@400;500;600&family=Source+Serif+4:opsz,wght@8..60,400;8..60,600&display=swap">
<style>
  :root {
    --paper: #f7f8fa;
    --card: #ffffff;
    --ink: #14181f;
    --ink-2: #4a5563;
    --ink-3: #7b8794;
    --rule: #dde2e9;
    --rule-2: #eaeef3;
    --del: #a3202b;
    --del-bg: #a3202b14;
    --ins: #1c4f8a;
    --ins-bg: #1c4f8a12;
    --mv: #0f6b52;
    --accent: #2f4a6d;
    --good: #4a7a5e;
    --partial: #97701f;
    --bad: #a3202b;
    --serif: 'Source Serif 4', Georgia, 'Times New Roman', serif;
    --sans: 'IBM Plex Sans', ui-sans-serif, system-ui, sans-serif;
    --mono: 'IBM Plex Mono', ui-monospace, 'SF Mono', Menlo, monospace;
  }
  @media (prefers-color-scheme: dark) {
    :root:not([data-theme="light"]) {
      --paper: #0e1116; --card: #161b23; --ink: #e6eaf0; --ink-2: #a6b1bf; --ink-3: #76828f;
      --rule: #262e39; --rule-2: #1e252e;
      --del: #f1808c; --del-bg: #f1808c1f; --ins: #7fb2ec; --ins-bg: #7fb2ec1a;
      --mv: #5fc3a4; --accent: #9dbbdd;
      --good: #7fb894; --partial: #d4ac52; --bad: #f1808c;
    }
  }
  :root[data-theme="dark"] {
    --paper: #0e1116; --card: #161b23; --ink: #e6eaf0; --ink-2: #a6b1bf; --ink-3: #76828f;
    --rule: #262e39; --rule-2: #1e252e;
    --del: #f1808c; --del-bg: #f1808c1f; --ins: #7fb2ec; --ins-bg: #7fb2ec1a;
    --mv: #5fc3a4; --accent: #9dbbdd;
    --good: #7fb894; --partial: #d4ac52; --bad: #f1808c;
  }

  body { background: var(--paper); color: var(--ink); font-family: var(--sans); line-height: 1.5; }
  .wrap { max-width: 62rem; margin: 0 auto; padding: 2.5rem 1.25rem 5rem; }

  .masthead { border-bottom: 2px solid var(--ink); padding-bottom: 1.1rem; margin-bottom: 1.6rem; }
  .kicker { font-size: .68rem; letter-spacing: .16em; text-transform: uppercase; color: var(--ink-3); font-weight: 600; }
  h1 { font-family: var(--serif); font-size: clamp(1.7rem, 4vw, 2.5rem); font-weight: 600; margin: .3rem 0 .5rem; letter-spacing: -.01em; text-wrap: balance; }
  .standfirst { color: var(--ink-2); max-width: 46rem; font-size: .95rem; }

  .board { display: flex; align-items: center; gap: 1rem; flex-wrap: wrap; background: var(--card); border: 1px solid var(--rule); border-radius: 4px; padding: 1rem 1.15rem; margin: 1.4rem 0 1.1rem; }
  .runstat .rl { font-size: .7rem; text-transform: uppercase; letter-spacing: .09em; color: var(--ink-3); font-weight: 600; }
  .runstat .rv { font-family: var(--mono); font-size: 1.55rem; font-weight: 600; font-variant-numeric: tabular-nums; line-height: 1.15; }
  .runstat .rk { font-size: .7rem; color: var(--ink-3); }
  .runstat.latest .rv { color: var(--accent); }
  .arrow { color: var(--ink-3); font-size: 1.1rem; }
  .note { margin-left: auto; font-size: .8rem; color: var(--ink-2); text-align: right; }

  .filters { display: flex; gap: .4rem; flex-wrap: wrap; margin-bottom: 1.6rem; }
  .filters button { font: inherit; font-size: .8rem; padding: .3rem .75rem; border: 1px solid var(--rule); background: var(--card); color: var(--ink-2); border-radius: 999px; cursor: pointer; }
  .filters button[aria-pressed="true"] { background: var(--ink); color: var(--paper); border-color: var(--ink); }
  .filters button:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }

  .case { border-top: 1px solid var(--rule); padding: 1.5rem 0 1.1rem; }
  .case header { display: flex; align-items: baseline; gap: .6rem; flex-wrap: wrap; }
  .cid { font-family: var(--mono); font-size: .78rem; font-weight: 600; color: var(--ink-3); }
  .case h2 { font-family: var(--serif); font-size: 1.12rem; font-weight: 600; margin: 0; letter-spacing: -.005em; }
  .chip { font-size: .67rem; letter-spacing: .06em; text-transform: uppercase; font-weight: 600; padding: .12rem .5rem; border-radius: 3px; border: 1px solid currentColor; }
  .g-good { color: var(--good); } .g-partial { color: var(--partial); } .g-bad { color: var(--bad); } .g-unreviewed { color: var(--accent); }
  .chip.changed { color: var(--accent); }
  .act { font-size: .8rem; color: var(--ink-3); margin: .25rem 0 .9rem; }
  .cites { font-family: var(--mono); font-size: .74rem; }

  .specimen { margin: 0 0 .7rem; }
  .runlabel { font-size: .68rem; text-transform: uppercase; letter-spacing: .09em; color: var(--ink-3); font-weight: 600; margin-bottom: .3rem; }
  .runlabel em { text-transform: none; letter-spacing: 0; font-weight: 400; }
  .doc { background: var(--card); border: 1px solid var(--rule); border-left: 3px solid var(--rule); border-radius: 3px; padding: .8rem 1rem; font-family: var(--serif); font-size: .97rem; line-height: 1.62; overflow-x: auto; }
  .specimen.latest .doc { border-left-color: var(--accent); }
  .specimen.wanted .doc { border-left-color: var(--good); border-style: dashed; background: transparent; }
  .doc p { margin: 0 0 .35rem; }
  .doc p:last-child { margin-bottom: 0; }
  del { color: var(--del); background: var(--del-bg); text-decoration: line-through; text-decoration-thickness: 1px; }
  ins { color: var(--ins); background: var(--ins-bg); text-decoration: underline; text-decoration-thickness: 1px; text-underline-offset: 2px; }
  .mv { color: var(--mv); background: color-mix(in srgb, var(--mv) 10%, transparent); }
  .mv-from { text-decoration: line-through; } .mv-to { text-decoration: underline; text-underline-offset: 2px; }

  .metrics { display: flex; gap: 1.1rem; font-family: var(--mono); font-size: .72rem; color: var(--ink-3); margin-top: .35rem; font-variant-numeric: tabular-nums; }
  .metrics b { color: var(--ink-2); font-weight: 600; }
  .delta.better { color: var(--good); } .delta.worse { color: var(--bad); }

  .legend { margin-top: 2.5rem; border-top: 1px solid var(--rule); padding-top: 1rem; font-size: .8rem; color: var(--ink-2); }
  .legend code { font-family: var(--mono); font-size: .75rem; }
</style>
<div class="wrap">
  <header class="masthead">
    <div class="kicker">Mutatis · Compare</div>
    <h1>Indemnity redline specimens</h1>
    <p class="standfirst">Twenty-two contract amendments, each isolating one editing act, rendered through the real Compare pipeline and the real RenderCore parser — so what you read below is the run structure the paginated preview and the vector PDF paint. Judge whether each one reads the way a conventional redline would.</p>
  </header>

  <div class="board">
    ${runSummary}
    <div class="note">${changedCount} of ${last.cases.length} case${changedCount === 1 ? '' : 's'} changed since the baseline.<br><span style="color:var(--ink-3)">Redundancy = the share of a struck span that is word-for-word identical to the insertion beside it. Floor is 0.11.</span></div>
  </div>

  <div class="filters" role="group" aria-label="Filter specimens">
    <button data-f="all" aria-pressed="true">All ${last.cases.length}</button>
    <button data-f="unreviewed" aria-pressed="false">Awaiting your verdict</button>
    <button data-f="bad" aria-pressed="false">Reads badly</button>
    <button data-f="partial" aria-pressed="false">Reads partly</button>
    <button data-f="moved" aria-pressed="false">Changed this session</button>
  </div>

${caseCards}

  <div class="legend">
    <p><del>struck</del> = deleted · <ins>inserted</ins> = inserted · <span class="mv mv-from">moved</span> from / <span class="mv mv-to">moved</span> to. Clause numbers are shown exactly as the redline emits them.</p>
    <p>Principles P1–P10 are in <code>docs/compare-readability-principles.md</code>.</p>
  </div>
</div>
<script>
  const buttons = [...document.querySelectorAll('.filters button')];
  const cases = [...document.querySelectorAll('.case')];
  function apply(f) {
    for (const c of cases) {
      const show =
        f === 'all' ? true :
        f === 'moved' ? c.dataset.moved === 'true' :
        c.dataset.grade === f;
      c.hidden = !show;
    }
    for (const b of buttons) b.setAttribute('aria-pressed', String(b.dataset.f === f));
    try { localStorage.setItem('dp-redline-filter', f); } catch {}
  }
  for (const b of buttons) b.addEventListener('click', () => apply(b.dataset.f));
  let saved = 'all';
  try { saved = localStorage.getItem('dp-redline-filter') || 'all'; } catch {}
  if (buttons.some((b) => b.dataset.f === saved)) apply(saved);
</script>
`;

mkdirSync(dirname(OUT), { recursive: true });
writeFileSync(OUT, html);
console.log(`${OUT}  (${runs.length} run${runs.length === 1 ? '' : 's'}, ${last.cases.length} cases)`);
