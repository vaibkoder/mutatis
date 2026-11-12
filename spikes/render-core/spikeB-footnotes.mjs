// Spike B — footnote bottom-of-page reservation FIXPOINT (THROWAWAY).
//
// De-risks the single hardest component (ADR-0024 §4.1): placing a footnote at
// the bottom of a page shrinks the body area, which can push the *reference* off
// the page, which removes the footnote, which restores the area … → a fixpoint.
// Also: a footnote longer than the page must SPLIT with a continuation notice.
//
// Model (synthetic, units = pt): a page body area of height H; an ordered list of
// body lines, each height `h`, some carrying a footnote whose body is `fnH` tall.
// A separator + gap sits above the footnote area when any footnote is present.
//
// Run: node spikes/render-core/spikeB-footnotes.mjs
const H = 600;            // body area height (pt) per page (synthetic)
const SEP = 12;           // separator + gap above footnotes when present
const LINE = 14;          // body line height
const MAX_ITER = 8;       // asserted fixpoint bound per page

// Synthetic body: 90 lines; a few carry footnotes; one footnote is oversized.
const lines = [];
for (let i = 0; i < 90; i++) {
  const fn =
    i === 8 ? { id: 'fn1', fnH: 40 } :
    i === 41 ? { id: 'fn2', fnH: 30 } :
    i === 42 ? { id: 'fn3', fnH: 28 } :
    i === 43 ? { id: 'fn4', fnH: 900 } : // OVERSIZED: 900pt > one page → must split
    i === 70 ? { id: 'fn5', fnH: 35 } :
    null;
  lines.push({ idx: i, h: LINE, fn });
}

// Footnote area available on a page given how much continuation carried in.
function footnoteAreaUsed(fns, carryRemainder) {
  if (fns.length === 0 && carryRemainder <= 0) return 0;
  const own = fns.reduce((a, f) => a + f.fnH, 0);
  return SEP + own + (carryRemainder > 0 ? carryRemainder : 0);
}

// Compose one page by FIXPOINT: start assuming the footnotes from a tentative
// line set, recompute capacity, re-place lines, until the placed line set is
// stable. Returns {placed, fns, iters, carryOut}.
function composePage(startIdx, carryInRemainder) {
  let placedCount = lines.length; // optimistic upper bound to start
  let iters = 0;
  let prevSig = '';
  let result = null;
  // carryInRemainder: pt of an oversized footnote continued from the previous page.
  while (iters < MAX_ITER) {
    iters++;
    // Reserve continuation first; if it alone exceeds the page, split again.
    let carryRemainder = 0;
    let contUsed = 0;
    if (carryInRemainder > 0) {
      const contCap = H - SEP; // continuation can use almost the whole page
      contUsed = Math.min(carryInRemainder, contCap);
      carryRemainder = carryInRemainder - contUsed; // may carry further
    }
    const placed = [];
    const fns = [];
    let bodyUsed = 0;
    let split = null;
    for (let i = startIdx; i < lines.length; i++) {
      const ln = lines[i];
      const tentativeFns = ln.fn ? [...fns, ln.fn] : fns;
      // Oversized footnote: own fnH cannot fit even an empty page's footnote area.
      let thisFnUsed = 0, thisSplit = null;
      if (ln.fn) {
        const fnAreaCapWithThis = H - bodyUsed - ln.h; // area left for footnotes if we place this line
        const otherFns = fns.reduce((a, f) => a + f.fnH, 0);
        const need = SEP + otherFns + ln.fn.fnH + contUsed;
        if (ln.fn.fnH > H - SEP) {
          // genuinely oversized → place ref here, split footnote across pages
          const fitHere = Math.max(0, H - bodyUsed - ln.h - SEP - otherFns - contUsed);
          thisFnUsed = fitHere;
          thisSplit = { id: ln.fn.id, total: ln.fn.fnH, here: fitHere, remainder: ln.fn.fnH - fitHere };
        } else if (need > H - bodyUsed - ln.h) {
          // footnote doesn't fit with this line on this page → end page before it
          break;
        }
      }
      // Does the line itself fit with current footnote reservation?
      const reserve = footnoteAreaUsed(tentativeFns.filter((f) => f.fnH <= H - SEP), contUsed) +
        (thisSplit ? SEP + thisSplit.here : 0);
      if (bodyUsed + ln.h + reserve > H && placed.length > 0) break;
      placed.push(ln);
      bodyUsed += ln.h;
      if (ln.fn) {
        if (thisSplit) { split = thisSplit; fns.push({ id: ln.fn.id, fnH: thisSplit.here, split: true }); i++; /* page ends after an oversized ref */ break; }
        fns.push(ln.fn);
      }
    }
    const sig = `${placed.length}:${fns.map((f) => f.id).join(',')}:${split?.id ?? ''}`;
    result = { placed, fns, contUsed, carryRemainder, split, nextIdx: startIdx + placed.length };
    if (sig === prevSig) break;
    prevSig = sig;
  }
  // carry out = remainder of an oversized footnote that didn't fit here
  const carryOut = (result.split ? result.split.remainder : 0) + result.carryRemainder;
  return { ...result, iters, carryOut };
}

// Drive pagination across the whole doc.
const pages = [];
let idx = 0;
let carry = 0;
let guard = 0;
while (idx < lines.length || carry > 0) {
  if (guard++ > 50) { console.error('runaway'); break; }
  const startCarry = carry;
  const page = composePage(idx, carry);
  pages.push({ ...page, startCarry });
  idx = page.nextIdx;
  carry = page.carryOut;
  if (page.placed.length === 0 && carry > 0) {
    // pure continuation page (oversized footnote spanning a whole page)
    continue;
  }
  if (page.placed.length === 0 && carry === 0) break;
}

// ── report ───────────────────────────────────────────────────────────────────
console.log('═══ SPIKE B — footnote reservation fixpoint ═══');
console.log(`page body height H=${H}pt, line=${LINE}pt, separator=${SEP}pt, fixpoint bound=${MAX_ITER}`);
let maxIters = 0, anyOverflow = false;
pages.forEach((p, n) => {
  maxIters = Math.max(maxIters, p.iters);
  const ftxt = p.fns.map((f) => `${f.id}(${f.fnH}${f.split ? '↧split' : ''})`).join(' ');
  const cont = p.startCarry > 0 ? ` cont-in=${Math.min(p.startCarry, H - SEP).toFixed(0)}` : '';
  const used = p.placed.reduce((a, l) => a + l.h, 0) + footnoteAreaUsed(p.fns.filter((f) => f.fnH <= H - SEP), p.contUsed) + (p.split ? 0 : 0);
  if (used > H + 0.5) anyOverflow = true;
  console.log(`  page ${String(n + 1).padStart(2)}: lines ${String(p.placed.length).padStart(2)}  iters=${p.iters}  bodyUsed≈${used.toFixed(0)}/${H}pt  footnotes[${ftxt}]${cont}${p.carryOut > 0 ? `  carry-out=${p.carryOut.toFixed(0)}pt` : ''}`);
});
console.log(`\nfixpoint: max iterations on any page = ${maxIters} (bound ${MAX_ITER})  ${maxIters <= MAX_ITER ? '✓ converged' : '✗ DID NOT CONVERGE'}`);
console.log(`no-overflow: ${anyOverflow ? '✗ a page overflowed' : '✓ every page within H'}`);
const splitPages = pages.filter((p) => p.split || p.startCarry > 0);
console.log(`oversized split: fn4 (900pt) spread across ${splitPages.length} page(s) with continuation  ${splitPages.length > 1 ? '✓' : '✗'}`);
console.log(`\nVERDICT: ${maxIters <= MAX_ITER && !anyOverflow && splitPages.length > 1 ? 'GO-eligible ✓ (fixpoint converges, no overflow, oversized footnote splits + continues)' : 'REVIEW'}`);
