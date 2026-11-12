/**
 * Feature 08 — Compare: pure OOXML revision-markup authoring helpers (C1).
 *
 * Every function here is PURE: it takes plain string XML and returns plain
 * string XML. No Office JS, no Date.now(), no Math.random(). `id`, `author`,
 * and `date` are always injected by the caller so the output is deterministic
 * (a hard constraint for testability — spec §5.5, key constraint 1).
 *
 * Spec: spec/features/08-compare.md §5.5 (output OOXML construction).
 */

/**
 * Escape a value destined for an XML attribute.
 * Author strings come from Config and may contain `&`, `<`, `"` etc.
 */
function escapeAttr(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Common attribute string shared by every revision element. */
function revAttrs(id: number, author: string, date: string): string {
  return `w:id="${id}" w:author="${escapeAttr(author)}" w:date="${escapeAttr(date)}"`;
}

/**
 * Wrap run XML in a `<w:ins>` tracked-change element.
 * The `<w:ins>` element wraps `<w:r>` runs, never `<w:p>` paragraphs
 * (key constraint 7).
 */
export function wrapIns(runXml: string, id: number, author: string, date: string): string {
  return `<w:ins ${revAttrs(id, author, date)}>${runXml}</w:ins>`;
}

/**
 * Wrap run XML in a `<w:del>` tracked-change element.
 * Per OOXML, deleted run text must use `<w:delText>` instead of `<w:t>`
 * (otherwise Word renders the text as still-present). We rewrite every
 * `<w:t …>…</w:t>` to `<w:delText …>…</w:delText>` inside the wrapped runs.
 */
export function wrapDel(runXml: string, id: number, author: string, date: string): string {
  return `<w:del ${revAttrs(id, author, date)}>${toDelText(runXml)}</w:del>`;
}

/**
 * Convert every `<w:t>` element in a run-XML fragment to `<w:delText>`, and
 * every `<w:instrText>` to `<w:delInstrText>`, preserving attributes (notably
 * `xml:space="preserve"`). Used by `wrapDel` and by the author when emitting
 * deletions. A deleted field's instruction must be `<w:delInstrText>`
 * (ECMA-376 §17.16.13). Until ADR-0045 this converted `<w:t>` only, so a
 * whole-paragraph deletion carrying a field struck its result and left its
 * instruction as live `<w:instrText>` inside the `<w:del>`.
 */
export function toDelText(runXml: string): string {
  return runXml
    .replace(/<w:instrText(\s[^>]*)?\/>/g, (_m, attrs: string | undefined) => `<w:delInstrText${attrs ?? ''}/>`)
    .replace(/<w:instrText(\s[^>]*)?>/g, (_m, attrs: string | undefined) => `<w:delInstrText${attrs ?? ''}>`)
    .replace(/<\/w:instrText>/g, '</w:delInstrText>')
    // self-closing empty text elements first: <w:t/> → <w:delText/>
    .replace(/<w:t(\s[^>]*)?\/>/g, (_m, attrs: string | undefined) => `<w:delText${attrs ?? ''}/>`)
    .replace(/<w:t(\s[^>]*)?>/g, (_m, attrs: string | undefined) => `<w:delText${attrs ?? ''}>`)
    .replace(/<\/w:t>/g, '</w:delText>');
}

/**
 * Emit a moveFrom range around a paragraph (the source location of a move):
 *   <w:moveFromRangeStart w:id w:name w:author w:date/>
 *   …paragraph with its runs wrapped in <w:moveFrom>…
 *   <w:moveFromRangeEnd w:id/>
 *
 * `name` is the stable identifier shared with the matching moveTo range.
 * The runs inside the paragraph are wrapped in `<w:moveFrom>` (the run-level
 * counterpart, analogous to `<w:del>`), so Word renders the relocation as a
 * move rather than a delete.
 */
export function wrapMoveFrom(
  paraXml: string,
  rangeId: number,
  name: string,
  author: string,
  date: string,
): string {
  const startId = rangeId;
  const endId = rangeId + 1;
  const moveId = rangeId + 2;
  const start = `<w:moveFromRangeStart w:id="${startId}" w:name="${escapeAttr(name)}" w:author="${escapeAttr(author)}" w:date="${escapeAttr(date)}"/>`;
  const end = `<w:moveFromRangeEnd w:id="${endId}"/>`;
  const inner = wrapMoveFromRuns(paraXml, moveId, author, date);
  return `${start}${inner}${end}`;
}

/**
 * Emit a moveTo range around a paragraph (the destination of a move):
 *   <w:moveToRangeStart w:id w:name w:author w:date/>
 *   …paragraph with its runs wrapped in <w:moveTo>…
 *   <w:moveToRangeEnd w:id/>
 */
export function wrapMoveTo(
  paraXml: string,
  rangeId: number,
  name: string,
  author: string,
  date: string,
): string {
  const startId = rangeId;
  const endId = rangeId + 1;
  const moveId = rangeId + 2;
  const start = `<w:moveToRangeStart w:id="${startId}" w:name="${escapeAttr(name)}" w:author="${escapeAttr(author)}" w:date="${escapeAttr(date)}"/>`;
  const end = `<w:moveToRangeEnd w:id="${endId}"/>`;
  const inner = wrapMoveToRuns(paraXml, moveId, author, date);
  return `${start}${inner}${end}`;
}

/**
 * Wrap the run content of a paragraph in `<w:moveFrom>`. The paragraph element
 * and its `<w:pPr>` are kept outside the revision element; only the runs move.
 */
function wrapMoveFromRuns(paraXml: string, id: number, author: string, date: string): string {
  return wrapParaRuns(paraXml, runs => `<w:moveFrom ${revAttrs(id, author, date)}>${runs}</w:moveFrom>`);
}

function wrapMoveToRuns(paraXml: string, id: number, author: string, date: string): string {
  return wrapParaRuns(paraXml, runs => `<w:moveTo ${revAttrs(id, author, date)}>${runs}</w:moveTo>`);
}

/**
 * The LEADING `<w:pPr>…</w:pPr>` (or `<w:pPr/>`) of a paragraph body, or `null`.
 *
 * `w:pPr` **nests**: `w:pPrChange` carries a complete `<w:pPr>` of its own — the
 * properties as they stood before a tracked property change. A lazy
 * `[\s\S]*?</w:pPr>` therefore stops at the **inner** close and returns a
 * truncated, unbalanced fragment. Depth-count instead, so callers always get a
 * whole element.
 */
export function matchLeadingPPr(rest: string): string | null {
  const self = /^\s*<w:pPr(?:\s[^>]*)?\/>/.exec(rest);
  if (self) return self[0];
  if (!/^\s*<w:pPr(?:\s[^>]*)?>/.test(rest)) return null;

  const tag = /<w:pPr(?:\s[^>]*?)?(\/)?>|<\/w:pPr>/g;
  let depth = 0;
  let m: RegExpExecArray | null;
  while ((m = tag.exec(rest)) !== null) {
    if (m[0].startsWith('</')) {
      depth--;
      if (depth === 0) return rest.slice(0, m.index + m[0].length);
    } else if (m[1] === undefined) {
      depth++;
    }
  }
  return null; // unbalanced — caller treats the paragraph as having no pPr
}

/**
 * An empty paragraph written self-closed — `<w:p/>`, `<w:p w:rsidR="…"/>`, how
 * Word writes a blank paragraph and very often the one a table cell must end
 * with — opened up as `<w:p …></w:p>`; anything else returned as it is. Every
 * function here that slices a paragraph into open tag, properties, runs and
 * close tag calls this first: read as an open tag, the self-closed form put a
 * revision OUTSIDE the paragraph (`<w:ins><w:p/></w:ins>`, which
 * `CT_RunTrackChange` does not allow) or properties after a closed element.
 */
export function openSelfClosedParagraph(paraXml: string): string {
  const m = /^(\s*)<w:p(\s[^>]*?)?\s*\/>(\s*)$/.exec(paraXml);
  return m === null ? paraXml : `${m[1]}<w:p${m[2] ?? ''}></w:p>${m[3]}`;
}

/**
 * Given a `<w:p …>…</w:p>` string, wrap the run-level content (everything that
 * is not the paragraph properties `<w:pPr>`) using the supplied wrapper and
 * return the rebuilt paragraph. Pure string surgery — no parser round-trip, to
 * avoid namespace-declaration quirks (key constraint 6).
 */
export function wrapParaRuns(xml: string, wrap: (runs: string) => string): string {
  const paraXml = openSelfClosedParagraph(xml);
  const openMatch = paraXml.match(/^(\s*<w:p(?:\s[^>]*)?>)/);
  const closeMatch = paraXml.match(/(<\/w:p>\s*)$/);
  if (!openMatch || !closeMatch) {
    // Not a recognisable paragraph — wrap the whole thing defensively.
    return wrap(paraXml);
  }
  const open = openMatch[1]!;
  const close = closeMatch[1]!;
  let inner = paraXml.slice(open.length, paraXml.length - close.length);

  // Split off a leading <w:pPr>…</w:pPr> (or self-closed) so paragraph
  // properties stay outside the revision element.
  let pPr = '';
  const pPrMatch = matchLeadingPPr(inner);
  if (pPrMatch !== null) {
    pPr = pPrMatch;
    inner = inner.slice(pPr.length);
  }

  return `${open}${pPr}${wrap(inner)}${close}`;
}

/**
 * Mark a paragraph's own PARAGRAPH MARK as inserted or deleted (ADR-0032,
 * paragraph merge / split).
 *
 * Word records the loss of a paragraph break as a revision on the paragraph
 * mark itself — `<w:pPr><w:rPr><w:del …/></w:rPr></w:pPr>` — not as a run-level
 * deletion. Accepting it joins the paragraph to the one below; rejecting it
 * restores the break. A merge authored WITHOUT this is not a redline of the
 * merge at all: the break simply vanishes and cannot be rejected.
 *
 * `CT_PPr` allows **one** `w:rPr` and orders it `rPr → sectPr → pPrChange`, so
 * this cannot simply append before `</w:pPr>`:
 *
 *  - a paragraph mark that already carries direct formatting has an `rPr`, and a
 *    second one is schema-invalid — merge into it instead (inside `CT_ParaRPr`,
 *    `w:ins`/`w:del` come first, so the revision goes at the front);
 *  - the last paragraph of a section ends with `sectPr`, and `pPrChange` may
 *    follow it — the new `rPr` must be inserted **before** either;
 *  - `pPrChange` contains a nested `<w:pPr>` (a `CT_PPrBase`, which has no
 *    `rPr` at all), so a lazy close-tag match lands the revision inside it and
 *    the paragraph mark is then not marked at all — the merge becomes
 *    unrejectable, the exact failure this function exists to prevent.
 */
export function markParagraphMark(
  xml: string,
  kind: 'ins' | 'del',
  id: number,
  author: string,
  date: string,
): string {
  const paraXml = openSelfClosedParagraph(xml);
  const rev = `<w:${kind} ${revAttrs(id, author, date)}/>`;
  const mark = `<w:rPr>${rev}</w:rPr>`;
  const openMatch = paraXml.match(/^(\s*<w:p(?:\s[^>]*)?>)/);
  if (!openMatch) return paraXml;
  const open = openMatch[1]!;
  const rest = paraXml.slice(open.length);

  const pPr = matchLeadingPPr(rest);
  // No properties, or an unbalanced fragment we will not edit blind — add a
  // fresh <w:pPr> carrying only the mark.
  if (pPr === null) return `${open}<w:pPr>${mark}</w:pPr>${rest}`;
  // Self-closed <w:pPr/> — expand it.
  if (pPr.endsWith('/>')) {
    return `${open}<w:pPr>${mark}</w:pPr>${rest.slice(pPr.length)}`;
  }

  const openTag = /^\s*<w:pPr(?:\s[^>]*)?>/.exec(pPr)![0];
  const closeTag = '</w:pPr>';
  const body = pPr.slice(openTag.length, pPr.length - closeTag.length);

  // Everything from the first sectPr/pPrChange onwards must stay AFTER the rPr.
  const at = (needle: string): number => {
    const i = body.indexOf(needle);
    return i === -1 ? body.length : i;
  };
  const boundary = Math.min(at('<w:sectPr'), at('<w:pPrChange'));
  const head = body.slice(0, boundary);
  const tail = body.slice(boundary);

  // An existing paragraph-mark rPr (only ever in `head`, since CT_PPr orders
  // rPr before sectPr) takes the revision as its first child.
  // `[^>]*` would swallow the `/` of a self-closed `<w:rPr w:x="1"/>` and treat
  // it as an open tag, putting `w:del` directly under `CT_PPr` — invalid, and
  // the paragraph mark then unmarked. The final `[^/]` forbids that.
  const paired = /<w:rPr(?:\s[^>]*[^/])?>/.exec(head);
  let merged: string;
  if (paired) {
    const cut = paired.index + paired[0].length;
    merged = head.slice(0, cut) + rev + head.slice(cut);
  } else {
    // Expand a self-closed rPr in place, KEEPING its attributes — replacing the
    // tag wholesale would silently drop them.
    const selfRPr = /<w:rPr((?:\s[^>]*?)?)\/>/.exec(head);
    merged = selfRPr
      ? head.slice(0, selfRPr.index) +
        `<w:rPr${selfRPr[1]!}>${rev}</w:rPr>` +
        head.slice(selfRPr.index + selfRPr[0].length)
      : head + mark;
  }

  return `${open}${openTag}${merged}${tail}${closeTag}${rest.slice(pPr.length)}`;
}
