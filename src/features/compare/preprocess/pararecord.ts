/**
 * Build ParaRecord[] from a (flattened, quarantined) document part.
 * Spec: spec/features/08-compare.md §4.1, §5.0, §5.2, §7.2, §9.6.
 *
 * Walks <w:p> elements in document order; for each paragraph extracts its
 * visible text (`inline-units.ts → paragraphText`: `<w:t>`, OMML `<m:t>`, field
 * results, text-box content — ADR-0045), normalises it, strips the numbering prefix, hashes the
 * prefix-free text (FNV-1a), and records structural flags (table cell, header/
 * footer, footnote, signature block). Pure — author/date are not involved here.
 */
import { XMLParser, XMLBuilder } from 'fast-xml-parser';

import type { CompareConfig } from '@shared/config/index.js';
import {
  resolveParaNumPr,
  resolveLevel,
  levelsReferencedBy,
  rendersToken,
} from '@shared/numbering/index.js';
import { PRESERVE_ORDER_OPTS, type PoNode as OoxmlPoNode } from '@shared/ooxml/po-xml.js';

import type { ParaNumbering, ParaRecord } from '../model.js';

import type { NumberingContext } from './numbering.js';

import { fnv1a } from './hash.js';
import { normaliseText, normaliseTextAggressive, stripNumberingPrefix } from './normalise.js';
import { BUILDER_OPTS, paragraphText } from './inline-units.js';

/**
 * The SHARED parser options, not a local copy of them.
 *
 * `extractPPr` hands its nodes to `@shared/ooxml/po-xml`'s `tagOf`/`attrOf` and
 * to `@shared/numbering`'s resolver, both of which assume the node shape those
 * exact options produce. A second literal that merely happened to match made
 * that coupling invisible: change `attributeNamePrefix` in one place and the
 * numbering resolution here silently reads `undefined` from every attribute
 * while the text extraction keeps working. Reuse before build — one home.
 */
const PARSER_OPTS = PRESERVE_ORDER_OPTS;


type PoNode = Record<string, unknown>;

/** §7.2 signature/execution-block keywords (lowercased match against para text). */
const SIGNATURE_KEYWORDS = [
  'executed',
  'signed',
  'as a deed',
  'in the presence of',
  'director',
  'secretary',
  'witness',
  'common seal',
  'authorised signatory',
  'authorized signatory',
];

/**
 * §7.2 — heuristic: is this paragraph text part of a signature/execution block?
 * Shared by the OOXML extractor (below) and the PDF adapter (io/pdf) so both
 * sides flag execution blocks identically. Input may be raw or normalised text.
 */
export function isSignatureBlockText(text: string): boolean {
  const lower = text.toLowerCase();
  return SIGNATURE_KEYWORDS.some((kw) => lower.includes(kw));
}

/**
 * Extract ParaRecord[] from a document part's XML.
 *
 * @param documentXml  the part XML AFTER flatten + quarantine.
 * @param partId       e.g. 'document' | 'header1' | 'footer1' | 'footnotes'.
 * @param config       compare config (reserved for future tuning; unused here).
 * @param numbering    styles + numbering scheme for THIS side (ADR-0042). When
 *                     omitted, every record's `numbering` is null — the exact
 *                     behaviour before ADR-0042, so a caller that has no parts
 *                     to give (text records, tests) is unaffected.
 */
export function extractParaRecords(
  documentXml: string,
  partId: string,
  config: CompareConfig,
  numbering?: NumberingContext,
): ParaRecord[] {
  void config; // reserved — kept in signature per the A1 contract.

  const parser = new XMLParser(PARSER_OPTS);
  const tree = parser.parse(documentXml) as PoNode[];
  const builder = new XMLBuilder(BUILDER_OPTS);

  const isHeaderFooter = partId.startsWith('header') || partId.startsWith('footer');
  const isFootnote = partId === 'footnotes' || /^(?:foot|end)note:/.test(partId);

  const records: ParaRecord[] = [];
  let index = 0;

  // Walk in document order, carrying whether we're inside a <w:tc>.
  const walk = (nodes: PoNode[], insideTableCell: boolean): void => {
    for (const node of nodes) {
      const tag = tagOf(node);
      if (tag === undefined) continue;

      if (tag === 'w:p') {
        const rawXml = builder.build([node]) as string;
        const rawText = paragraphText(node);
        const pPr = extractPPr(node);
        const styleId = extractStyleId(pPr);

        const normWhole = normaliseText(rawText);
        const { prefix, body } = stripNumberingPrefix(normWhole);
        const normText = body;
        const normTextAggressive = normaliseTextAggressive(body);

        const isSignatureBlock = isSignatureBlockText(normWhole);

        records.push({
          index: index++,
          partId,
          rawXml,
          normText,
          normTextAggressive,
          numberingPrefix: prefix,
          hash: fnv1a(normText),
          styleId,
          // ADR-0042 — identity and comparison key; the LABEL needs the whole
          // part in document order and arrives from `assignNumberLabels`.
          numbering: numbering ? resolveNumbering(numbering, styleId, pPr) : null,
          isTableCell: insideTableCell,
          isSignatureBlock,
          isHeaderFooter,
          isFootnote,
        });
        // A paragraph can itself contain nested content, but not nested <w:p>;
        // no need to recurse into it for paragraph discovery.
        continue;
      }

      // Recurse into children, tracking table-cell context.
      const value = node[tag];
      if (Array.isArray(value)) {
        walk(value as PoNode[], insideTableCell || tag === 'w:tc');
      }
    }
  };

  walk(tree, false);
  return records;
}

/**
 * Resolve a paragraph's numbering identity: the instance and level the style
 * chain resolves to, plus the cross-document comparison key. The label and the
 * counter path are a property of the SEQUENCE and arrive later, from
 * `assignNumberLabels`. See `numberingKey` for what is comparable and why.
 */
function resolveNumbering(
  ctx: NumberingContext,
  styleId: string | null,
  pPr: readonly OoxmlPoNode[],
): ParaNumbering | null {
  const ref = resolveParaNumPr(ctx.styles, styleId, pPr);
  if (ref === null) return null;
  return {
    numId: ref.numId,
    ilvl: ref.ilvl,
    key: numberingKey(ctx, ref.numId, ref.ilvl),
    label: null,
    path: null,
    renders: renderedLevels(ctx, ref.numId, ref.ilvl),
  };
}

/**
 * The levels 0…`ilvl` whose counters this paragraph's label actually SHOWS.
 *
 * Two filters, and each one was a measured header-versus-rows contradiction:
 *
 * 1. **Substituted.** `levelsReferencedBy` reads the same `%N` tokens
 *    `applyLvlText` does, and the intersection with 0…`ilvl` is deliberate: a
 *    `%3` in a level-1 template references a counter the renderer has no token
 *    for and renders `''` on both sides.
 * 2. **Rendered non-empty.** A substituted level whose effective format is
 *    `none`, `bullet` or anything `formatInt` cannot express occupies its `%N`
 *    slot and prints nothing. Splitting a clause list under `%1.%2` with level 0
 *    at `numFmt="none"` moved three sub-clauses `.c`→`.a`, `.d`→`.b`, `.e`→`.c`
 *    — minus two each — under a header reading *"shifted (+1)"*, the invisible
 *    level-0 counter. Same defect as filter 1, one step further in.
 *
 * The effective format is the one the renderer computes, `isLgl` folded in
 * exactly as `numberingKey` folds it — `isLgl` forces every substituted token to
 * `decimal`, which always renders, so it can turn an unprintable level printable.
 *
 * Both filters read the DEFINITION only, never a counter, which is what lets one
 * side's answer stand for the pair (equal identity ⇒ equal template and equal
 * formats). The residue is stated in `rendersToken`: a letter or roman level
 * whose counter is at or below zero also prints nothing, and no definition-level
 * predicate can see that.
 *
 * Null when the level does not resolve — there is no label, so there is nothing
 * it is rendered from, and `path` is null too.
 */
function renderedLevels(
  ctx: NumberingContext,
  numId: string,
  ilvl: number,
): readonly number[] | null {
  const own = resolveLevel(ctx.scheme, numId, ilvl);
  if (own === undefined) return null;
  const referenced = levelsReferencedBy(own.lvlText);
  const out: number[] = [];
  for (let k = 0; k <= ilvl; k++) {
    if (!referenced.has(k)) continue;
    const level = k === ilvl ? own : resolveLevel(ctx.scheme, numId, k);
    const fmt = own.isLgl ? 'decimal' : (level?.numFmt ?? 'decimal');
    if (!rendersToken(fmt)) continue;
    out.push(k);
  }
  return out;
}

/**
 * The field separator. A C0 control character, **illegal in XML 1.0 content**,
 * so a conforming `numbering.xml` cannot put it inside a `lvlText` or a `numFmt`
 * and no two different definitions can collide on one key. A delimiter a value
 * *could* contain would make a collision possible.
 *
 * What a collision costs, stated in both directions because this comment has
 * been wrong in each of them. Where the two definitions render DIFFERENT labels
 * on a pair this pass examines, a collision misattributes and cannot hide:
 * `numberLabel` is compared from the two RENDERED labels and never from the key,
 * so the op survives but reads *"its number moved because a sibling did"* when
 * the list was in fact redefined, and §8.2 hides one of those two by default.
 * Where they render the SAME label — two `numFmt` values this build cannot
 * express, at different `start`s — the key is the only detector there is, and a
 * collision drops the op outright: `summary.total` 1 to 0. That IS a change
 * found and then not reported, and it is why this byte is load-bearing rather
 * than tidy. Both directions are pinned by tests that empty this constant.
 *
 * **The collision those tests use moved once, and the move is the lesson.**
 * Until the twenty-second pass they collided the OWN level's `numFmt` against
 * its `start` — `fooX`+`11` against `fooX1`+`1`, which concatenate the same.
 * Adding the own level's declared format as a trailing field (see
 * `numberingKey`) broke that: a difference in the LAST field has nothing after
 * it to absorb the boundary, so the two keys stayed apart with the separator
 * emptied and both tests went on passing while testing nothing. The suite could
 * not say so — running it proved they still passed, not that they still had a
 * premise. The collision now lives in an ANCESTOR level's block, which carries
 * no trailing declared format: level 0 at `fooX`+`11` against `fooX1`+`1` under
 * a `%1.%2` own level identical on both sides. When a change widens this key,
 * re-derive every guard's premise, not just its pass/fail.
 *
 * **There was a second constant here, and it could never have had a guard.**
 * The key is `${ilvl}` + separator + the fields, so emptying a separator that
 * appears ONLY there can change an answer only when one side's `ilvl` decimal
 * string is a proper prefix of the other's — impossible while both are single
 * digits, and Word's nine levels (`ilvl` 0–8) are all single digits. Same
 * `ilvl` ⇒ the two keys were equal or unequal already; different single-digit
 * `ilvl` ⇒ they differ at index 0. So it decided nothing, no ablation could
 * turn it red, and seven documents said it was pinned. The level is now an
 * ordinary FIRST FIELD behind this one separator, which is what the documents
 * always described.
 *
 * (The premise is the producer's conformance, not the reader's:
 * `fast-xml-parser` does not validate, so a raw control byte in a malformed
 * package would survive it — and would also land in the rendered label.)
 */
const FIELD_SEP = '\u0001';

/**
 * The cross-document comparison key for one paragraph's numbering: its level,
 * plus the definition fields its rendered number is actually drawn from.
 *
 * -- What it excludes, and why ------------------------------------------------
 * **`numId`.** A `w:num` instance id is document-local and Word reassigns it
 * freely — a Save-As onto another template renumbers every instance while every
 * clause keeps its number — so comparing it across two packages reports changes
 * that did not happen (measured on a synthetic pair with no real change).
 *
 * **Anything the renderer does not read for THIS paragraph.** The key answers
 * one question: did this paragraph's own numbering change, or did its number
 * merely move because a sibling did? Every field in it must therefore be one
 * that can alter what `renderListNumbers` produces here. Read that module
 * alongside this one — the label is
 *
 *     applyLvlText(own.lvlText, tokens)
 *     tokens[k] = formatInt(counter_k, own.isLgl ? 'decimal' : level_k.numFmt)
 *
 * so the fields that can move it are: the OWN level's `lvlText`; and, for each
 * level the template actually substitutes, that level's EFFECTIVE format (its
 * `numFmt`, or `decimal` when the own level's `isLgl` overrides every token) and
 * whatever sets its counter — `start` and `lvlRestart`.
 *
 * Plus one field that is not in that reading of the renderer, and is here
 * because the renderer is wrong: the OWN level's DECLARED `numFmt`, carried
 * beside the folded effective format — for a substituted own level only, like
 * every other format field, because an own format the `lvlText` never
 * substitutes moves no label in this build or in Word (measured: under
 * `lvlText` `%1.`, a level-1 paragraph's own `decimalZero` → `decimal` gives
 * `summary.total` 0; under `%1.%2` the same edit gives 2). `render-list-numbers.ts` applies `isLgl`
 * to the paragraph's own token as well as its ancestors', Word does not, and a
 * key that mirrors that fold goes blind exactly where the renderer does —
 * measured, a `decimalZero` to `decimal` edit under `<w:isLgl/>` produced NO OP
 * AT ALL while Word shows `1.01` against `1.1`. See the field at `k === ilvl`.
 *
 * `isLgl` and `lvlRestart` are stored as the behaviour the renderer computes
 * rather than as written, because several spellings of each mean the same thing
 * and a key that separates them reports a document that renders identically —
 * with the one exception above, where the behaviour computed is a defect and the
 * declared value is carried as well. *Whose* behaviour matters: the document's
 * is what the key describes, this build's is what it must not inherit.
 *
 * One field is NOT about this paragraph's own token: whether each level 0…`ilvl`
 * has a definition at all. The counter machine SKIPS a paragraph whose level does
 * not resolve, so deleting a level definition moves the counters of the
 * paragraphs below it. See the loop.
 *
 * -- What this key covers, exactly ------------------------------------------
 * What the renderer reads and writes WHEN RENDERING THIS PARAGRAPH: levels
 * 0…`ilvl`, whose formats it substitutes and whose counters it touches. Levels
 * DEEPER than `ilvl` are touched by other paragraphs' render calls, and a
 * per-paragraph key is not a function of those. The counter machine is a
 * sequence, and a sequence's state is not a property of any one member of it.
 *
 * That boundary has a visible consequence, recorded and tested rather than
 * papered over (`compare.diff.property.test.ts`, *"RECORDS its limit"*;
 * ADR-0042 § Accepted limits): deleting a level definition deeper than a
 * paragraph's own level moves that paragraph's counter — a level-2 point
 * initialises level 1's counter before a level-1 sub-clause reaches it, so
 * `1.2` → `1.1` — with nothing in its own key moving. It is reported, because
 * `numberLabel` is compared from the rendered labels; it is reported as a
 * bystander of a sibling's renumbering, and no sibling renumbered.
 *
 * The honest boundary: this is what the renderer reads, PLUS whatever `formatInt`
 * cannot express. Two `numFmt` values it does not support (`cardinalText`, say)
 * both render `''`, so swapping one for the other changes the key and no label.
 * That is left reporting — the definition really did change, and Word would show
 * it — but it is a place where the key is wider than this build's output.
 *
 * -- Three corrections got it here, all one mistake at different widths -------
 * A key NARROWER than the label leaves two different numberings equal, and the
 * change is then reported as a bystander cascade — "its number moved because a
 * sibling did" — of which the document contains no instance. A key WIDER than
 * the label reports paragraphs that render identically end to end. The first is
 * a false statement; the second inflates the §8.2 headline count. Measured:
 *
 * - `numFmt`/`lvlText`/`isLgl` only: `<w:start>` 1 to 5 read as a +4 cascade;
 * - the own level only: redefining level 0 of a `%1.%2` list read as twelve
 *   bystander cascades, with no sibling anywhere;
 * - every field of every level 0 to `ilvl`: re-wording a level-0 prefix `%1.` to
 *   `Clause %1` reported **240** changes on a 40x5 document where 40 paragraphs
 *   changed, because an ancestor's `lvlText` is never substituted into a
 *   descendant's label and an ancestor's `isLgl` is never read for it.
 *
 * Narrowing cannot hide a label that MOVED **on a pair `detectPropertyChanges`
 * examines**: `numberLabel` is compared from the two RENDERED labels, never from
 * this key, so for those the key decides only which of the two things it was.
 * Outside that reach the sentence says nothing — see below.
 *
 * It is not a licence to narrow further, and this comment has twice read as one.
 * A definition change that moves no rendered label is detected by this key and
 * by NOTHING else — measured on a level-1 paragraph of a `%1.%2` list whose
 * LEVEL-0 definition carries two unexpressible `numFmt` values at different
 * `start`s: the `%1` token renders `''` on both sides, both labels read `.1`,
 * and with the field separator emptied the op disappears entirely,
 * `summary.total` 1 to 0.
 *
 * There is no bridging rule that makes the narrowing automatically safe, and two
 * attempts to write one were falsified by measurement — the second claimed this
 * key's span is a superset of everything that can move this paragraph's label,
 * which ADR-0042 § Accepted limits TWO has always contradicted (a level deleted
 * DEEPER than this paragraph moves its counter, is not in this key, and moves
 * the label anyway). The two true statements are: **on a pair this pass
 * examines**, a label that MOVED is always reported, whatever this key carries;
 * and a change that moves no label is reported only if this key carries the
 * field. Every field below is a judgement against the recorded costs, not a
 * deduction.
 *
 * That first qualifier was dropped three times running, each time by the pass
 * correcting the previous one, and pass 17 falsified the unqualified form with
 * ADR-0042's Accepted limit FOUR: `detectPropertyChanges` converts `equal` ops
 * only, so a clause whose words were edited AND whose number moved reports the
 * text change and nothing about the number. Measured, and now pinned by a test.
 * See ADR-0042.
 */
function numberingKey(ctx: NumberingContext, numId: string, ilvl: number): string {
  const own = resolveLevel(ctx.scheme, numId, ilvl);
  // No resolvable definition means no rendered label, so there is nothing to
  // describe but the level itself, which the style chain always gives us.
  if (own === undefined) return `${ilvl}${FIELD_SEP}`;

  // The template is read from THIS level only. `isLgl` is NOT a field: it exists
  // to force every substituted token to `decimal`, so its whole effect is folded
  // into the per-level format below. Storing it raw as well reported four
  // paragraphs when `<w:isLgl/>` was added to a list whose formats were already
  // decimal — a definition edit that renders nothing differently.
  //
  // But folding it over the OWN level's token as well hid a change outright, and
  // that was hard to find. See the `own.isLgl` field pushed
  // at `k === ilvl` below.
  const fields: string[] = [own.lvlText];

  const referenced = levelsReferencedBy(own.lvlText);
  for (let k = 0; k <= ilvl; k++) {
    const level = k === ilvl ? own : resolveLevel(ctx.scheme, numId, k);
    // WHETHER level k has a definition at all, for every k — substituted or not.
    // This is the one input that is not about this paragraph's own token: the
    // counter machine SKIPS a paragraph whose level does not resolve
    // (`render-list-numbers.ts`, `if (!level) continue`), so it neither advances
    // that level nor resets the deeper ones. Deleting a level definition
    // therefore moves the counters of paragraphs BELOW it while leaving their own
    // fields untouched. Measured: removing level 1's `<w:lvl>` from a 3-level
    // list moved two `(%3)` sub-points `(a)`→`(c)`, and with only the substituted
    // levels in the key those read as `crossRefRenumber` — "its number moved
    // because a sibling did" — when no sibling had moved and the list had been
    // redefined.
    //
    // The loop stops at `ilvl` and not at the scheme's last level, which is the
    // recorded limit above: a level DEEPER than this paragraph's is read by other
    // paragraphs' render calls, not by this one's, and widening the key to cover
    // them would mark every member of a list on a definition edit no member
    // renders.
    fields.push(String(k), level === undefined ? 'absent' : 'present');
    // A level the template never substitutes cannot appear in this label. Its
    // counter still moves and `path` still sees that — but a counter nobody
    // renders here is not a change to this paragraph.
    if (!referenced.has(k)) continue;
    fields.push(
      // The EFFECTIVE format for this level's token, exactly as
      // `render-list-numbers.ts` computes it: the own level's `isLgl` renders
      // every substituted token `decimal`, so no level's own `numFmt` is visible
      // through it.
      own.isLgl ? 'decimal' : (level?.numFmt ?? 'decimal'),
      // What sets this level's counter. `resolveLevel` folds any
      // `<w:startOverride>` into `start`, so a `lvlOverride` is covered here.
      // The renderer defaults an absent value to 1, so this must too — otherwise
      // a level whose definition is gone differs from one that says `start=1`
      // in the key while rendering identically.
      String(level?.start ?? 1),
      // `lvlRestart` as its EFFECTIVE behaviour, not as written. See
      // `effectiveRestart`.
      String(effectiveRestart(level?.lvlRestart, k)),
    );
    // THE OWN LEVEL'S DECLARED FORMAT, when `isLgl` would otherwise erase it.
    //
    // `render-list-numbers.ts` applies `isLgl` to every substituted token
    // INCLUDING the paragraph's own, and Word does not: measured against
    // Word's rendering, `numFmt=decimalZero` + `isLgl` over an `upperRoman`
    // ancestor gives `1.01` — the ancestor converted, the own format kept. That
    // renderer defect is recorded and is 08c's to fix.
    //
    // What was NOT recorded, and is this key's problem rather than the
    // renderer's: the field above mirrored the fold, so both detectors went
    // blind at once and a change to the own level's `numFmt` produced NO OP AT
    // ALL. Measured — four paragraphs, bodies byte-identical, `decimalZero` to
    // `decimal` on an `isLgl` sub-level:
    //
    //     with `<w:isLgl/>`     labels 1.1 1.2 1.3 -> 1.1 1.2 1.3   total 0
    //     the same edit without labels 1.01 …      -> 1.1 …         total 3
    //
    // Word renders the first pair `1.01 …` against `1.1 …`, so a reader sees a
    // change on every sub-clause and the summary said nothing. ADR-0008
    // suppression, and on a shape that exists three times over in the benchmark
    // packages' own `numbering.xml`.
    //
    // Pushed UNCONDITIONALLY, and the first attempt at this fix pushed it only
    // when `own.isLgl` was set — which made the presence of `<w:isLgl/>` itself
    // visible and invented a change on a list whose formats were already
    // decimal, the exact false positive the comment at the top of this function
    // records. Unconditional, the field says only *"what Word renders this token
    // as"*: identical on both sides unless the own level's declared format
    // actually changed. It duplicates the folded value above for every list that
    // has no `<w:isLgl/>`, which is most of them, and that redundancy is the
    // price of the field carrying one fact rather than two.
    //
    // The folded field above stays exactly as it is: it is what makes key
    // equality force `renders` equality, which `counterShift` relies on when it
    // reads one side's `renders`.
    if (k === ilvl) fields.push(level?.numFmt ?? 'decimal');
  }
  return `${ilvl}${FIELD_SEP}${fields.join(FIELD_SEP)}`;
}

/**
 * `<w:lvlRestart>` reduced to the behaviour the renderer actually gets, because
 * several written values mean the same thing and a key that distinguishes them
 * reports a change to a document that renders identically.
 *
 * `render-list-numbers.ts` resolves it as `def?.lvlRestart ?? deeperIlvl` and
 * then asks `advancedIlvl <= restart - 1`, where `advancedIlvl` only ever ranges
 * over `0…k-1`. So at level `k`: `0` means never restart; any value at or above
 * `k` — **and an absent element** — means restart whenever any shallower level
 * advances; only `1…k-1` are distinguishable from each other. At level 0 every
 * value collapses to "never", there being nothing shallower to restart it
 * (ECMA-376 §17.9.10).
 *
 * Measured before this normalisation: adding `<w:lvlRestart w:val="1"/>` to a
 * level-1 definition — a semantic no-op — reported **twenty** `formatting`
 * changes with both label lists byte-identical. Same shape as storing `isLgl`
 * raw beside the format, and the same fix: store what the renderer computes.
 */
function effectiveRestart(raw: number | undefined, ilvl: number): number {
  if (raw !== undefined && raw <= 0) return 0;
  return Math.min(raw ?? ilvl, ilvl);
}

/**
 * The children of a paragraph's `<w:pPr>`, or []. Read once per paragraph and
 * shared by the style-id read and the numbering resolution — this module parses
 * with `PRESERVE_ORDER_OPTS` itself (see `PARSER_OPTS`), so the nodes are by
 * construction the shape that module's helpers expect.
 */
function extractPPr(paraNode: PoNode): OoxmlPoNode[] {
  const children = paraNode['w:p'];
  if (!Array.isArray(children)) return [];
  for (const child of children as PoNode[]) {
    if (tagOf(child) !== 'w:pPr') continue;
    const pPrChildren = child['w:pPr'];
    return Array.isArray(pPrChildren) ? (pPrChildren as OoxmlPoNode[]) : [];
  }
  return [];
}

/** Read `<w:pStyle w:val="…">` from a paragraph's pPr children; null when absent. */
function extractStyleId(pPr: readonly OoxmlPoNode[]): string | null {
  for (const pPrChild of pPr) {
    if (tagOf(pPrChild as PoNode) !== 'w:pStyle') continue;
    const attrs = (pPrChild as Record<string, unknown>)[':@'] as
      | Record<string, unknown>
      | undefined;
    const val = attrs?.['@_w:val'];
    return typeof val === 'string' ? val : null;
  }
  return null;
}

/** The element tag of a preserveOrder node (single non-attrs key), or undefined. */
function tagOf(node: PoNode): string | undefined {
  for (const key of Object.keys(node)) {
    if (key === ':@') continue;
    return key;
  }
  return undefined;
}
