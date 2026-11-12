/**
 * ADR-0045 — a seeded random sweep over paragraphs mixing text with the objects
 * the replace author keeps: equations, links, REF fields, footnote references,
 * bookmarks, text boxes (plain and with a VML fallback), content controls,
 * inline pictures and links holding only a space. Each case edits words, moves and adds spaces at object
 * edges, and adds, drops or swaps objects.
 *
 * Why it exists: three reviews in a row each found inputs that tripped
 * `inline-objects.ts`'s self-check — and a tripped check sends the paragraph
 * back to the pre-ADR rebuild, which drops every object in it. The ADR twice
 * claimed no input could, from the cases its tests happened to construct. A
 * claim over an input space has to be measured over that space: this is the
 * measurement, and every case asserts the outcome a fallback would break.
 *
 * Per case, for the whole redline:
 *  - accept-all reads as the revised document and reject-all as the original,
 *    in `normText` — the model the diff itself compares in (the document's own
 *    spelling is asserted by `pipeline.inline-objects.test.ts`, case by case;
 *    a space added beside an equation is outside the diff's model, ADR-0045
 *    § Accepted limits 12);
 *  - the object census matches both ways — so no object is lost or doubled,
 *    which is what a fallback does — and every footnote reference and drawing
 *    sits among the same characters as in its source;
 *  - no `<w:ins>`/`<w:del>` wraps a hyperlink or simple field;
 *  - no deletion is directly followed by an insertion (ADR-0043's order);
 *  - run-level revision ids open in document order.
 *
 * Synthetic vocabulary only.
 */
import { describe, it, expect } from 'vitest';
import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import { runPipeline } from '@features/compare/pipeline.js';
import { flattenTrackedChanges, extractParaRecords } from '@features/compare/preprocess/index.js';
import type { ParaRecord } from '@features/compare/model.js';
import { parseNodes, paragraphText, zeroTextObjects } from '@features/compare/preprocess/inline-units.js';
import { tagOf, childrenOf, type PoNode } from '@shared/ooxml/po-xml.js';
import {
  inlineParts, para, run, math, footnoteRef, hyperlink, refField, bookmarked, textBox, alternateTextBox,
} from '../../fixtures/inline-objects-docx.js';

const CONFIG = getCompareConfig(SG_PROFILE);
// Drafting punctuation decides how the word diff splits a space at an object's
// edge — a quoted defined term at a link's end isolates one (found in review) —
// so the vocabulary carries it.
const WORDS = [
  'the', 'Seller', 'Buyer', 'shall', 'may', 'pay', 'Price', 'portal', 'notice', 'within', 'ten', 'days', 'of',
  'Completion', '"Goods"', "Seller's", 'days.', 'Clause', 'S$100', '(the "Portal")', 'and;',
];

function rng(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

type Seg =
  | { k: 'run' | 'num' | 'link' | 'field' | 'math' | 'box' | 'altbox' | 'sdt' | 'mark'; words: string[]; lead: string; trail: string }
  | { k: 'fn' | 'pic' | 'shape' | 'spacelink' };

const picture = '<w:r><w:drawing><wp:inline><a:graphic><a:graphicData/></a:graphic></wp:inline></w:drawing></w:r>';
const shape =
  '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor/></w:drawing></mc:Choice>' +
  '<mc:Fallback><w:pict><v:shape/></w:pict></mc:Fallback></mc:AlternateContent></w:r>';

function render(seg: Seg): string {
  if (!('words' in seg)) return seg.k === 'fn' ? footnoteRef() : seg.k === 'pic' ? picture : seg.k === 'shape' ? shape : hyperlink(' ');
  const t = `${seg.lead}${seg.words.join(' ')}${seg.trail}`;
  switch (seg.k) {
    case 'run':
    case 'num': return run(t);
    case 'link': return hyperlink(t);
    case 'field': return refField(t);
    case 'math': return math(seg.words.join(''));
    case 'box': return textBox(t);
    case 'altbox': return alternateTextBox(t);
    case 'sdt': return `<w:sdt><w:sdtPr><w:alias w:val="Field"/></w:sdtPr><w:sdtContent>${run(t)}</w:sdtContent></w:sdt>`;
    case 'mark': return bookmarked(run(t));
  }
}

/**
 * `objectsOnly`: every edit adds, drops or moves a zero-width object and no
 * word changes — the shape ADR-0045 § Amendment makes a reported, marked
 * `property` op instead of an invisible `equal`.
 */
function generate(next: () => number, objectsOnly = false): { orig: string; rev: string } {
  const pick = <T,>(xs: readonly T[]): T => xs[Math.floor(next() * xs.length)]!;
  const words = (): string[] => Array.from({ length: 1 + Math.floor(next() * 4) }, () => pick(WORDS));
  const textKinds = ['run', 'run', 'run', 'link', 'field', 'math', 'box', 'altbox', 'sdt', 'mark'] as const;
  const zeroKinds = ['fn', 'pic', 'shape', 'spacelink'] as const;
  const segs: Seg[] = [];
  // A manual clause number, which `normText` strips: an object can sit before
  // it, or move across it (found in review).
  if (next() < 0.25) segs.push({ k: 'num', words: ['1.'], lead: '', trail: ' ' });
  const n = 2 + Math.floor(next() * 5);
  for (let i = 0; i < n; i++) {
    if (next() < 0.2) segs.push({ k: pick(zeroKinds) });
    else segs.push({ k: pick(textKinds), words: words(), lead: next() < 0.3 ? ' ' : '', trail: next() < 0.5 ? ' ' : '' });
  }
  const revised: Seg[] = segs.map((s) => ('words' in s ? { ...s, words: [...s.words] } : { ...s }));
  const edits = 1 + Math.floor(next() * 3);
  for (let e = 0; e < edits; e++) {
    if (revised.length === 0) revised.push({ k: 'run', words: words(), lead: '', trail: '' });
    const i = Math.floor(next() * revised.length);
    const s = revised[i]!;
    // The number itself is never edited: a manual number removed or rewritten
    // in an edited clause is emitted from the revised side unmarked (ADR-0028),
    // which loses it from the rejected view — a separate, recorded follow-up.
    // Objects may still be put before it, dropped beside it or moved across it.
    if (s.k === 'num') continue;
    if (objectsOnly) {
      // Not a link holding only a space: that is an object WITH text (ADR-0045
      // § Accepted limits 1 keeps it out of the amendment).
      const zero = revised.map((x, k) => ('words' in x || x.k === 'spacelink' ? -1 : k)).filter((k) => k >= 0);
      const roll = next();
      if (zero.length > 0 && roll < 0.4) revised.splice(pick(zero), 1);
      else if (zero.length > 0 && roll < 0.7) {
        const [moved] = revised.splice(pick(zero), 1);
        revised.splice(Math.floor(next() * (revised.length + 1)), 0, moved!);
      } else revised.splice(Math.floor(next() * (revised.length + 1)), 0, { k: pick(['fn', 'pic', 'shape'] as const) });
      continue;
    }
    const roll = next();
    if ('words' in s && roll < 0.5) s.words[Math.floor(next() * s.words.length)] = pick(WORDS);
    else if ('words' in s && roll < 0.65) s.lead = s.lead === '' ? ' ' : '';
    else if ('words' in s && roll < 0.8) s.trail = s.trail === '' ? ' ' : '';
    else if (roll < 0.9) revised.splice(i, 1);
    else revised.splice(i, 0, { k: pick(zeroKinds) });
  }
  return { orig: para(segs.map(render).join('')), rev: para(revised.map(render).join('')) };
}

/** The first top-level `<w:p>` of a body, matched by depth (a text box nests paragraphs). */
function firstParagraph(xml: string): string {
  const start = xml.indexOf('<w:p');
  const tag = /<w:p(?:\s[^>]*)?>|<\/w:p>/g;
  tag.lastIndex = start;
  let depth = 0;
  for (let m = tag.exec(xml); m !== null; m = tag.exec(xml)) {
    depth += m[0].startsWith('</') ? -1 : 1;
    if (depth === 0) return xml.slice(start, m.index + m[0].length);
  }
  return xml;
}

/**
 * ADR-0043's order, read structurally: among a paragraph's (or a container's)
 * children, no deletion is followed by an insertion with only bookmarks
 * between. A kept footnote reference or image between them is the recorded
 * exception (ADR-0045 § Accepted limits 13). A string test for
 * `</w:del><w:ins ` is blind to the bookmark case, which is how it was missed.
 */
const ORDER_CONTAINERS = new Set(['w:hyperlink', 'w:fldSimple', 'w:smartTag', 'w:customXml', 'w:sdt', 'w:sdtContent', 'm:oMath']);
function deletionBeforeInsertion(nodes: readonly PoNode[]): boolean {
  let pendingDel = false;
  for (const node of nodes) {
    const tag = tagOf(node);
    if (tag === undefined || tag === 'w:pPr' || tag === 'w:sdtPr') continue;
    if (ORDER_CONTAINERS.has(tag) && deletionBeforeInsertion(childrenOf(node))) return true;
    if (tag === 'w:del') pendingDel = true;
    else if (tag === 'w:ins') {
      if (pendingDel) return true;
    } else if (!tag.startsWith('w:bookmark')) {
      pendingDel = false;
    }
  }
  return false;
}

/**
 * Where each zero-width object sits among the characters: the text with every
 * space removed (spacing beside an object is the diff's business, limit 12)
 * and a mark for each footnote reference or drawing. A census counts objects
 * and cannot see one in the wrong place — review found a footnote
 * that never moved misplaced once every change was rejected, and every count
 * was right.
 */
function places(xml: string): string {
  const out: string[] = [];
  const walk = (nodes: readonly PoNode[]): void => {
    for (const node of nodes) {
      const tag = tagOf(node);
      if (tag === undefined || tag === 'mc:Fallback') continue;
      if (tag === 'w:footnoteReference' || tag === 'w:endnoteReference') out.push('\u0001N');
      else if (tag === 'w:drawing') out.push('\u0001D');
      if (tag === 'w:t' || tag === 'm:t') out.push(paragraphText({ 'w:r': [node] } as PoNode, false));
      else walk(childrenOf(node));
    }
  };
  walk(parseNodes(xml));
  return out.join('').replace(/\s+/g, '');
}

function census(raw: string, bookmarks: boolean): Record<string, number> {
  const shells = /<(w:hyperlink|w:fldSimple|m:oMath|w:sdtContent)(?:\s[^>]*)?(?:\/>|>(?:<w:sdtPr>.*?<\/w:sdtPr>)?<\/\1>)/g;
  let xml = raw;
  for (let prev = ''; prev !== xml; ) {
    prev = xml;
    xml = xml.replace(shells, '');
  }
  const n = (re: RegExp) => (xml.match(re) ?? []).length;
  const open = (tag: string) => n(new RegExp(`<${tag}[\\s/>]`, 'g'));
  return {
    footnote: n(/<w:footnoteReference[\s/>]/g),
    fieldBegin: n(/w:fldCharType="begin"/g),
    hyperlink: open('w:hyperlink'),
    math: open('m:oMath'),
    drawing: open('w:drawing'),
    pict: open('w:pict'),
    textBox: open('w:txbxContent'),
    sdt: open('w:sdtContent'),
    ...(bookmarks ? { bookmark: n(/<w:bookmarkStart[\s/>]/g) } : {}),
  };
}

/**
 * Object-only pairs, per 200, that fell back to striking every word and
 * inserting every word again. Measured after the first review: 0 and
 * 1 on the two object-only seeds — 148 of 171 and 148 of 169 object ops spliced
 * into the revised paragraph, the rest marked by the object-aware replace,
 * chiefly pairs holding a link of only a space, which the splice declines. A
 * rise is a regression.
 */
const WHOLE_PARAGRAPH_CEILING = 1;

describe('ADR-0045 sweep: no generated paragraph loses or doubles an object', () => {
  for (const [seed, objectsOnly] of [[20261017, false], [451, false], [9001, false], [20261018, true], [77, true]] as const) {
    it(`seed ${seed}: 200 random pairs${objectsOnly ? ', objects only' : ''}`, async () => {
      const next = rng(seed);
     
      let replaces = 0;
      let objectOps = 0;
      let wholeParagraph = 0;
      for (let c = 0; c < 200; c++) {
        const { orig, rev } = generate(next, objectsOnly);
        const r = await runPipeline({
          originalParts: inlineParts(orig),
          revisedParts: inlineParts(rev),
          config: CONFIG,
          author: 'Mutatis Compare',
          date: '2026-10-17T00:00:00Z',
        });
        replaces += r.ops.filter((o) => o.type === 'replace').length;
        const objectOp = r.ops.find((o) => o.propertyChange?.objects === true);
        // Detection is complete: a text-identical pair whose zero-text objects
        // differ is reported. Without this the census skip below would let a
        // missed one pass silently.
        const [o0, v0] = [r.originalParas[0]!, r.revisedParas[0]!];
        if (o0.normText === v0.normText) {
          const objects = (p: ParaRecord) => JSON.stringify(zeroTextObjects(p.rawXml, p.numberingPrefix));
          if (objects(o0) !== objects(v0)) expect(objectOp, `undetected, seed ${seed} case ${c}\n${orig}\n${rev}`).toBeDefined();
        }
        if (objectOp !== undefined) {
          objectOps++;
          // The last resort strikes every word of the original and inserts every
          // word of the revised paragraph; the other tries keep the words.
          const marked = (tag: 'delText' | 'ins') => {
            const xml = firstParagraph(r.authoredBodyXml);
            const texts = tag === 'delText'
              ? [...xml.matchAll(/<w:delText[^>]*>([^<]*)<\/w:delText>/g)].map((m) => m[1])
              : [...xml.matchAll(/<w:ins [^>]*>((?:(?!<\/w:ins>).)*)<\/w:ins>/g)].flatMap((m) => [...m[1]!.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((t) => t[1]));
            return texts.join('').replace(/\s+/g, '').length;
          };
          const words = r.originalParas[0]!.normText.replace(/\s+/g, '').length;
          if (words > 0 && marked('delText') >= words && marked('ins') >= words) wholeParagraph++;
        }
        // A pair still `equal` is emitted from the revised side whole. Since the
        // amendment that holds only where no footnote, picture or object was
        // added, removed or moved — a link of only a space dropped at the edge
        // changes no text and stays the recorded limit (ADR-0045 § Accepted
        // limits 1). An object op is held to every check below.
        const unchanged = r.ops.every((o) => o.type === 'equal' || (o.type === 'property' && o.propertyChange?.objects !== true));
        const where = `seed ${seed} case ${c}\n${orig}\n${rev}`;
        for (const [policy, source] of [
          ['acceptAll', r.revisedParas],
          ['rejectAll', r.originalParas],
        ] as const) {
          const flat = flattenTrackedChanges(r.authoredBodyXml, policy);
          const texts = (ps: readonly ParaRecord[]) => ps.map((p) => p.normText).filter((t) => t !== '');
          expect(texts(extractParaRecords(flat, 'document', CONFIG)), `${policy} text, ${where}`).toEqual(texts(source));
          const accept = policy === 'acceptAll';
          if (unchanged) continue;
          expect(census(flat, accept), `${policy} census, ${where}`).toEqual(
            census(source.map((p) => p.rawXml).join(''), accept),
          );
          expect(places(flat), `${policy} places, ${where}`).toBe(places(source.map((p) => p.rawXml).join('')));
        }
        expect(r.authoredBodyXml, where).not.toMatch(/<w:(ins|del) [^>]*><w:(hyperlink|fldSimple)/);
        // ADR-0043: a changed region reads insertion first, however the walk
        // made its revisions.
        const body = childrenOf(parseNodes(r.authoredBodyXml)[0]!);
        expect(body.filter((p) => tagOf(p) === 'w:p').some((p) => deletionBeforeInsertion(childrenOf(p))), where).toBe(false);
        const ids = [...r.authoredBodyXml.matchAll(/<w:(?:ins|del) w:id="(\d+)"[^>]*[^/]>/g)].map((m) => Number(m[1]));
        expect(ids, where).toEqual([...ids].sort((a, b) => a - b));
      }
      // The sweep is only evidence if it reaches the author it is testing.
      if (objectsOnly) {
        expect(objectOps).toBeGreaterThan(100);
        // Measured; a rise means more pairs struck whole than the splice should leave.
        expect(wholeParagraph).toBeLessThanOrEqual(WHOLE_PARAGRAPH_CEILING);
      } else {
        expect(replaces).toBeGreaterThan(100);
      }
    });
  }
});
