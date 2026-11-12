import type { DocxPara, TableCoord } from './read-docx.js';
import {
  compose,
  nextLetter,
  numberShape,
  stripLabelPunct,
  tidyLabel,
  typedNumber,
  type NumberShape,
} from './labels.js';
import type {
  ContractStructure,
  LabelSource,
  NodeKind,
  ParagraphStructure,
  Region,
  StructureNode,
} from './model.js';

/**
 * The contract-structure analyser: one pass over the paragraphs with a stack of
 * open nodes. Word's own numbering is the main evidence for depth; typed
 * numbers, heading styles and bold run-in titles are read and placed too.
 *
 * Design borrowed from Donna's clause map (`app/structure/tree.py`) and the
 * Matter add-in's dissector built on it: every paragraph is owned (an unnumbered
 * paragraph belongs to the deepest open node), front matter, parties, recitals,
 * definitions and the execution block are nodes, and a clause's citation is
 * composed from its parents so a limb of a definition is "1.1(a)" and a
 * paragraph of a schedule is "Schedule 2 para 3(a)".
 */

const CONTAINER_WORDS: Record<string, NodeKind> = {
  schedule: 'schedule',
  annex: 'annex',
  annexure: 'annex',
  exhibit: 'annex',
  appendix: 'appendix',
  attachment: 'appendix',
};

const WORD_IDS =
  'one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|first|second|third|fourth|fifth|sixth|seventh|eighth|ninth|tenth';
const CONTAINER_RE = new RegExp(
  `^\\s*(schedule|annexure|annex|appendix|exhibit|attachment)\\s+((?:[0-9]{1,3}[A-Za-z]?|[A-Z]|[IVXLC]{1,6})(?:\\([A-Z0-9]{1,2}\\))?|${WORD_IDS})(?![\\w-])(?!\\.\\d)\\s*[-–—:.]?\\s*(.*)$`,
  'is',
);
/** "SCHEDULE: DELIVERABLES", "THE SCHEDULE", "SCHEDULE" — a container title with no number. */
const CONTAINER_NOID_RE =
  /^\s*(?:the\s+)?(schedule|annexure|annex|appendix|exhibit)\s*(?:[:\-–—.]\s*(.+))?$/i;
const CONTAINER_WORD_START = /^\s*[[(]?\s*(schedule|annexure|annex|appendix|exhibit|attachment)\b/i;
const PART_RE = /^\s*part\s+([0-9]{1,3}[A-Za-z]?|[A-Z]|[IVXLC]{1,6})(?![\w-])(?!\.\d)\s*[-–—:.]?\s*(.*)$/is;

const PARTIES_HEADING = /^\s*(?:the\s+)?parties\s*:?\s*$/i;
const BETWEEN_LINE = /^\s*between\s*:?\s*$/i;
/** "This Cover Page identifies the Parties to this Agreement:", "The parties to this agreement are:". */
const PARTIES_INTRO = /^(?=.{0,120}$).*\bparties\b[^.]{0,40}:\s*$/i;
const RECITALS_HEADING = /^\s*(?:background|recitals?|whereas|preamble|introduction)\s*:?\s*$/i;
const OPERATIVE_HEADING =
  /^\s*(?:agreed terms|operative (?:provisions|terms)|it is (?:hereby )?agreed(?: as follows)?|now (?:it is|this deed witnesses)[^.]*|the parties (?:hereby )?agree as follows|.{0,60}\bagree as follows)\s*[:.]?\s*$/i;
const EXECUTION_START =
  /^\s*(?:in witness (?:whereof|of which|hereof)|as witness|this (?:deed|agreement|document) (?:has been|was|is) (?:executed|signed|entered)|executed (?:as an? \[?#*\s*(?:deed|agreement)|by|on behalf)|signed (?:by|for and on behalf|on behalf|as a deed)|execution(?: page| pages|\s+of\s+the\s+agreement)?\s*$|signatories\s*$|signature page|signed\s*$)/i;
/** Inside a schedule, only these open the agreement's execution block; "Signed for and on behalf" may be a form's. */
const EXECUTION_STRONG =
  /^\s*(?:in witness|as witness|executed (?:as an? \[?#*\s*(?:deed|agreement))|this (?:deed|agreement|contract) (?:has been|was|is) (?:executed|signed)|execution(?: page| pages)?\s*$|signatories\s*$|signature page)/i;
const SIGNATURE_WORDS =
  /\b(?:signed|signature|signatory|director|secretary|witness|name|title|date|for and on behalf|authori[sz]ed|position|print|seal)\b/i;

const RUNIN_MAX = 80;

interface Frame {
  node: StructureNode;
  listGroup: string | null;
  ilvl: number | null;
  shape: NumberShape | null;
  /** Rank of an unnumbered heading-styled title (outline level), 9 for a bold run-in. */
  headingRank: number | null;
  /** Can a numbered paragraph sit beneath this node? Headings and top-level parts: no. */
  parentable: boolean;
}

function frame(node: StructureNode, extra: Partial<Frame> = {}): Frame {
  return {
    node,
    listGroup: null,
    ilvl: null,
    shape: null,
    headingRank: null,
    parentable: true,
    ...extra,
  };
}

function textOf(p: DocxPara): string {
  return p.inAccepted && p.text.trim() !== '' ? p.text : p.textOriginal || p.text;
}

function flat(s: string): string {
  return s.replace(/[\s\u00a0]+/g, ' ').trim();
}

/** A line that reads as a title: short, no sentence punctuation at the end, starts with a capital. */
function titleLike(s: string, maxWords = 12): boolean {
  const t = flat(s);
  if (t === '' || t.length > 100) return false;
  if (t.split(' ').length > maxWords) return false;
  if (/[.;,]$/.test(t) && !/\b(?:etc|Ltd|Inc|Co)\.$/.test(t)) return false;
  if (/[;,]\s*(?:and|or|but)$/i.test(t)) return false; // the tail of a list item, not a title
  if (/^(?:or|and|and\/or|either|but)$/i.test(t)) return false; // "OR" between alternative drafting
  if (/\b(?:shall|must|will)\b/.test(t)) return false;
  return /^[“"'(]*[A-Z0-9]/.test(t);
}

function isUpperTitle(t: string): boolean {
  return /[A-Z]{2}/.test(t) && t === t.toUpperCase() && !/[a-z]/.test(t);
}

/** Formatting that marks a line as a title: a heading style, bold, capitals or centring. */
function titleEvidence(p: DocxPara, ft: string): boolean {
  return (
    p.isHeadingStyle ||
    p.allBold ||
    p.allCaps ||
    isUpperTitle(ft) ||
    /heading|title|schedule|annex|appendix|part/i.test(p.styleName ?? '')
  );
}

const QUOTE_PAIRS: Record<string, string> = {
  '“': '”',
  '"': '"',
  '‘': '’',
  "'": "'",
  '«': '»',
  '„': '“',
};
const DEFINING_VERB =
  /^\s*(?:\[[^\]]{0,800}\]\s*)?(?:\([^)]{0,80}\)\s*)?(?:(?:,|and|or)\s*[“"‘'][^”"’']{1,80}[”"’']\s*)*(?:,\s*)?(?:means?|shall mean|includes?|shall include|has the meaning|have the meanings?|has the same meaning|refers? to|is defined|shall have the meaning|shall be construed|:|–|—|-\s)/i;

/** The term a definition paragraph defines: "“Business Day” means ..." -> "Business Day". */
function definitionTerm(text: string): string | null {
  // an optional definition in square brackets: '["Temporary Land" means ...]'
  const t = text.replace(/^\s+/, '').replace(/^\[(?=[“"‘'])/, '');
  const open = t.charAt(0);
  const close = QUOTE_PAIRS[open];
  if (!close) return null;
  // the closing quote is the first one followed by something other than a letter (an apostrophe is not a close)
  let end = -1;
  for (let k = 1; k < Math.min(t.length, 120); k += 1) {
    const ch = t.charAt(k);
    if (ch === close || (close === '’' && ch === "'") || (close === '”' && ch === '"')) {
      if (/[A-Za-z]/.test(t.charAt(k + 1)) && (ch === '’' || ch === "'")) continue;
      end = k;
      break;
    }
  }
  if (end <= 1) return null;
  const term = t.slice(1, end).trim();
  if (term === '' || term.length > 90 || /[.;]$/.test(term)) return null;
  if (!DEFINING_VERB.test(t.slice(end + 1))) return null;
  return term;
}

const BOLD_TERM_VERB =
  /^\s*(?:means?|shall mean|includes?|has the meaning|have the meanings?|has the same meaning|refers? to)\b/i;

/** A bold term with no quotes, as Australian forms write them: "**Claim** includes any claim ...". */
function boldTerm(p: DocxPara, ft: string): string | null {
  const lead = flat(p.boldLead);
  if (lead === '' || !ft.startsWith(lead) || lead.split(' ').length > 8 || /[.;:,]$/.test(lead)) return null;
  if (!/^[A-Z“"‘']/.test(lead)) return null;
  // a capitalised word the bold missed ("**Adjustment Event** Table means") is part of the term
  const missed = /^(\s+[A-Z][\w-]*){1,2}(?=\s+(?:means?|includes?)\b)/.exec(ft.slice(lead.length))?.[0] ?? '';
  if (!BOLD_TERM_VERB.test(ft.slice(lead.length + missed.length))) return null;
  return (lead + missed).replace(/^[“"‘']+|[”"’']+$/g, '').trim();
}

const MADE_LINE =
  /^\s*this\s+(?:contract|agreement|deed|lease|licen[cs]e)\b.{0,60}\b(?:is made|is dated|dated|made on)\b/i;

/**
 * Where a numbered cover ends: the parties block ("BETWEEN:", "PARTIES", "THIS AGREEMENT is made
 * on") opens after numbered lines (cover notes, a contents list) that are not clauses of the
 * agreement — nothing before the parties is citable. Returns 0 when the cover has no numbers.
 */
function contentsListEnd(paras: readonly DocxPara[]): number {
  const limit = Math.min(paras.length, 400);
  let numbered = 0;
  for (let i = 0; i < limit; i += 1) {
    const p = paras[i]!;
    const t = flat(textOf(p));
    if (t === '') continue;
    if (
      PARTIES_HEADING.test(t) ||
      (BETWEEN_LINE.test(t) && partyItemFollows(paras, i)) ||
      MADE_LINE.test(t)
    ) {
      return numbered > 0 ? i : 0;
    }
    if ((p.label && !p.bullet && !p.isToc) || typedNumber(p.text)) numbered += 1;
  }
  return 0;
}

const SIGNATURE_CELL = /^\s*(?:signature|signed|print(?:ed)? name|name|title|position|date|by)\s*:?\s*$/i;

/** Is the next thing after `i` a signature table ("Signature", "Print Name", "Title", "Date" cells)? */
function signatureTableFollows(paras: readonly DocxPara[], i: number): boolean {
  let j = i + 1;
  while (j < paras.length && !paras[j]!.table && flat(paras[j]!.text) === '') j += 1;
  const first = paras[j];
  if (!first?.table) return false;
  const tableId = first.table.tableId;
  let hits = 0;
  for (let k = j; k < paras.length && k < j + 80; k += 1) {
    const q = paras[k]!;
    if (q.table?.tableId !== tableId) break;
    if (SIGNATURE_CELL.test(flat(q.text))) hits += 1;
  }
  return hits >= 3;
}

/** A bold capitals title across the full width of a table ("SERVICES" heading a form's section). */
function tableBanner(p: DocxPara, ft: string): boolean {
  return (
    p.table !== null &&
    p.table.rowCells === 1 &&
    p.table.parentTableId === null &&
    p.allBold &&
    isUpperTitle(ft) &&
    titleLike(ft, 8) &&
    !SIGNATURE_WORDS.test(ft) &&
    !EXECUTION_START.test(ft) &&
    !/\b(?:pte|ltd|limited|llp|inc|plc|corp|gmbh)\b/i.test(ft)
  );
}

function sameRowLaterCell(a: DocxPara, b: DocxPara): boolean {
  return (
    a.table !== null && b.table !== null && a.table.tableId === b.table.tableId && a.table.row === b.table.row && b.table.col > a.table.col
  );
}

/** Do signature lines ("…………", "______", "Name:", "Signature") follow `i` closely? */
function signatureLinesFollow(paras: readonly DocxPara[], i: number): boolean {
  let seen = 0;
  for (let j = i + 1; j < paras.length && seen < 3; j += 1) {
    const t = flat(paras[j]!.text);
    if (t === '') continue;
    seen += 1;
    if (/…{3,}|\.{5,}|_{5,}/.test(t) || SIGNATURE_WORDS.test(t)) return true;
  }
  return false;
}

function partyItemFollows(paras: readonly DocxPara[], i: number): boolean {
  for (let j = i + 1; j < paras.length && j < i + 4; j += 1) {
    const q = paras[j]!;
    if (flat(q.text) === '') continue;
    return q.label !== null || typedNumber(q.text) !== null;
  }
  return false;
}

export function analyseStructure(paras: readonly DocxPara[]): ContractStructure {
  const nodes: StructureNode[] = [];
  const owner: (string | null)[] = paras.map(() => null);
  const paraLabel: {
    label: string | null;
    source: LabelSource;
    party: string | null;
  }[] = paras.map(() => ({
    label: null,
    source: 'none',
    party: null,
  }));

  // Does the document number its clauses? If not, heading styles carry the hierarchy.
  let numberedCount = 0;
  for (const p of paras) {
    if (p.label && /\d/.test(p.label) && !p.bullet) numberedCount += 1;
    else if (!p.label && typedNumber(p.text) && /\d/.test(typedNumber(p.text)!.token)) numberedCount += 1;
  }
  const numberedDoc = numberedCount >= 4;

  const contentsEnd = contentsListEnd(paras);

  let stack: Frame[] = [];
  let phase: 'front' | 'parties' | 'recitals' | 'body' | 'execution' = 'front';
  let container = null as Frame | null;
  let part = null as Frame | null;
  let front = null as StructureNode | null;
  let parties = null as StructureNode | null;
  let pendingTitle = null as StructureNode | null;
  /** The last title-like line seen in the front matter: it may turn out to head the body. */
  let frontTitle: number | null = null;
  /** Inside a contents list on the cover: its entries are front matter, not clauses or headings. */
  let inContents = false;
  /** Tables whose rows are numbered, and the stack depth their row nodes sit at. */
  const rowTables = new Map<number, number>();

  const newNode = (
    kind: NodeKind,
    start: number,
    region: Region,
    fields: Partial<Pick<StructureNode, 'label' | 'heading' | 'term' | 'parentId'>> & {
      source?: LabelSource;
    } = {},
  ): StructureNode => {
    const node: StructureNode = {
      id: `n${nodes.length}`,
      kind,
      label: fields.label ?? null,
      fullLabel: null,
      heading: fields.heading ?? null,
      term: fields.term ?? null,
      parentId: fields.parentId ?? null,
      region,
      start,
      end: start,
      numberingSource: fields.source ?? 'none',
    };
    nodes.push(node);
    return node;
  };

  const popTo = (size: number): void => {
    if (stack.length > size) stack = stack.slice(0, Math.max(0, size));
  };
  const scopeBase = (): number => {
    if (part && stack.includes(part)) return stack.indexOf(part) + 1;
    if (container && stack.includes(container)) return stack.indexOf(container) + 1;
    return 0;
  };
  const scopeNode = (): Frame | null =>
    part && stack.includes(part) ? part : container && stack.includes(container) ? container : null;
  const top = (): Frame | null => stack[stack.length - 1] ?? null;
  const closeAll = (): void => {
    stack = [];
    container = null;
    part = null;
  };
  const region = (): Region => (container || (part && part.parentable) ? 'schedule' : 'body');
  const regionKind = (word: string): NodeKind => CONTAINER_WORDS[word.toLowerCase()] ?? 'schedule';

  const startContainer = (
    i: number,
    kind: NodeKind,
    label: string,
    heading: string | null,
    source: LabelSource,
  ): void => {
    closeAll();
    phase = 'body';
    const node = newNode(kind, i, 'schedule', { label, heading, source });
    container = frame(node);
    stack = [container];
    pendingTitle = heading ? null : node;
  };

  const startPart = (i: number, label: string, heading: string | null, source: LabelSource): void => {
    if (container) popTo(stack.indexOf(container) + 1);
    else closeAll();
    if (phase !== 'body') phase = 'body';
    const node = newNode('part', i, container ? 'schedule' : 'body', {
      label,
      heading,
      source,
      parentId: container?.node.id ?? null,
    });
    part = frame(node);
    stack.push(part);
    pendingTitle = heading ? null : node;
  };

  /** Start the body: the front matter, parties and recitals are closed. */
  const beginBody = (i: number): void => {
    if (phase === 'body') return;
    const wasFront = phase === 'front';
    popTo(0);
    phase = 'body';
    // a title line late in the front matter ("CONTRACT", "TERMS AND CONDITIONS") heads the body
    if (wasFront && front && frontTitle !== null && frontTitle > front.start) {
      const node = newNode('heading', frontTitle, 'body', {
        heading: flat(textOf(paras[frontTitle]!)),
      });
      for (let j = frontTitle; j < i; j += 1) if (owner[j] === front.id) owner[j] = node.id;
      stack = [frame(node, { parentable: false })];
    }
    frontTitle = null;
  };

  const findSibling = (shape: NumberShape, base: number, below: number): number => {
    for (let j = below - 1; j >= base; j -= 1) {
      const f = stack[j]!;
      if (f.shape && sameKind(shape, f.shape)) return j;
      if (f.node.kind === 'definition' || !f.parentable) return -1;
    }
    return -1;
  };

  /**
   * A definition met under an operative subclause of a definitions clause ("4.K On termination, the
   * Academy Trust must…" then "“Land” means…") belongs to the definitions clause, unless that
   * subclause itself introduces definitions ("In this clause 4.K:").
   */
  const definitionsHolder = (k: number, base: number): number => {
    const holder = stack[k];
    if (!holder || holder.node.kind !== 'subclause') return k;
    const text = paras[holder.node.start]?.text ?? '';
    if (/[:：]\s*$|\bmeans?\b|\bdefin|\bfollowing\s+(?:words|terms|expressions)\b|\bin\s+this\b/i.test(text)) return k;
    for (let a = k - 1; a >= base; a -= 1) if (stack[a]!.node.kind === 'definitions') return a;
    return k;
  };

  /** Does this numbered paragraph introduce the definitions that follow it, or merely precede them? */
  const introducesTable = (start: number): boolean => {
    const text = (typedNumber(paras[start]?.text ?? '')?.rest ?? paras[start]?.text ?? '').trim();
    if (/[:：]\s*$|\bmean(?:s|ing|ings)?\b|\bdefin/i.test(text)) return true;
    return text.split(/\s+/).length <= 6;
  };

  /** Where a numbered paragraph goes: pop siblings and deeper levels, return the parent frame. */
  const placeNumbered = (
    shape: NumberShape,
    listGroup: string | null,
    ilvl: number | null,
    cell: TableCoord | null = null,
  ): Frame | null => {
    const base = scopeBase();
    // inside a numbered table, a limb never climbs out of its row: a restarted "(i)" in a row
    // with no number of its own belongs to the last numbered row, not to the clause around the table
    const rowFloor = cell && cell.col > 0 ? rowTables.get(cell.tableId) : undefined;
    const floor = rowFloor !== undefined && stack[rowFloor]?.node.rowNumber ? rowFloor : -1;
    for (let k = stack.length - 1; k >= base; k -= 1) {
      const c = stack[k]!;
      if (k === floor) {
        popTo(k + 1);
        return c;
      }
      if (!c.parentable) continue;
      // a number typed in a later cell of a numbered table row belongs to that row ("6." then "1.")
      const rowCell = paras[c.node.start]?.table;
      if (cell && rowCell && cell.col > 0 && rowCell.col === 0 && rowCell.tableId === cell.tableId && rowCell.row === cell.row) {
        popTo(k + 1);
        return c;
      }
      if (c.node.kind === 'definition') {
        if (shape.cls !== 'decimal' && shape.cls !== 'dotted') {
          popTo(k + 1);
          return c;
        }
        continue;
      }
      if (c.shape === null) {
        popTo(k + 1);
        return c;
      }
      if (listGroup !== null && c.listGroup === listGroup && c.ilvl !== null && ilvl !== null) {
        if (c.ilvl < ilvl) {
          popTo(k + 1);
          return c;
        }
        continue;
      }
      if (shape.cls === 'dotted') {
        if (c.shape.cls === 'dotted' || c.shape.cls === 'decimal') {
          const cParts = c.shape.cls === 'decimal' ? [c.shape.core] : c.shape.parts;
          if (cParts.length < shape.parts.length && cParts.every((x, j) => x === shape.parts[j])) {
            popTo(k + 1);
            return c;
          }
        }
        continue;
      }
      if (sameKind(shape, c.shape)) continue;
      // a plain Word "2" never sits under a lettered limb of another list (a definition's "(c)")
      // and a typed "2." never sits under a bracketed "(b)"
      if (
        ((listGroup !== null && c.listGroup !== null && c.listGroup !== listGroup) ||
          (listGroup === null && shape.shape === '#.' && c.shape.shape.startsWith('('))) &&
        shape.cls === 'decimal' &&
        !shape.shape.startsWith('(') &&
        c.shape.cls !== 'decimal' &&
        c.shape.cls !== 'dotted'
      ) {
        continue;
      }
      if (
        shape.cls === 'decimal' &&
        (c.shape.cls === 'decimal' || c.shape.cls === 'dotted') &&
        shape.shape === '#.'
      )
        continue;
      const sib = findSibling(shape, base, k);
      if (sib >= 0) {
        k = sib;
        continue;
      }
      popTo(k + 1);
      return c;
    }
    popTo(base);
    return scopeNode();
  };

  const owns = (i: number, id: string | null): void => {
    owner[i] = id;
  };

  for (let i = 0; i < paras.length; i += 1) {
    const p = paras[i]!;
    const text = textOf(p);
    const ft = flat(text);
    const wordLabel = p.inAccepted ? p.label : p.labelOriginal;
    const blank = ft === '';
    const current = (): string | null => top()?.node.id ?? front?.id ?? null;

    if (blank && !wordLabel) {
      owns(i, current());
      continue;
    }
    if (p.isToc) {
      if (!front && phase === 'front' && nodes.length === 0) {
        front = newNode('front_matter', i, 'front');
        stack = [frame(front)];
      }
      owns(i, current());
      continue;
    }

    if (phase === 'front' && !wordLabel && !p.table && /^(?:table of )?contents:?$/i.test(ft)) {
      if (!front) {
        front = newNode('front_matter', i, 'front');
        stack = [frame(front)];
      }
      inContents = true;
      owns(i, front.id);
      continue;
    }
    if (inContents) {
      // an entry ends in its page number ("5. The buyer's obligations 6")
      if (phase === 'front' && (p.isToc || /[\t ]\d{1,4}\s*$/.test(text))) {
        owns(i, current());
        continue;
      }
      inContents = false;
    }
    if (i < contentsEnd && phase === 'front') {
      // the cover before the parties block: typed numbers keep their label but cite nothing
      if (!front) {
        front = newNode('front_matter', i, 'front');
        stack = [frame(front)];
      }
      if (wordLabel) {
        if (p.inAccepted) paraLabel[i] = { label: wordLabel, source: 'word', party: null };
      } else {
        const typedCover = !p.bullet && !p.table ? typedNumber(text) : null;
        // a contents line with its page number ("2.1 Delivery\t3") is the contents, and carries no number
        if (typedCover && !/\t\s*\d{1,4}\s*$/.test(text))
          paraLabel[i] = {
            label: typedCover.token,
            source: 'typed',
            party: null,
          };
      }
      owns(i, front.id);
      continue;
    }
    // -- a title line taken by the schedule/part that just opened
    if (
      pendingTitle &&
      !wordLabel &&
      (!p.table || tableBanner(p, ft) || sameRowLaterCell(paras[pendingTitle.start]!, p)) &&
      titleLike(ft, 16) &&
      !CONTAINER_RE.test(ft) &&
      !PART_RE.test(ft)
    ) {
      pendingTitle.heading = ft;
      pendingTitle = null;
      owns(i, current());
      continue;
    }
    pendingTitle = null;

    // -- the execution block ends at the next schedule, part or heading of the agreement
    if (phase === 'execution') {
      const endsHere =
        tableBanner(p, ft) ||
        (!p.table &&
          ((wordLabel !== null &&
            !p.bullet &&
            (p.isHeadingStyle || /heading|part|title/i.test(p.styleName ?? '') || p.listLevel === 0) &&
            ft !== '' &&
            !signatureLinesFollow(paras, i)) || // "I. Official and Mailing Address" then "Name:" lines
            (wordLabel === null &&
              titleLike(ft, 10) &&
              ((p.isHeadingStyle && !/signat/i.test(p.styleName ?? '')) || isUpperTitle(ft)) &&
              !/^[[(]/.test(ft) &&
              !SIGNATURE_WORDS.test(ft) &&
              !EXECUTION_START.test(ft) &&
              ft.split(' ').length >= 2 &&
              !signatureLinesFollow(paras, i)) ||
            CONTAINER_RE.test(ft)));
      // the cover's title repeated after an order form's signature block opens a second cover
      const coverAgain =
        front !== null && !p.table && ft === flat(textOf(paras[front.start]!)) && titleLike(ft, 10);
      if (coverAgain) {
        closeAll();
        phase = 'front';
        front = newNode('front_matter', i, 'front');
        stack = [frame(front)];
        owns(i, front.id);
        continue;
      }
      if (!endsHere) {
        owns(i, current());
        continue;
      }
      closeAll();
      phase = 'body';
    }

    // -- execution block
    if (
      !wordLabel &&
      ((phase !== 'front' && EXECUTION_START.test(ft) && (!container || EXECUTION_STRONG.test(ft))) ||
        // a schedule may hold a form with its own signature table; an annex listed on a cover-page table may not
        (!p.table && (!container || paras[container.node.start]!.table !== null) && signatureTableFollows(paras, i)))
    ) {
      closeAll();
      phase = 'execution';
      const node = newNode('execution', i, 'execution');
      stack = [frame(node)];
      owns(i, node.id);
      continue;
    }

    // -- schedule / annex / appendix title (painted by Word, or typed)
    const paintedContainer = wordLabel ? CONTAINER_RE.exec(wordLabel) : null;
    if (paintedContainer && !p.table) {
      if (!(container && container.node.label === wordLabel)) {
        startContainer(i, regionKind(paintedContainer[1]!), wordLabel!, ft || null, 'word');
        paraLabel[i] = { label: wordLabel!, source: 'word', party: null };
        owns(i, container!.node.id);
        continue;
      }
    }
    // a part-styled Word number whose text names a schedule or annex ("V. Annex – Processing")
    if (
      wordLabel &&
      !p.table &&
      /part|schedule|annex|appendix/i.test(p.styleName ?? '') &&
      CONTAINER_WORD_START.test(ft) &&
      !/^\s*part\b/i.test(wordLabel)
    ) {
      const word = CONTAINER_WORD_START.exec(ft)![1]!;
      startContainer(i, regionKind(word), wordLabel, ft, 'word');
      paraLabel[i] = { label: wordLabel, source: 'word', party: null };
      owns(i, container!.node.id);
      continue;
    }
    // an optional schedule or part in square brackets: "[SCHEDULE 7]", "[Part 3 - Policies]"
    const fb = /^\[[^\]]+\]$/.test(ft) ? ft.slice(1, -1).trim() : ft;
    // typed, or typed into the first cell of a cover-page table ("Annex I(A) List of Parties")
    const typedContainer =
      !wordLabel && (!p.table || (p.table.col === 0 && p.table.parentTableId === null)) ? CONTAINER_RE.exec(fb) : null;
    if (typedContainer && titleEvidence(p, ft)) {
      const rest = typedContainer[3]!.trim();
      const head = fb.slice(0, fb.length - typedContainer[3]!.length);
      const sep = /[-–—:.]\s*$/.test(head) || /\t/.test(text);
      if (rest === '' || sep || titleLike(rest, 14)) {
        const label = head.replace(/[\s\-–—:.]+$/, '').trim();
        if (container && flat(container.node.label ?? '').toLowerCase() === label.toLowerCase()) {
          owns(i, current()); // a running title repeated at the top of a page
          continue;
        }
        startContainer(i, regionKind(typedContainer[1]!), label, rest === '' ? null : rest, 'typed');
        paraLabel[i] = { label, source: 'typed', party: null };
        owns(i, container!.node.id);
        continue;
      }
    }
    const unnumberedContainer = !wordLabel && !p.table ? CONTAINER_NOID_RE.exec(ft) : null;
    if (unnumberedContainer && titleEvidence(p, ft) && phase !== 'front') {
      startContainer(
        i,
        regionKind(unnumberedContainer[1]!),
        '',
        unnumberedContainer[2]?.trim() || null,
        'none',
      );
      container!.node.label = null;
      owns(i, container!.node.id);
      continue;
    }
    // inside a schedule, a title in the schedule title's own style opens the next schedule
    if (
      container &&
      !wordLabel &&
      !p.table &&
      container.node.start !== i &&
      p.styleId !== null &&
      p.styleId === paras[container.node.start]!.styleId &&
      (p.isHeadingStyle || /schedule|annex|appendix/i.test(p.styleName ?? '')) &&
      titleLike(ft, 10)
    ) {
      startContainer(i, container.node.kind, '', ft, 'none');
      container!.node.label = null;
      owns(i, container!.node.id);
      continue;
    }
    // "Section C:" painted on a title: a division heading, not a clause
    if (
      wordLabel &&
      !p.table &&
      /^(?:section|chapter|division)\s+[A-Z][:.]?$/i.test(wordLabel) &&
      titleLike(ft, 14)
    ) {
      popTo(scopeBase());
      const node = newNode('heading', i, region(), {
        heading: ft,
        parentId: scopeNode()?.node.id ?? null,
      });
      stack.push(frame(node, { parentable: false }));
      paraLabel[i] = { label: wordLabel, source: 'word', party: null };
      owns(i, node.id);
      continue;
    }
    // a form inside a schedule divided into "Section 1 – Details of ...", "Section 2 – ..."
    const formSection = container && !wordLabel ? /^section\s+(\d{1,2})\s*[-–—:]\s*\S/i.exec(ft) : null;
    if (formSection && titleLike(ft, 14) && (!p.table || p.table.col === 0)) {
      popTo(scopeBase());
      const label = `Section ${formSection[1]}`;
      const node = newNode('clause', i, region(), {
        label,
        heading: ft.replace(/^section\s+\d{1,2}\s*[-–—:]\s*/i, ''),
        source: 'typed',
        parentId: scopeNode()?.node.id ?? null,
      });
      stack.push(frame(node, { shape: numberShape(label) }));
      paraLabel[i] = { label, source: 'typed', party: null };
      owns(i, node.id);
      continue;
    }
    // a schedule divided by a lettered Word list of titles ("A. Conditions Precedent", "B. ...")
    if (
      container &&
      wordLabel &&
      !p.table &&
      p.listLevel === 0 &&
      /^[A-Z]\.$/.test(wordLabel) &&
      titleLike(ft, 10) &&
      !/[.:]$/.test(ft) &&
      (part === null || /^[A-Z]\.$/.test(part.node.label ?? '') || stack.indexOf(part) < 0)
    ) {
      startPart(i, wordLabel, ft, 'word');
      paraLabel[i] = { label: wordLabel, source: 'word', party: null };
      owns(i, part!.node.id);
      continue;
    }
    const paintedPart = wordLabel ? PART_RE.exec(wordLabel) : null;
    if (paintedPart && !p.table) {
      startPart(i, wordLabel!, ft || null, 'word');
      paraLabel[i] = { label: wordLabel!, source: 'word', party: null };
      owns(i, part!.node.id);
      continue;
    }
    // a top-level Word number in a part style ("I. Index", "IV. Short form Terms"): a division of the
    // document that heads what follows without being the parent of its clauses
    if (wordLabel && !p.table && !container && /\bpart\b/i.test(p.styleName ?? '') && ft !== '') {
      closeAll();
      const kindRegion: Region =
        phase === 'front' || phase === 'parties' || phase === 'recitals' ? 'front' : 'body';
      if (front && stack.length === 0 && kindRegion === 'front') {
        /* front matter closes */
      }
      const node = newNode('part', i, kindRegion, {
        label: wordLabel,
        heading: ft,
        source: 'word',
      });
      stack = [frame(node, { parentable: false })];
      paraLabel[i] = { label: wordLabel, source: 'word', party: null };
      owns(i, node.id);
      continue;
    }
    // an optional part opens with a bracket that closes pages later ("[Part 3 - the Company’s Policies")
    const fp = /^\[\s*part\s/i.test(fb) && !fb.includes(']') ? fb.slice(1).trim() : fb;
    const typedPart = !wordLabel && !p.table ? PART_RE.exec(fp) : null;
    if (typedPart && container && (titleEvidence(p, ft) || (titleLike(fp, 14) && /^\s*part\s+\S+\s*[:\-–—]/i.test(fp)))) {
      const rest = typedPart[2]!.trim();
      if (rest === '' || titleLike(rest, 14) || /[-–—:]/.test(fp.slice(0, fp.length - rest.length))) {
        const label = fp
          .slice(0, fp.length - typedPart[2]!.length)
          .replace(/[\s\-–—:.]+$/, '')
          .trim();
        startPart(i, label, rest === '' ? null : rest, 'typed');
        paraLabel[i] = { label, source: 'typed', party: null };
        owns(i, part!.node.id);
        continue;
      }
    }

    // -- front matter markers
    if (phase === 'front' || phase === 'parties' || phase === 'recitals') {
      if (
        !wordLabel &&
        (PARTIES_HEADING.test(ft) ||
          (BETWEEN_LINE.test(ft) && (partyItemFollows(paras, i) || (phase === 'front' && !parties && front !== null))) ||
          (phase === 'front' && !parties && !p.table && PARTIES_INTRO.test(ft)))
      ) {
        phase = 'parties';
        parties = newNode('parties', i, 'front', {
          heading: PARTIES_INTRO.test(ft) && !PARTIES_HEADING.test(ft) ? null : ft.replace(/\s*:$/, ''),
        });
        stack = [frame(parties)];
        owns(i, parties.id);
        continue;
      }
      if (!wordLabel && RECITALS_HEADING.test(ft)) {
        popTo(0);
        phase = 'recitals';
        const node = newNode('heading', i, 'front', { heading: ft });
        stack = [frame(node, { parentable: false })];
        owns(i, node.id);
        continue;
      }
      if (!wordLabel && OPERATIVE_HEADING.test(ft)) {
        popTo(0);
        phase = 'body';
        frontTitle = null;
        const node = newNode('heading', i, 'front', {
          heading: ft.replace(/\s*:$/, ''),
        });
        owns(i, node.id);
        continue;
      }
    }

    // -- numbers
    let typed = !wordLabel && !p.bullet ? typedNumber(text) : null;
    // a reference typed alone in a cell ("1.I", "2. P.1", "3.A – 3.F" in a checklist) is not a number
    if (typed && p.table && /^(?:$|[–—-]\s*\d|[A-Z]{1,2}\.\d+$)/.test(flat(typed.rest))) typed = null;
    const label = wordLabel ?? typed?.token ?? null;
    const source: LabelSource = wordLabel ? 'word' : typed ? 'typed' : 'none';
    const rest = typed ? flat(typed.rest) : ft;

    if (label !== null) {
      const shape = numberShape(label, p.numFmt);

      if (phase === 'front' || phase === 'parties' || phase === 'recitals') {
        const upper = shape.cls === 'upperLetter';
        const numeric = shape.cls === 'decimal' && shape.shape !== '#.#';
        // a numbered capitals title ("1. BACKGROUND", "1. DEFINITIONS") is the first clause, not a party or recital
        const clauseTitle =
          isUpperTitle(rest) && titleLike(rest, 8) && !/\b(?:pte|ltd|limited|llp|inc|plc|corp)\b/i.test(rest);
        if (!clauseTitle && (phase === 'recitals' || (upper && phase !== 'parties' && shape.shape.startsWith('(')))) {
          if (upper || shape.cls === 'decimal') {
            popTo(0);
            phase = 'recitals';
            const node = newNode('recital', i, 'front', { label, source });
            stack = [frame(node, { listGroup: p.listGroup, ilvl: p.listLevel, shape })];
            paraLabel[i] = {
              label: p.inAccepted ? label : null,
              source,
              party: null,
            };
            owns(i, node.id);
            continue;
          }
        }
        if (!clauseTitle && (phase === 'parties' || (numeric && shape.shape.startsWith('(') && phase === 'front'))) {
          if (!parties) {
            parties = newNode('parties', i, 'front');
            stack = [frame(parties)];
          }
          phase = 'parties';
          paraLabel[i] = {
            label: p.inAccepted ? label : null,
            source,
            party: `Party ${stripLabelPunct(label)}`,
          };
          owns(i, parties.id);
          continue;
        }
        beginBody(i);
      }

      if (blank) {
        // a number painted in the first cell of a table row numbers the row; elsewhere, nothing to cite
        if (!(p.table && p.table.col === 0 && wordLabel)) {
          paraLabel[i] = {
            label: p.inAccepted ? label : null,
            source,
            party: null,
          };
          owns(i, current());
          continue;
        }
      }

      // the numbered rows of a table sit together under whatever was open when the table began
      const rowNumber = blank && p.table !== null && p.table.col === 0 && wordLabel !== null;
      let parent: Frame | null;
      const rowDepth = rowNumber ? rowTables.get(p.table!.tableId) : undefined;
      if (rowNumber && rowDepth !== undefined && rowDepth <= stack.length) {
        popTo(rowDepth);
        parent = top();
      } else if (rowNumber) {
        let k = stack.length - 1;
        while (k >= 0 && !stack[k]!.parentable) k -= 1;
        popTo(k + 1);
        parent = k >= 0 ? stack[k]! : null;
        rowTables.set(p.table!.tableId, k + 1);
      } else {
        parent = placeNumbered(shape, p.listGroup, p.listLevel, p.table);
      }
      const term = definitionTerm(rest);
      const kind: NodeKind =
        parent === null || parent === container || parent === part ? 'clause' : 'subclause';
      const created = newNode(kind, i, region(), {
        label,
        heading: numberedHeading(p, rest, typed !== null),
        term,
        source,
        parentId: parent?.node.id ?? null,
      });
      const node: StructureNode = rowNumber ? Object.assign(created, { rowNumber: true }) : created;
      stack.push(frame(node, { listGroup: p.listGroup, ilvl: p.listLevel, shape }));
      paraLabel[i] = {
        label: p.inAccepted ? label : null,
        source,
        party: null,
      };
      owns(i, node.id);
      continue;
    }

    // -- unnumbered paragraphs
    const runIn = !p.table ? runInTitle(p) : null;
    if (phase === 'front' && front && p.isHeadingStyle && !p.table && /:$/.test(ft) && titleLike(ft.slice(0, -1), 10)) {
      // "Information about the Academy:" heads a section of the cover, not the agreement
      const node = newNode('heading', i, 'front', { heading: ft.replace(/\s*:$/, '') });
      stack = [frame(node, { parentable: false })];
      frontTitle = null;
      owns(i, node.id);
      continue;
    }
    if (phase === 'front' && (p.isHeadingStyle || runIn !== null) && !p.isTitleStyle && !p.table && front) {
      beginBody(i);
    }
    // an unnumbered recital: a sentence straight after the recitals heading, or after another unnumbered one
    if (
      phase === 'recitals' &&
      !p.table &&
      !titleLike(ft, 10) &&
      (top()?.node.kind === 'heading' || (top()?.node.kind === 'recital' && top()?.node.label === null))
    ) {
      const node = newNode('recital', i, 'front');
      stack = [frame(node)];
      owns(i, node.id);
      continue;
    }
    if (phase === 'parties' || phase === 'recitals' || phase === 'front') {
      if (!front && phase === 'front') {
        front = newNode('front_matter', i, 'front');
        stack = [frame(front)];
      } else if (
        phase === 'front' &&
        front &&
        !p.table &&
        titleLike(ft, 8) &&
        titleEvidence(p, ft) &&
        !p.isTitleStyle &&
        !/\b(?:19|20)\d{2}\b|\bv(?:ersion)?\s*\d/i.test(ft)
      ) {
        frontTitle = i;
      }
      owns(i, current());
      continue;
    }

    // definitions
    const term = !p.table || p.table.col === 0 ? (definitionTerm(ft) ?? boldTerm(p, ft)) : null;
    if (term) {
      const base = scopeBase();
      let k = stack.length - 1;
      while (
        k >= base &&
        (stack[k]!.node.kind === 'definition' || isUnderDefinition(stack, k, base) || !stack[k]!.parentable)
      )
        k -= 1;
      k = definitionsHolder(k, base);
      const lead = p.table && k >= base ? stack[k]! : null;
      if (lead && (lead.node.kind === 'clause' || lead.node.kind === 'subclause') && !introducesTable(lead.node.start))
        k -= 1;
      popTo(k + 1);
      const holder = k >= base ? stack[k]! : null;
      if (holder && (holder.node.kind === 'clause' || holder.node.kind === 'subclause'))
        holder.node.kind = 'definitions';
      const node = newNode('definition', i, region(), {
        term,
        parentId: holder?.node.id ?? scopeNode()?.node.id ?? null,
      });
      stack.push(frame(node));
      owns(i, node.id);
      continue;
    }
    if (
      p.table &&
      p.table.col === 0 &&
      p.table.rowCells >= 2 &&
      p.table.parentTableId === null &&
      tableDefinition(paras, i, stack)
    ) {
      const base = scopeBase();
      let k = stack.length - 1;
      while (k >= base && (stack[k]!.node.kind === 'definition' || isUnderDefinition(stack, k, base))) k -= 1;
      // a definitions table after an unrelated numbered sentence ("3. The headings … will not affect
      // its interpretation") belongs to what holds that sentence, not to the sentence
      const lead = k >= base ? stack[k]! : null;
      if (lead && (lead.node.kind === 'clause' || lead.node.kind === 'subclause') && !introducesTable(lead.node.start))
        k -= 1;
      popTo(k + 1);
      const holder = k >= base ? stack[k]! : null;
      if (holder && (holder.node.kind === 'clause' || holder.node.kind === 'subclause'))
        holder.node.kind = 'definitions';
      const node = newNode('definition', i, region(), {
        term: ft.replace(/^[“"‘']+|[”"’']+$/g, ''),
        parentId: holder?.node.id ?? scopeNode()?.node.id ?? null,
      });
      stack.push(frame(node));
      owns(i, node.id);
      continue;
    }

    // headings without numbers
    if (!numberedDoc && p.isHeadingStyle && !p.table && titleLike(ft, 16)) {
      const rank = p.outlineLevel ?? 0;
      const base = scopeBase();
      let k = stack.length - 1;
      while (
        k >= base &&
        !(stack[k]!.headingRank !== null && stack[k]!.headingRank! < rank && stack[k]!.shape === null)
      )
        k -= 1;
      popTo(k + 1);
      const parent = k >= base ? stack[k]! : scopeNode();
      const kind: NodeKind =
        parent === null || parent === container || parent === part ? 'clause' : 'subclause';
      const node = newNode(kind, i, region(), {
        heading: ft,
        parentId: parent?.node.id ?? null,
      });
      stack.push(frame(node, { headingRank: rank }));
      owns(i, node.id);
      continue;
    }
    if (
      numberedDoc &&
      ((!p.table &&
        titleLike(ft, 10) &&
        !/:$/.test(ft) &&
        !/^[^:]{1,40}:\s*\S/.test(ft) && // "UEN: 201534417G" is a field and its value
        (p.isHeadingStyle || p.allBold || isUpperTitle(ft))) ||
        (!container && tableBanner(p, ft)))
    ) {
      // a title in a numbered document heads what follows; it is not the parent of the clauses after it.
      // Between the sub-clauses of an open top-level clause it is a heading within that clause.
      const base = scopeBase();
      const outer = stack[base];
      const within =
        outer && outer.shape && outer.parentable && !nextClosesClause(paras, i, outer) ? outer : null;
      const scope = within ?? scopeNode();
      popTo(within ? base + 1 : base);
      const node = newNode('heading', i, region(), {
        heading: ft,
        parentId: scope?.node.id ?? null,
      });
      stack.push(frame(node, { parentable: false }));
      owns(i, node.id);
      continue;
    }
    if (runIn !== null) {
      const base = scopeBase();
      let k = stack.length - 1;
      while (
        k >= base &&
        !(stack[k]!.headingRank !== null && stack[k]!.headingRank! < 9 && stack[k]!.shape === null)
      )
        k -= 1;
      popTo(k + 1);
      const parent = k >= base ? stack[k]! : scopeNode();
      const kind: NodeKind =
        parent === null || parent === container || parent === part ? 'clause' : 'subclause';
      const node = newNode(kind, i, region(), {
        heading: runIn,
        parentId: parent?.node.id ?? null,
      });
      stack.push(frame(node, { headingRank: 9 }));
      owns(i, node.id);
      continue;
    }

    owns(i, current());
  }

  fixSpans(nodes, owner, paras.length);
  for (const n of nodes) n.fullLabel = fullLabelFor(n, nodes);

  const byId = new Map(nodes.map((n) => [n.id, n]));
  const paragraphs: ParagraphStructure[] = paras.map((p, i) => {
    const l = paraLabel[i]!;
    const nodeId = owner[i] ?? null;
    const n = nodeId ? byId.get(nodeId) : undefined;
    let fullLabel: string | null = null;
    if (l.party) fullLabel = p.inAccepted ? l.party : null;
    else if (l.label && n && n.start === i) fullLabel = n.fullLabel;
    return {
      label: l.label,
      fullLabel,
      labelSource: l.label ? l.source : 'none',
      nodeId,
    };
  });
  return { nodes, paragraphs };
}

function sameKind(a: NumberShape, b: NumberShape): boolean {
  if (a.cls === 'dotted' || b.cls === 'dotted') return a.cls === b.cls && a.parts.length === b.parts.length;
  if (a.cls !== b.cls) {
    // "(i)" after "(h)" is the letter i; after nothing, the numeral (unless Word says which)
    if (a.certain) return false;
    const letterRoman =
      (a.cls === 'lowerRoman' && b.cls === 'lowerLetter' && nextLetter(b.core) === a.core) ||
      (a.cls === 'upperRoman' && b.cls === 'upperLetter' && nextLetter(b.core) === a.core);
    if (!letterRoman) return false;
  }
  return a.shape === b.shape;
}

/** Does the next numbered paragraph after `i` close the top-level clause `outer` (a sibling of it)? */
function nextClosesClause(paras: readonly DocxPara[], i: number, outer: Frame): boolean {
  for (let j = i + 1; j < paras.length; j += 1) {
    const q = paras[j]!;
    const t = flat(textOf(q));
    if (t === '') continue;
    const label = q.inAccepted ? q.label : q.labelOriginal;
    const typed = !label && !q.bullet ? typedNumber(q.text) : null;
    if (!label && !typed) {
      if (CONTAINER_RE.test(t) || EXECUTION_START.test(t)) return true;
      continue;
    }
    if (label && (CONTAINER_RE.test(label) || PART_RE.test(label))) return true;
    if (
      label &&
      q.listGroup !== null &&
      q.listGroup === outer.listGroup &&
      q.listLevel !== null &&
      outer.ilvl !== null
    ) {
      return q.listLevel <= outer.ilvl;
    }
    const shape = numberShape(label ?? typed!.token, q.numFmt);
    if (shape.cls === 'dotted') return false;
    return sameKind(shape, outer.shape!);
  }
  return true;
}

function isUnderDefinition(stack: readonly Frame[], k: number, base: number): boolean {
  for (let j = k; j >= base; j -= 1) {
    if (stack[j]!.node.kind === 'definition') return true;
    if (stack[j]!.shape && (stack[j]!.shape!.cls === 'decimal' || stack[j]!.shape!.cls === 'dotted'))
      return false;
  }
  return false;
}

const DEF_HEADER = /^(term|meaning|definition|defined term|expression|word|interpretation)s?$/i;

/** Is `i` in a table whose header row reads "Term | Meaning"? */
function definitionTableHeader(paras: readonly DocxPara[], i: number): boolean {
  const cell = paras[i]!.table!;
  for (let j = i - 1; j >= 0 && j > i - 400; j -= 1) {
    const q = paras[j]!;
    if (q.table?.tableId !== cell.tableId) return false;
    if (q.table.row === 0 && q.table.col === 0) {
      const next = paras[j + 1];
      return DEF_HEADER.test(flat(q.text)) && next?.table?.row === 0 && DEF_HEADER.test(flat(next.text));
    }
  }
  return false;
}

function tableDefinition(paras: readonly DocxPara[], i: number, stack: readonly Frame[]): boolean {
  const p = paras[i]!;
  const t = flat(p.text);
  if (t === '' || t.length > 80 || /[.;]$/.test(t)) return false;
  if (p.table!.row === 0 && DEF_HEADER.test(t)) return false;
  if (definitionTableHeader(paras, i)) return true;
  const next = paras[i + 1];
  if (!next || !next.table || next.table.tableId !== p.table!.tableId || next.table.row !== p.table!.row)
    return false;
  // a quoted term alone in the first cell ("“Beneficiary”") is a definition row
  if (/^[“"‘'][^“”"]{1,78}[”"’']$/.test(t)) return true;
  for (let j = stack.length - 1; j >= 0; j -= 1) {
    const n = stack[j]!.node;
    if (n.kind === 'definitions' || n.kind === 'definition') return true;
    // limbs of an earlier row's definition in this same table
    if (paras[n.start]?.table?.tableId === p.table!.tableId && n.kind === 'subclause') continue;
    const start = paras[n.start];
    if (
      start &&
      /\b(meaning|definitions?|defined|terms?)\b/i.test(start.text) &&
      /:\s*$/.test(start.text.trim())
    )
      return true;
    if (
      start &&
      /\bfollowing (?:additional )?terms\b|\bhave the (?:following )?meanings?\b/i.test(start.text)
    )
      return true;
    break;
  }
  return false;
}

/** The title of a numbered paragraph: the whole line when it reads as one, else a bold run-in. */
function numberedHeading(p: DocxPara, rest: string, typed: boolean): string | null {
  const t = flat(rest);
  if (t === '') return null;
  if (titleLike(t, 12) && !definitionTerm(t)) return t.replace(/[.:]$/, '');
  return runInTitle(p, typed ? rest : null);
}

/** A bold or underlined lead ending in a full stop: "Entire agreement. This Agreement ...". */
function runInTitle(p: DocxPara, restOverride: string | null = null): string | null {
  const lead = flat(p.boldLead || p.underlineLead);
  if (lead === '') return null;
  const whole = flat(restOverride ?? p.text);
  let title = lead;
  const typed = typedNumber(title);
  if (typed) title = flat(typed.rest);
  if (!/\.$/.test(title)) {
    const after = whole.slice(whole.indexOf(title) + title.length);
    if (!/^\s*\.\s/.test(after)) return null;
  }
  title = title.replace(/[\s.]+$/, '');
  if (title === '' || title.length > RUNIN_MAX || title.split(' ').length > 10) return null;
  if (!/^[A-Z“"']/.test(title)) return null;
  if (/^(?:note|important|warning|n\.b|nb)$/i.test(title)) return null;
  // a party's name in capitals ("BRIGHTWATER HOLDINGS PTE. LTD.") is not a title
  if (isUpperTitle(title) || /\b(?:pte|ltd|limited|llp|inc|plc|corp|gmbh|bv|sa)$/i.test(title)) return null;
  if (whole.length <= title.length + 2) return null; // the whole line is bold: a heading, not run-in
  return title;
}

/** A node runs to the paragraph before the next node at the same or a higher level; owners stay inside. */
function fixSpans(nodes: StructureNode[], owner: (string | null)[], n: number): void {
  const byId = new Map(nodes.map((x) => [x.id, x]));
  const isDescendant = (x: StructureNode, of: StructureNode): boolean => {
    for (let p = x.parentId; p; p = byId.get(p)?.parentId ?? null) if (p === of.id) return true;
    return false;
  };
  const sorted = [...nodes].sort((a, b) => a.start - b.start);
  for (let k = 0; k < sorted.length; k += 1) {
    const x = sorted[k]!;
    let end = n - 1;
    for (let j = k + 1; j < sorted.length; j += 1) {
      if (!isDescendant(sorted[j]!, x)) {
        end = sorted[j]!.start - 1;
        break;
      }
    }
    x.end = Math.max(x.start, end);
  }
  owner.forEach((id, i) => {
    if (!id) return;
    const x = byId.get(id);
    if (x && i > x.end) x.end = i;
  });
  for (const x of nodes) {
    for (
      let p = x.parentId ? byId.get(x.parentId) : undefined;
      p;
      p = p.parentId ? byId.get(p.parentId) : undefined
    ) {
      if (p.end < x.end) p.end = x.end;
    }
  }
}

const CONTAINERS = new Set<NodeKind>(['schedule', 'annex', 'appendix']);

export function fullLabelFor(node: StructureNode, nodes: readonly StructureNode[]): string | null {
  const byId = new Map(nodes.map((x) => [x.id, x]));
  const chain: StructureNode[] = [];
  for (let x: StructureNode | undefined = node; x; x = x.parentId ? byId.get(x.parentId) : undefined)
    chain.unshift(x);
  const label = node.label;
  switch (node.kind) {
    case 'parties':
      return label ? `Party ${stripLabelPunct(label)}` : null;
    case 'front_matter':
    case 'heading':
    case 'execution':
    case 'other':
      return null;
    case 'definition':
      if (!label) return null;
      break;
    default:
      break;
  }
  if (CONTAINERS.has(node.kind)) return tidyLabel(label) || null;
  if (node.kind === 'part') {
    const t = tidyLabel(label);
    if (!t) return null;
    const holder = [...chain.slice(0, -1)].reverse().find((a) => CONTAINERS.has(a.kind));
    const holderFull = holder ? (holder.fullLabel ?? fullLabelFor(holder, nodes)) : null;
    return holderFull ? `${holderFull} ${t}` : t;
  }
  if (node.kind === 'recital') {
    const t = stripLabelPunct(label ?? '');
    return t ? `Recital ${t}` : null;
  }
  if (!label) return null;
  let prefix: string | null = null;
  let start = 0;
  for (let k = chain.length - 2; k >= 0; k -= 1) {
    const a = chain[k]!;
    if (a.kind === 'part' || CONTAINERS.has(a.kind)) {
      prefix = a.fullLabel ?? fullLabelFor(a, nodes);
      start = k + 1;
      break;
    }
  }
  // a table's row number starts the citation again ("Schedule 1 para 6", not "... para 4(6)")
  for (let k = chain.length - 1; k >= start; k -= 1) {
    if (chain[k]!.rowNumber) {
      start = k;
      break;
    }
  }
  const composed = compose(chain.slice(start).map((x) => x.label));
  if (composed === null) return null;
  return prefix ? `${prefix} para ${composed}` : composed;
}

