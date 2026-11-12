/**
 * Pairing sketches (ADR-0037; spec §6.2) — the only thing pairing ever reads.
 *
 * A sketch is a small, plain-data summary of one document. Reducing every
 * document to one of these before scoring is what keeps M×N comparisons cheap
 * at bundle scale — but the more important property is that a sketch contains
 * nothing a human could not check. Every number the pairing grid shows is
 * derived from these three sets, which is why a proposed pair can be explained
 * rather than merely asserted (ADR-0008: no opaque judgement about documents).
 *
 * All three signals come from work the product already does:
 *
 *  - **paragraph hashes** are computed during `preprocessSide` regardless, so
 *    the dominant signal is free;
 *  - **defined terms** come from `analyseDefinedTerms`, so a sketch recognises
 *    SG/UK definition shapes exactly as the Defined Terms checker does rather
 *    than re-deriving them with a weaker regex;
 *  - **party names** come from the party-role harvester, for the same reason.
 *
 * Host-free: the `.docx` is parsed with the OOXML service and read through
 * `host-web/ooxml-loader`, neither of which touches Office JS.
 */
import { PizZipBundle } from '@shared/ooxml/bundle.js';
import type { CompareConfig, FlattenPolicy, JurisdictionProfile } from '@shared/config/index.js';
import { loadDocSnapshotFromOoxml } from '@host-web/ooxml-loader.js';
// Import the pure `analyse.js` submodules directly, not the feature barrels
// (`index.js`) — each barrel also re-exports that feature's store-touching
// `clearThenEmit`, which value-imports `@shared/findings`'s store module and
// the module-level `signal()` call it makes (`@preact/signals`, untree-
// shakeable). Pairing sketches run host-free at bundle scale and must never
// carry the add-in's UI state layer into their module graph.
import { analyseDefinedTerms } from '@features/defined-terms/analyse.js';
import { analysePartyRole } from '@features/party-role/analyse.js';
import { preprocessSide, buildNumberingContext } from '../../preprocess/index.js';
import { extractParts } from '../../io/parts.js';
import { contentTokens } from '../../diff/replace-threshold.js';
import type { ParaRecord } from '../../model.js';
import { resolvePairingSettings } from '../config.js';

/** Which side of the pairing a document came from. */
export type BundleSide = 'a' | 'b';

/**
 * What format a bundle document is.
 *
 * `other` exists so an unexpected extension is a *shown, unrunnable* row rather
 * than an exception during ingest — a document the user handed us must appear
 * in the grid whatever it turns out to be.
 */
export type SourceKind = 'docx' | 'pdf' | 'other';

/** Classify by extension, which is all that is known before the bytes are read. */
export function sourceKindOf(fileName: string): SourceKind {
  const lower = fileName.toLowerCase();
  if (lower.endsWith('.docx')) return 'docx';
  if (lower.endsWith('.pdf')) return 'pdf';
  return 'other';
}

/**
 * What pairing knows about one document.
 *
 * `sourceKind` is the format-agnostic seam (spec §2.1). Cross-format PAIRING is
 * designed for; cross-format COMPARISON is not built, and the two must not be
 * confused. A `.pdf` in a bundle therefore gets a real sketch — so it can be
 * paired, scored and shown — but a filename-only one: extracting its text to
 * feed the content signal would be building half of the comparison ADR-0017
 * gated behind an alignment bar that was never measured. It pairs on its name,
 * and the pair it lands in is marked not-comparable (`propose.ts`).
 */
export interface PairingSketch {
  docId: string;
  side: BundleSide;
  sourceKind: SourceKind;
  fileName: string;
  byteLength: number;
  /** `fnv1a` of `normText`, boilerplate excluded (see below). */
  paraHashes: ReadonlySet<string>;
  /** Normalised defined terms, from the Defined Terms harvester. */
  definedTerms: ReadonlySet<string>;
  /** Normalised party-role labels, from the party-role harvester. */
  partyNames: ReadonlySet<string>;
  /**
   * How many paragraphs the preprocess pass quarantined. A number, kept so the
   * grid and the index can report it without re-reading the document.
   *
   * The **records themselves are deliberately not here.** They were, so that a
   * confirmed run would not preprocess a document a second time (spec §6.7) —
   * and a review measured what that costs: at the §4.4 cap of fifty documents a
   * side, a hundred document-sized `ParaRecord[]` (each record carrying its
   * paragraph's `rawXml`) stay live from the moment the pairing is proposed
   * until the driver is discarded. Spec §4.3 opens by naming that exact shape
   * as "an out-of-memory crash, not a slowdown", and its whole policy is built
   * to keep the peak at one pair's working set times the pool size. A second
   * preprocess pass per document is a slowdown. The two are not comparable, so
   * the records are built per pair, inside the lane that needs them.
   */
  quarantinedCount: number;
  /**
   * Set when this document could not be read at all — a file named `.docx`
   * that is not a Word package, a truncated download, an encrypted one.
   *
   * The document still gets a sketch, and still gets a row in the proposal. It
   * is simply not a pairing candidate: there is nothing to pair on and nothing
   * to compare. Spec §8 case 2 is the rule — one bad file must not lose
   * ninety-nine good redlines — and before this field existed a single corrupt
   * document threw out of the bundle loop and left `proposal` null, so the
   * whole grid disappeared and every other document with it.
   */
  unreadable?: string;
}

export interface BuildSketchInput {
  docId: string;
  side: BundleSide;
  fileName: string;
  bytes: ArrayBuffer;
  config: CompareConfig;
  profile: JurisdictionProfile;
  flatten: FlattenPolicy;
}

/**
 * Whether a paragraph earns a place in the content signal.
 *
 * Empty paragraphs, single-word headings and numbering stubs are identical
 * across every document in a bundle. Left in, they pull every Jaccard score
 * toward every other and flatten exactly the distinctions pairing exists to
 * make. The default four-token floor matches the absorbed-paragraph guard
 * ADR-0032 already uses, for the same reason: below it there is not enough text
 * to carry meaning. It comes from `compare.batch.pairing.minSketchTokens` via
 * the config resolver, not from the constant — a profile that sets it must
 * actually change what a sketch contains (spec §7).
 */
export function isSketchWorthy(record: ParaRecord, config: CompareConfig): boolean {
  return (
    contentTokens(record.normText, config).length
      >= resolvePairingSettings(config).minSketchTokens
  );
}

/** Lowercase and collapse whitespace, so casing drift does not split a term. */
function normaliseName(value: string): string {
  return value.trim().toLowerCase().replace(/\s+/g, ' ');
}

/**
 * A sketch for a document that could not be read.
 *
 * Deliberately a sketch rather than an absence. The grid's rule is that every
 * document the user handed us appears exactly once, and a file that turned out
 * not to be a Word package is still a file they handed us — dropping it would
 * hide it, and throwing loses the rest of the bundle with it (spec §8 case 2).
 * It carries no hashes and no cast, so it can pair with nothing; `propose.ts`
 * keeps it out of the assignment and gives it its own row with this reason.
 */
export function unreadableSketch(identity: SketchIdentity, reason: string): PairingSketch {
  return { ...filenameOnlySketch(identity), unreadable: reason };
}

/** What a sketch needs before any bytes have been read. */
export interface SketchIdentity {
  docId: string;
  side: BundleSide;
  fileName: string;
  byteLength: number;
}

/**
 * A sketch with no content signal — the filename and nothing else.
 *
 * Two callers, for two different reasons. A `.pdf` gets one because extracting
 * its text would be building half the cross-format comparison ADR-0017 gates. A
 * `.docx` over the per-document cap gets one because it is never going to be
 * compared, and spec §4.4 asks the guard to fail *before* any partial work:
 * preprocessing it anyway would read every byte and then retain its records for
 * the driver's life, which is exactly the peak §4.3 exists to bound.
 */
export function filenameOnlySketch(identity: SketchIdentity): PairingSketch {
  return {
    docId: identity.docId,
    side: identity.side,
    sourceKind: sourceKindOf(identity.fileName),
    fileName: identity.fileName,
    byteLength: identity.byteLength,
    paraHashes: new Set<string>(),
    definedTerms: new Set<string>(),
    partyNames: new Set<string>(),
    quarantinedCount: 0,
  };
}

/** One bundle document's preprocessed records. */
export interface DocumentRecords {
  paras: ParaRecord[];
  quarantinedCount: number;
}

/**
 * Preprocess one bundle document's bytes.
 *
 * The seam the run uses. A pair's two documents are preprocessed when a lane
 * picks that pair up and released when it finishes, so the peak is bounded by
 * the pool rather than by the bundle (spec §4.3). Sketch building uses it too,
 * and keeps only the hashes.
 */
export function preprocessBundleDocument(
  bytes: ArrayBuffer,
  config: CompareConfig,
  flatten: FlattenPolicy,
): DocumentRecords {
  const parts = extractParts(new PizZipBundle(bytes));
  // ADR-0042 — the same numbering context `runPipeline` builds for a single
  // pair. Without it a batch comparison of the very same two documents reports
  // a re-levelled clause as no change at all, while the single-pair panel
  // reports it: one product answering differently depending on which panel the
  // user opened, and the silent answer is the wrong one.
  return preprocessSide(
    parts.documentXml,
    'document',
    flatten,
    config,
    buildNumberingContext(parts.stylesXml, parts.numberingXml),
  );
}

/**
 * Build a sketch from one `.docx`'s bytes.
 *
 * The cast signals are best-effort: a document the harvesters cannot read
 * yields empty sets rather than an error, because the cast is a tie-breaker
 * (spec §6.3) and losing it must never cost a pairing that content and
 * filename can decide on their own.
 */
export function buildPairingSketch(input: BuildSketchInput): PairingSketch {
  const sourceKind = sourceKindOf(input.fileName);

  // A non-`.docx` pairs on its filename alone. Reading its content would mean
  // building the cross-format content signal, which is the comparison side of
  // the line spec §2.1 draws — and the empty hash set is honest rather than
  // lossy: `jaccard` scores an empty set 0, so such a document never
  // auto-pairs on content it does not have.
  if (sourceKind !== 'docx') {
    return filenameOnlySketch({
      docId: input.docId,
      side: input.side,
      fileName: input.fileName,
      byteLength: input.bytes.byteLength,
    });
  }

  const bundle = new PizZipBundle(input.bytes);
  const parts = extractParts(bundle);

  // The records are hashed and then DROPPED — see `quarantinedCount` on the
  // type for why keeping them is the one thing this module must not do.
  //
  // Deliberately NO numbering context (ADR-0042), unlike the comparison seam
  // above. Pairing reads `record.hash` and nothing else, and hashes are text —
  // resolving numbering here would parse two more parts per document on a path
  // that runs M×N times, to build a field the next line throws away.
  const { paras, quarantinedCount } = preprocessSide(
    parts.documentXml,
    'document',
    input.flatten,
    input.config,
  );

  const paraHashes = new Set<string>();
  for (const record of paras) {
    if (isSketchWorthy(record, input.config)) paraHashes.add(record.hash);
  }

  const definedTerms = new Set<string>();
  const partyNames = new Set<string>();
  try {
    const snapshot = loadDocSnapshotFromOoxml(bundle);
    const paraNodes = [...snapshot.paragraphs, ...snapshot.footnotes];

    for (const key of analyseDefinedTerms(paraNodes, input.profile).definitionsMap.keys()) {
      definedTerms.add(normaliseName(key));
    }
    for (const role of analysePartyRole(paraNodes, input.profile).roles) {
      partyNames.add(normaliseName(role.key));
    }
  } catch {
    // Best-effort by design: the cast is a tie-breaker, and a document whose
    // definitions cannot be harvested must still pair on content and filename.
  }

  return {
    docId: input.docId,
    side: input.side,
    sourceKind,
    fileName: input.fileName,
    byteLength: input.bytes.byteLength,
    paraHashes,
    definedTerms,
    partyNames,
    quarantinedCount,
  };
}
