import PizZip from 'pizzip';

/**
 * Build a minimal but genuinely valid `.docx` in memory.
 *
 * Tests need real OOXML, not a stub: the compare pipeline parses the package,
 * and a fixture that skips the content-types part or the package relationship
 * would fail in a way that tells us nothing about our own code. This is the
 * smallest package Word and the engine both accept.
 *
 * Self-contained on purpose: a unit suite that needs another repository's test
 * data breaks for reasons unrelated to the code under test.
 */

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;

const PACKAGE_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOCUMENT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
</w:styles>`;

/** A4 with 1-inch margins, in twentieths of a point. */
const SECT_PR =
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';

function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

export interface Para {
  readonly text: string;
  readonly styleId?: string;
  /** Preserve leading/trailing spaces, as Word requires for such runs. */
  readonly preserveSpace?: boolean;
}

function paragraphXml(para: Para): string {
  const ppr = para.styleId ? `<w:pPr><w:pStyle w:val="${para.styleId}"/></w:pPr>` : '';
  const space = para.preserveSpace ? ' xml:space="preserve"' : '';
  return `<w:p>${ppr}<w:r><w:t${space}>${escapeXml(para.text)}</w:t></w:r></w:p>`;
}

/** Build a `.docx` as raw bytes from an ordered list of paragraphs. */
export function makeDocxBytes(paras: readonly (Para | string)[]): Uint8Array {
  const body = paras
    .map((p) => paragraphXml(typeof p === 'string' ? { text: p } : p))
    .join('');

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}${SECT_PR}</w:body></w:document>`;

  const zip = new PizZip();
  zip.file('[Content_Types].xml', CONTENT_TYPES);
  zip.file('_rels/.rels', PACKAGE_RELS);
  zip.file('word/_rels/document.xml.rels', DOCUMENT_RELS);
  zip.file('word/document.xml', documentXml);
  zip.file('word/styles.xml', STYLES);

  return zip.generate({ type: 'uint8array', compression: 'DEFLATE' });
}

/** The same bytes wrapped as a `File`, the way the UI receives them. */
export function makeDocxFile(name: string, paras: readonly (Para | string)[]): File {
  const bytes = makeDocxBytes(paras);
  return new File([bytes as BlobPart], name, {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

/**
 * A small agreement pair with one change of each kind the classifier cares
 * about: a numeric value, an obligation verb, a defined term, an inserted
 * clause and a deleted clause.
 */
export const SAMPLE_ORIGINAL: readonly (Para | string)[] = [
  { text: 'SERVICES AGREEMENT', styleId: 'Heading1' },
  'This Agreement is made on 1 January 2026 between Alpha Pte Ltd and Beta Limited.',
  'The Supplier shall deliver the Services within 30 days of the Commencement Date.',
  'The Fee payable under this Agreement is £40,000.',
  'The Customer may terminate this Agreement on 60 days notice.',
  'This clause will be deleted in the revised version.',
];

export const SAMPLE_REVISED: readonly (Para | string)[] = [
  { text: 'SERVICES AGREEMENT', styleId: 'Heading1' },
  'This Agreement is made on 1 January 2026 between Alpha Pte Ltd and Beta Limited.',
  'The Supplier must deliver the Services within 45 days of the Commencement Date.',
  'The Fee payable under this Agreement is £55,000.',
  'The Customer may terminate this Agreement on 60 days notice.',
  'A wholly new clause has been inserted here by the revising party.',
];
