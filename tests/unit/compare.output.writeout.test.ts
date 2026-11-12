/**
 * Feature 08 — Compare C2: materialise / write-out.
 * Installs a HostPort via setActiveHost and asserts materialise routes the write
 * through getActiveHost().writeBack (12f). Verifies mode mapping + body replacement.
 * Acceptance: T11 — default = new doc; insert path = 'replace' with fallback.
 */
import { describe, it, expect, vi, afterEach } from 'vitest';
import { PizZipBundle } from '@shared/ooxml/service.js';
import { setActiveHost, _resetActiveHost } from '@shared/host/index.js';
import type { HostPort } from '@shared/host/index.js';
import { materialise, replaceBody } from '@features/compare/output/writeout.js';

const AUTHOR = 'Mutatis Compare';
const DATE = '2026-06-08T10:00:00Z';

/** Install a stub host exposing only a spied writeBack; returns the spy. */
function installHostWithWriteBack() {
  const writeBack = vi.fn().mockResolvedValue(undefined);
  setActiveHost({ writeBack } as unknown as HostPort);
  return writeBack;
}

const DOC_XML =
  '<?xml version="1.0"?>' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:body><w:p><w:r><w:t>original</w:t></w:r></w:p></w:body>' +
  '</w:document>';

function makeBundle(): PizZipBundle {
  const b = new PizZipBundle();
  b.setPart('word/document.xml', DOC_XML);
  b.setPart('word/styles.xml', '<w:styles/>');
  return b;
}

afterEach(() => {
  _resetActiveHost();
  vi.restoreAllMocks();
});

describe('replaceBody', () => {
  it('swaps the <w:body> subtree while preserving the document root', () => {
    const out = replaceBody(DOC_XML, '<w:body><w:p><w:r><w:t>new</w:t></w:r></w:p></w:body>');
    expect(out).toContain('<w:document');
    expect(out).toContain('xmlns:w=');
    expect(out).toContain('<w:t>new</w:t>');
    expect(out).not.toContain('<w:t>original</w:t>');
  });

  it('throws when there is no <w:body>', () => {
    expect(() => replaceBody('<w:document></w:document>', '<w:body/>')).toThrow();
  });
});

describe('materialise — newDocument (default)', () => {
  it('calls writeBack with mode newDocument', async () => {
    const spy = installHostWithWriteBack();
    await materialise(
      '<w:body><w:p><w:r><w:t>redline</w:t></w:r></w:p></w:body>',
      makeBundle(),
      { mode: 'newDocument', author: AUTHOR, date: DATE },
    );
    expect(spy).toHaveBeenCalledTimes(1);
    const [bundleArg, optsArg] = spy.mock.calls[0]!;
    expect(optsArg).toEqual({ mode: 'newDocument' });
    // The authored body landed in the cloned bundle.
    expect((bundleArg as PizZipBundle).part('word/document.xml')).toContain('<w:t>redline</w:t>');
    // Cloned bundle preserves other parts (styles).
    expect((bundleArg as PizZipBundle).part('word/styles.xml')).toBe('<w:styles/>');
  });
});

describe('materialise — insertIntoActive', () => {
  it('calls writeBack with mode replace and the artifact fallback', async () => {
    const spy = installHostWithWriteBack();
    await materialise(
      '<w:body><w:p/></w:body>',
      makeBundle(),
      { mode: 'insertIntoActive', onArtifactDetected: 'fallbackToNewDoc', author: AUTHOR, date: DATE },
    );
    expect(spy).toHaveBeenCalledTimes(1);
    const [, optsArg] = spy.mock.calls[0]!;
    expect(optsArg).toEqual({ mode: 'replace', onArtifactDetected: 'fallbackToNewDoc' });
  });

  it('defaults onArtifactDetected to fallbackToNewDoc when omitted', async () => {
    const spy = installHostWithWriteBack();
    await materialise('<w:body><w:p/></w:body>', makeBundle(), {
      mode: 'insertIntoActive',
      author: AUTHOR,
      date: DATE,
    });
    const [, optsArg] = spy.mock.calls[0]!;
    expect(optsArg).toEqual({ mode: 'replace', onArtifactDetected: 'fallbackToNewDoc' });
  });
});

describe('materialise — error path', () => {
  it('throws when the original bundle has no document.xml', async () => {
    const empty = new PizZipBundle();
    await expect(
      materialise('<w:body/>', empty, { mode: 'newDocument', author: AUTHOR, date: DATE }),
    ).rejects.toThrow(/document\.xml/);
  });
});
