// Centralised pdfjs import.
//
// In Node, pdf.js falls back to a "fake worker" by dynamically importing
// pdf.worker.mjs (see pdf.mjs: PDFWorker._setupFakeWorkerGlobal → `await
// import(this.workerSrc)`). That dynamic path is invisible to Next.js's
// file tracer, so on Vercel the worker file isn't shipped to /var/task/
// and getDocument() blows up with "Cannot find module ...pdf.worker.mjs".
//
// Static-importing the worker fixes both ends: it gives the tracer a real
// dependency edge to follow, and pre-populates globalThis.pdfjsWorker so
// PDFWorker skips the dynamic import entirely (pdf.mjs:17508).
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { pathToFileURL } from 'node:url';
import * as pdfjsWorker from 'pdfjs-dist/legacy/build/pdf.worker.mjs';
import { getDocument as pdfjsGetDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';

globalThis.pdfjsWorker = pdfjsWorker;

// Some PDFs use built-in fonts (Helvetica, Times, etc.) that aren't embedded.
// On Vercel's Lambda there are no system fonts to substitute, so pdfjs needs
// the standard fonts shipped inside pdfjs-dist. Without this, text from those
// fonts can come back missing or as junk glyphs — which is what we suspect
// is making "S100917-06" disappear from the Office Holdings POs on Vercel.
const require = createRequire(import.meta.url);
const standardFontsDir = join(
  dirname(require.resolve('pdfjs-dist/legacy/build/pdf.mjs')),
  '..',
  '..',
  'standard_fonts'
);
const STANDARD_FONT_DATA_URL = pathToFileURL(standardFontsDir).href + '/';

export function getDocument(params) {
  return pdfjsGetDocument({
    standardFontDataUrl: STANDARD_FONT_DATA_URL,
    ...params,
  });
}
