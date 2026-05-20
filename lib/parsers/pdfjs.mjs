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
import * as pdfjsWorker from 'pdfjs-dist/legacy/build/pdf.worker.mjs';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';

globalThis.pdfjsWorker = pdfjsWorker;

export { getDocument };
