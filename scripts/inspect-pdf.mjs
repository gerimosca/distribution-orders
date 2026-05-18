import { readFileSync } from 'node:fs';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';

const path = process.argv[2] ?? 'PO071467.pdf';
const data = Uint8Array.from(readFileSync(path));
const doc = await getDocument({ data, useSystemFonts: true }).promise;

console.log(`=== ${path} :: ${doc.numPages} pages ===`);
const maxPages = Number(process.argv[3] ?? doc.numPages);
for (let p = 1; p <= Math.min(doc.numPages, maxPages); p++) {
  const page = await doc.getPage(p);
  const tc = await page.getTextContent();
  const lines = new Map();
  for (const it of tc.items) {
    if (!it.str || !it.str.trim()) continue;
    const x = it.transform[4];
    const y = Math.round(it.transform[5]);
    if (!lines.has(y)) lines.set(y, []);
    lines.get(y).push({ x, s: it.str });
  }
  const ys = [...lines.keys()].sort((a, b) => b - a);
  console.log(`\n----- PAGE ${p} -----`);
  for (const y of ys) {
    const parts = lines.get(y).sort((a, b) => a.x - b.x);
    console.log(`y=${y}  ` + parts.map((q) => `[x=${Math.round(q.x)}]${q.s}`).join(' '));
  }
}
