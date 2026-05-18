import { readFileSync, writeFileSync } from 'node:fs';
import * as XLSX from 'xlsx';
import { parseKurtGeiger, buildWorkbook } from '../lib/parsers/kurt-geiger.mjs';

const inputPath = process.argv[2] ?? 'AW26 ALOHA PRE.xlsx';
const outputPath = process.argv[3] ?? 'output-kurt-geiger.xlsx';

const buf = readFileSync(inputPath);
const parsed = parseKurtGeiger(buf);

console.log('Parsed warehouses:');
for (const wh of parsed.warehouses) {
  const w = wh.warnings.length ? ` ⚠${wh.warnings.length}` : '';
  console.log(`  - [${wh.kind}] ${wh.warehouse}: ${wh.items.length} líneas SKU, qty = ${wh.totalQty}, POs = [${wh.pos.join(', ')}]${w}`);
}
console.log(`\nGrand total: ${parsed.grandTotal}`);

const totalWarnings = parsed.warehouses.reduce((s, w) => s + w.warnings.length, 0);
console.log(`\nAvisos por pestaña: ${totalWarnings}`);
for (const wh of parsed.warehouses) {
  for (const msg of wh.warnings) {
    console.log(`  ⚠ [${wh.warehouse}] ${msg}`);
  }
}
if (parsed.globalIssues.length) {
  console.log(`\nIncidencias globales: ${parsed.globalIssues.length}`);
  for (const m of parsed.globalIssues) console.log('  ⚠ ' + m);
}

const wb = buildWorkbook(parsed);
const out = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
writeFileSync(outputPath, out);
console.log(`\nGenerado: ${outputPath}`);
