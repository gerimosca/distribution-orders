import { readFileSync, writeFileSync } from 'node:fs';
import * as XLSX from 'xlsx';
import { getClient } from '../lib/parsers/index.mjs';
import { createZip, safeFileName } from '../lib/zip.mjs';

const client = getClient('john-lewis');

const excelFile = 'Distribution John Lewis.xlsx';
const ocFiles = ['Order - S21508.pdf', 'Order - S21510.pdf'];

const parsed = client.parse(readFileSync(excelFile));
console.log('=== POs (Excel) ===');
let excelGrand = 0;
for (const po of parsed.pos) {
  excelGrand += po.totalQty;
  console.log(`\n${po.poNumber} (total ${po.totalQty})${po.warnings.length ? ' ⚠ ' + po.warnings.join(' | ') : ''}`);
  for (const it of po.items) console.log(`  ${it.sku} = ${it.qty}`);
}
console.log(`\nPOs=${parsed.pos.length}  grandTotal=${parsed.grandTotal}  (sum check ${excelGrand})`);
if (parsed.globalIssues.length) console.log('globalIssues:', parsed.globalIssues);

const ocs = [];
for (const f of ocFiles) ocs.push(await client.parseOc(readFileSync(f), f.replace(/\.pdf$/i, '')));
let ocGrand = 0;
for (const oc of ocs) {
  let t = 0;
  for (const q of oc.items.values()) t += q;
  ocGrand += t;
  console.log(`\n=== OC ${oc.orderNumber} | SKUs=${oc.items.size} | priced=${oc.pricing.size} | units=${t} ===`);
}
console.log(`OC grand units=${ocGrand}`);

parsed.discrepancies = client.reconcile(ocs, parsed.pos);
parsed.ocPricing = client.ocPricingBySku(ocs).pricing;

const d = parsed.discrepancies;
console.log('\n=== Conciliación ===');
console.log('referenceLabel:', d.referenceLabel);
console.log('columnas:', d.columnLabels.join(', '));
console.log('Cuadran:', d.ok, '| Discrepancias:', d.rows.length);
for (const r of d.rows) {
  console.log([r.sku, ...r.cols.map((v) => v ?? '--'), r.reference, r.dif, r.motivo].join(' | '));
}

const wb = client.build(parsed);
console.log('\nPestañas:', wb.SheetNames.join(', '));
writeFileSync('output-john-lewis.xlsx', XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));

// Spot-check one PO tab (PRICE/DISCOUNT must be filled from the OC).
const sample = parsed.pos[0].poNumber;
console.log(`\n--- Pestaña ${sample} ---`);
XLSX.utils
  .sheet_to_json(wb.Sheets[sample], { header: 1, defval: null })
  .forEach((r) => console.log(JSON.stringify(r)));

console.log('\n--- Pestaña Discrepancias ---');
XLSX.utils
  .sheet_to_json(wb.Sheets['Discrepancias'], { header: 1, defval: null })
  .forEach((r) => console.log(JSON.stringify(r)));

// ZIP of CSVs (one per tab) — mirrors the route's output=zip path.
const files = wb.SheetNames.map((name) => ({
  name: `${safeFileName(name)}.csv`,
  data: '﻿' + XLSX.utils.sheet_to_csv(wb.Sheets[name]),
}));
writeFileSync('output-john-lewis-csv.zip', createZip(files));
console.log(`\nZIP escrito con ${files.length} CSV(s): ${files.map((f) => f.name).join(', ')}`);

// Assertions.
if (d.rows.length === 0) console.log('\n✔ Sin discrepancias: Excel cuadra con la OC.');
else console.log(`\n⚠ ${d.rows.length} discrepancia(s).`);
const anyPriced = parsed.pos.some((po) =>
  po.items.some((it) => {
    const v = parsed.ocPricing.get(it.sku);
    return v && v.price != null;
  })
);
console.log(anyPriced ? '✔ PRICE/DISCOUNT poblados desde la OC.' : '⚠ No se pobló PRICE desde la OC.');
