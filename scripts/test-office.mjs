import { readFileSync, writeFileSync } from 'node:fs';
import * as XLSX from 'xlsx';
import { getClient } from '../lib/parsers/index.mjs';

const client = getClient('office-holdings');

const poFiles = ['PO071467.pdf', 'PO071468.pdf'];
const pos = [];
for (const f of poFiles) {
  pos.push(await client.parsePo(readFileSync(f), f.replace(/\.pdf$/i, '')));
}

console.log('=== POs ===');
for (const po of pos) {
  console.log(`\n${po.poNumber} (total ${po.totalQty})${po.warnings.length ? ' ⚠ ' + po.warnings.join(' | ') : ''}`);
  for (const it of po.items) console.log(`  ${it.sku} = ${it.qty}`);
}

const ocs = [await client.parseOc(readFileSync('Order - S21936.pdf'), 'S21936')];
console.log('\n=== OC ===', ocs[0].orderNumber, '| SKU:', ocs[0].items.size);

const parsed = {
  pos,
  discrepancies: client.reconcile(ocs, pos),
  ocPricing: client.ocPricingBySku(ocs).pricing,
};
const d = parsed.discrepancies;
console.log('\n=== Conciliación ===');
console.log('referenceLabel:', d.referenceLabel, '| columnas:', d.columnLabels);
console.log('Cuadran:', d.ok, '| Discrepancias:', d.rows.length);
for (const r of d.rows) {
  console.log([r.sku, ...r.cols.map((v) => v ?? '--'), r.reference, r.dif, r.motivo].join(' | '));
}

const wb = client.build(parsed);
console.log('\nPestañas:', wb.SheetNames);
writeFileSync('output-office.xlsx', XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
const a = XLSX.utils.sheet_to_json(wb.Sheets['Discrepancias'], { header: 1, defval: null });
console.log('\n--- Pestaña Discrepancias ---');
a.forEach((r) => console.log(JSON.stringify(r)));
const t = XLSX.utils.sheet_to_json(wb.Sheets['PO071467'], { header: 1, defval: null });
console.log('\n--- Pestaña PO071467 ---');
t.forEach((r) => console.log(JSON.stringify(r)));
