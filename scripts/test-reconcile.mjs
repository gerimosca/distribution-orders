import { readFileSync, writeFileSync } from 'node:fs';
import * as XLSX from 'xlsx';
import { getClient } from '../lib/parsers/index.mjs';

const client = getClient('kurt-geiger');

const parsed = client.parse(readFileSync('AW26 ALOHA PRE.xlsx'));

const ocFiles = ['Order - S23609.pdf', 'Order - S23610.pdf'];
const ocs = [];
for (const f of ocFiles) {
  ocs.push(await client.parseOc(readFileSync(f), f));
}

const excelTotals = client.excelTotalsBySku(parsed);
parsed.discrepancies = client.reconcile(excelTotals, ocs);
parsed.ocPricing = client.ocPricingBySku(ocs).pricing;

const d = parsed.discrepancies;
console.log('Column labels:', d.columnLabels);
console.log('SKU que cuadran (ok):', d.ok);
console.log('Discrepancias:', d.rows.length);
console.log('Genéricas:', d.generic);
console.log('\n--- Pestaña Discrepancias (primeras 25 filas) ---');
console.log(['SKU', ...d.columnLabels, d.referenceLabel, 'Dif', 'Motivo'].join(' | '));
for (const r of d.rows.slice(0, 25)) {
  console.log(
    [r.sku, ...r.cols.map((v) => (v == null ? '--' : v)), r.reference || '--', r.dif, r.motivo].join(' | ')
  );
}
if (d.rows.length > 25) console.log(`... (+${d.rows.length - 25} más)`);

const wb = client.build(parsed);
console.log('\nPestañas del Excel de salida:', wb.SheetNames);
const out = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
writeFileSync('output-con-discrepancias.xlsx', out);
console.log('Generado: output-con-discrepancias.xlsx');
