import * as XLSX from 'xlsx';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { buildDiscrepancySheet } from './discrepancy-sheet.mjs';

// Office Holdings purchase-order PDF. One PDF = one PO.
// - PO number: header "ORDER NO: POxxxxxx".
// - SKU base: the FACTORY STYLE column (e.g. S100917-01).
// - Sizes: UK sizes in the column header, converted to EU.
// - Quantity: per size column, on the FACTORY STYLE row.
const UK_TO_EU = { '3.5': 36, '4': 37, '5': 38, '6': 39, '7': 40, '8': 41 };
const FACTORY_STYLE = /^S\d{3,}-\d{1,3}[A-Z]?$/;
const FACTORY_X_MIN = 150;
const FACTORY_X_MAX = 235;
const SIZE_X_TOL = 14;

// Several uploaded PDFs may resolve to the same PO number (re-upload, or two
// files with the same ORDER NO). Keep them as separate tabs with a numeric
// suffix; the "PO:" cell still shows the real number. Pure & deterministic on
// the pos array order, so build() and reconcile() agree on labels.
function uniquePoLabels(pos) {
  const seen = new Map();
  return pos.map((p) => {
    const base = String(p.poNumber);
    const n = (seen.get(base) ?? 0) + 1;
    seen.set(base, n);
    const label = n === 1 ? base : `${base} (${n})`;
    return label.replace(/[:\\/?*[\]]/g, '-').slice(0, 31);
  });
}

function linesByY(items) {
  const lines = new Map();
  for (const it of items) {
    if (!it.str || !it.str.trim()) continue;
    const x = it.transform[4];
    const y = Math.round(it.transform[5]);
    if (!lines.has(y)) lines.set(y, []);
    lines.get(y).push({ x, s: it.str.trim() });
  }
  for (const arr of lines.values()) arr.sort((a, b) => a.x - b.x);
  return lines;
}

export async function parsePoPdf(buffer, fallbackLabel) {
  const data =
    buffer instanceof ArrayBuffer ? new Uint8Array(buffer) : Uint8Array.from(buffer);
  const doc = await getDocument({ data, useSystemFonts: true }).promise;

  let poNumber = null;
  let sizeCols = null; // [{ x, eu }]
  let grandTotal = null;
  const agg = new Map(); // sku -> qty
  let itemsSum = 0;

  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p);
    const tc = await page.getTextContent();
    const lines = linesByY(tc.items);

    for (const parts of lines.values()) {
      const joined = parts.map((q) => q.s).join(' ');

      if (!poNumber) {
        const m = joined.match(/ORDER NO:\s*([A-Z0-9]+)/i);
        if (m) poNumber = m[1];
      }
      if (grandTotal == null && /GRAND TOTAL:/i.test(joined)) {
        const nums = parts.filter((q) => /^\d+$/.test(q.s));
        if (nums.length) grandTotal = Number(nums[nums.length - 1].s);
      }

      // Size header row: tokens that are exactly UK sizes we know.
      const sizeHits = parts.filter((q) => Object.hasOwn(UK_TO_EU, q.s));
      if (!sizeCols && sizeHits.length >= 4) {
        sizeCols = sizeHits.map((q) => ({ x: q.x, eu: UK_TO_EU[q.s] }));
      }

      // Item row: has a FACTORY STYLE token in its column.
      const fs = parts.find(
        (q) =>
          FACTORY_STYLE.test(q.s) && q.x >= FACTORY_X_MIN && q.x <= FACTORY_X_MAX
      );
      if (!fs || !sizeCols) continue;

      for (const col of sizeCols) {
        const tok = parts.find(
          (q) => Math.abs(q.x - col.x) <= SIZE_X_TOL && /^\d+$/.test(q.s)
        );
        if (!tok) continue;
        const qty = Number(tok.s);
        itemsSum += qty;
        if (qty <= 0) continue;
        const sku = `${fs.s}${col.eu}`;
        agg.set(sku, (agg.get(sku) ?? 0) + qty);
      }
    }
  }

  poNumber = poNumber ?? fallbackLabel ?? 'PO';
  const items = [...agg].map(([sku, qty]) => ({ sku, qty }));
  const totalQty = items.reduce((s, it) => s + it.qty, 0);
  const warnings = [];
  if (grandTotal != null && grandTotal !== totalQty) {
    warnings.push(
      `GRAND TOTAL del PDF=${grandTotal} ≠ suma de líneas=${totalQty}. Revisar.`
    );
  }
  if (!sizeCols) {
    warnings.push('No se detectó la fila de tallas (3.5 4 5 6 7 8). Revisar formato.');
  }

  return { poNumber, items, totalQty, warnings };
}

// Aggregate qty per SKU across every parsed OC (reusing oc-pdf's shape).
export function ocTotalsBySku(ocs) {
  const total = new Map();
  for (const oc of ocs) {
    for (const [sku, qty] of oc.items) {
      total.set(sku, (total.get(sku) ?? 0) + qty);
    }
  }
  return total;
}

// Columns = the POs, reference = the OC total. dif = sum(PO) - OC.
export function reconcile(ocs, pos) {
  const ocTotals = ocTotalsBySku(ocs);
  const ocLabel = ocs.length
    ? 'OC ' + ocs.map((o) => o.orderNumber).join(' + ')
    : 'OC';
  const columnLabels = uniquePoLabels(pos);
  const poMaps = pos.map((p) => {
    const m = new Map();
    for (const it of p.items) m.set(it.sku, (m.get(it.sku) ?? 0) + it.qty);
    return m;
  });

  const skus = new Set();
  for (const m of poMaps) for (const s of m.keys()) skus.add(s);
  for (const s of ocTotals.keys()) skus.add(s);

  const rows = [];
  let ok = 0;
  for (const sku of [...skus].sort()) {
    const cols = poMaps.map((m) => (m.has(sku) ? m.get(sku) : null));
    const poSum = cols.reduce((s, v) => s + (v ?? 0), 0);
    const ocTotal = ocTotals.get(sku) ?? 0;
    const inAnyPo = cols.some((v) => v != null);
    if (!inAnyPo && ocTotal === 0) continue;
    const dif = poSum - ocTotal;
    if (dif === 0 && inAnyPo) {
      ok++;
      continue;
    }
    let rank, motivo;
    if (!inAnyPo) {
      rank = 2;
      motivo = `Solo en OC, ningún PO lo trae (faltan ${ocTotal})`;
    } else if (ocTotal === 0) {
      rank = 3;
      motivo = 'En PO, sin OC';
    } else {
      rank = 1;
      motivo =
        dif > 0
          ? `PO > OC (sobran ${dif} en PO)`
          : `OC > PO (faltan ${-dif} en PO)`;
    }
    rows.push({ sku, cols, reference: ocTotal, dif, motivo, rank });
  }
  rows.sort((a, b) => a.rank - b.rank || a.sku.localeCompare(b.sku));

  return {
    rows,
    ok,
    columnLabels,
    referenceLabel: ocLabel,
    title: 'DISCREPANCIAS — POs vs OC',
    legend:
      'Diferencia = (suma de POs) − OC.  Positiva: los POs piden de más.  Negativa: la OC pide más que los POs.',
    summaryLabels: {
      real: 'Reales PO≠OC',
      refOnly: 'Solo en OC',
      colOnly: 'En PO sin OC',
    },
  };
}

export function build(parsed) {
  const wb = XLSX.utils.book_new();

  if (parsed.discrepancies) {
    const ws = XLSX.utils.aoa_to_sheet(buildDiscrepancySheet(parsed.discrepancies));
    XLSX.utils.book_append_sheet(wb, ws, 'Discrepancias');
  }

  const pricing = parsed.ocPricing; // Map<sku,{price,disc}> | undefined
  const priceOf = (sku) => {
    const v = pricing && pricing.get(sku);
    return [v && v.price != null ? v.price : '', v && v.disc != null ? v.disc : ''];
  };

  const tabLabels = uniquePoLabels(parsed.pos);
  parsed.pos.forEach((po, idx) => {
    const aoa = [];
    aoa.push(['SKU', 'QUANTITY', 'PRICE', 'DISCOUNT', null, 'PO:', po.poNumber]);
    const warningText = po.warnings && po.warnings.length
      ? 'Revisar: ' + po.warnings.join(' | ')
      : null;
    const agg = new Map();
    for (const it of po.items) agg.set(it.sku, (agg.get(it.sku) ?? 0) + it.qty);
    const entries = [...agg];
    if (entries.length === 0 && warningText) {
      aoa.push([null, null, null, null, null, 'Aviso:', warningText]);
    }
    entries.forEach(([sku, qty], i) => {
      const [price, disc] = priceOf(sku);
      if (i === 0 && warningText) {
        aoa.push([sku, qty, price, disc, null, 'Aviso:', warningText]);
      } else {
        aoa.push([sku, qty, price, disc]);
      }
    });
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    XLSX.utils.book_append_sheet(wb, ws, tabLabels[idx]);
  });
  return wb;
}
