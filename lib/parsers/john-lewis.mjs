import * as XLSX from 'xlsx';
import { buildDiscrepancySheet } from './discrepancy-sheet.mjs';

// John Lewis distribution. Main input: one client Excel ("Distribution John
// Lewis.xlsx"); the optional Aloha OC PDFs supply PRICE/DISCOUNT and validate
// the quantities.
//
// The Excel carries an "All Orders" master sheet plus one sheet per PO. Each
// data row is one size of one style for one PO:
//   - PO Number               -> the tab the row belongs to
//   - Supplier Model Number   -> the SKU base (e.g. "S101329-03")
//   - Ordering Description     -> ends in the EU size ("... EU37"); last 2
//                                 chars give the size ("37")
//   - Order quantity (Units)   -> units
// Final SKU = base + size, no separator: "S101329-03" + "37" = "S101329-0337"
// — the exact form the Aloha OC prints (e.g. "[S101329-0337]").

// Column headers we anchor on (matched after whitespace-normalisation, so a
// trailing CR/LF or doubled space in the export doesn't break detection).
const COL_PO = 'PO Number';
const COL_DESC = 'Ordering Description';
const COL_MODEL = 'Supplier Model Number';
const COL_QTY = 'Order quantity (Units)';

function norm(s) {
  return String(s ?? '').replace(/\s+/g, ' ').trim();
}

// Locate the header row + the columns we need within a sheet's row matrix.
// Returns null if the sheet isn't a PO-detail sheet (e.g. an empty tab).
function detectLayout(rows) {
  let headerIdx = -1;
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    if (row && row.some((v) => norm(v) === COL_PO)) {
      headerIdx = r;
      break;
    }
  }
  if (headerIdx < 0) return null;
  const header = rows[headerIdx];
  const col = (name) => header.findIndex((v) => norm(v) === name);
  const layout = {
    headerIdx,
    poCol: col(COL_PO),
    descCol: col(COL_DESC),
    modelCol: col(COL_MODEL),
    qtyCol: col(COL_QTY),
  };
  if (layout.descCol < 0 || layout.modelCol < 0 || layout.qtyCol < 0) return null;
  return layout;
}

// Strip the size off a full SKU to get the style base ("S101329-0337" ->
// "S101329-03"). Used so PRICE/DISCOUNT can fall back to the style price when
// the OC didn't price that exact size.
function skuBase(sku) {
  return sku.slice(0, -2);
}

export function parseJohnLewis(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });

  // The "All Orders" sheet is the master list of every PO row. When present
  // we read only it (reading the per-PO sheets too would double-count). If a
  // future export drops it, fall back to the union of the other sheets.
  const masterName = wb.SheetNames.find((n) => norm(n).toLowerCase() === 'all orders');
  const sheetNames = masterName ? [masterName] : wb.SheetNames;

  const globalIssues = [];
  const poOrder = []; // first-seen order of PO numbers
  const poData = new Map(); // poNumber -> { items: Map<sku, qty>, warnings: [] }
  let rowsSeen = 0;

  for (const name of sheetNames) {
    const ws = wb.Sheets[name];
    if (!ws) continue;
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: false });
    const L = detectLayout(rows);
    if (!L) continue;

    for (let r = L.headerIdx + 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row) continue;
      const poRaw = L.poCol >= 0 ? row[L.poCol] : null;
      const model = row[L.modelCol];
      const desc = row[L.descCol];
      // A data row needs at least a model + description; skip subtotal/blank
      // rows that lack them.
      if (model == null || desc == null || norm(model) === '' || norm(desc) === '') {
        continue;
      }
      rowsSeen++;

      const po = norm(poRaw) || '(sin PO)';
      const base = norm(model);
      const size = norm(desc).slice(-2);
      if (!/^\d{2}$/.test(size)) {
        globalIssues.push(
          `Fila ${r + 1} (${base}): no pude leer la talla de "${norm(desc)}" (esperaba 2 dígitos al final).`
        );
        continue;
      }
      const sku = `${base}${size}`;

      const qtyRaw = row[L.qtyCol];
      const qty = Math.round(Number(qtyRaw));
      if (!Number.isFinite(qty)) {
        globalIssues.push(`Fila ${r + 1} (${sku}): cantidad no numérica ("${qtyRaw}").`);
        continue;
      }

      if (!poData.has(po)) {
        poData.set(po, { items: new Map(), warnings: [] });
        poOrder.push(po);
      }
      const d = poData.get(po);
      d.items.set(sku, (d.items.get(sku) ?? 0) + qty);
    }
  }

  if (rowsSeen === 0) {
    throw new Error(
      'No reconozco el formato: no encontré filas con "PO Number" / "Supplier Model Number" / "Ordering Description".'
    );
  }

  // Sort POs ascending so the tabs come out in a predictable order.
  poOrder.sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));

  let grandTotal = 0;
  const pos = poOrder.map((poNumber) => {
    const d = poData.get(poNumber);
    const items = [...d.items].map(([sku, qty]) => ({ sku, qty }));
    const totalQty = items.reduce((s, it) => s + it.qty, 0);
    grandTotal += totalQty;
    return { poNumber, items, warnings: d.warnings, totalQty };
  });

  return { pos, grandTotal, globalIssues };
}

// Total qty per SKU across every PO (the Excel side of the reconciliation).
export function excelTotalsBySku(parsed) {
  const total = new Map();
  for (const po of parsed.pos) {
    for (const it of po.items) total.set(it.sku, (total.get(it.sku) ?? 0) + it.qty);
  }
  return total;
}

// Reconcile the distribution against the OC. Mirrors the Office Holdings sheet:
// one column per PO, the OC total as the reference, dif = sum(POs) - OC.
export function reconcile(ocs, pos) {
  const ocTotals = new Map();
  for (const oc of ocs) {
    for (const [sku, qty] of oc.items) ocTotals.set(sku, (ocTotals.get(sku) ?? 0) + qty);
  }
  const ocLabel = ocs.length ? 'OC ' + ocs.map((o) => o.orderNumber).join(' + ') : 'OC';
  const columnLabels = pos.map((p) => String(p.poNumber).slice(0, 31));
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
      motivo = `Solo en OC Alohas, ningún PO de John Lewis lo reparte (faltan ${ocTotal} en distribución)`;
    } else if (ocTotal === 0) {
      rank = 3;
      motivo = 'En distribución John Lewis, sin OC Alohas';
    } else {
      rank = 1;
      motivo =
        dif > 0
          ? `John Lewis reparte ${dif} más de lo que Alohas confirma (sobran en distribución)`
          : `Alohas confirma ${-dif} más de lo que John Lewis reparte (faltan por repartir)`;
    }
    rows.push({ sku, cols, reference: ocTotal, dif, motivo, rank });
  }
  rows.sort((a, b) => a.rank - b.rank || a.sku.localeCompare(b.sku));

  // GEN* lines from the OC: surfaced but never reconciled.
  const genericSkus = new Set();
  for (const o of ocs) for (const s of o.generic.keys()) genericSkus.add(s);
  const generic = [...genericSkus].sort().map((sku) => ({
    sku,
    perCol: pos.map(() => null),
  }));

  return {
    rows,
    ok,
    generic,
    genericLabel: 'Líneas genéricas de la OC (no conciliadas):',
    columnLabels,
    referenceLabel: ocLabel,
    title: 'DISCREPANCIAS — Distribución John Lewis vs OC Alohas',
    legend:
      'Diferencia = (suma de las POs John Lewis) − (OC Alohas). ' +
      'Positiva: John Lewis reparte MÁS de lo que Alohas confirma (sobran en distribución). ' +
      'Negativa: Alohas confirma MÁS de lo que John Lewis reparte (faltan por repartir).',
    summaryLabels: {
      real: 'Reales John Lewis ≠ Alohas',
      refOnly: 'Solo en OC Alohas',
      colOnly: 'Solo en distribución John Lewis',
    },
  };
}

export function build(parsed) {
  const wb = XLSX.utils.book_new();
  const usedTabs = new Set();

  if (parsed.discrepancies) {
    const ws = XLSX.utils.aoa_to_sheet(buildDiscrepancySheet(parsed.discrepancies));
    XLSX.utils.book_append_sheet(wb, ws, 'Discrepancias');
    usedTabs.add('Discrepancias');
  }

  // PRICE/DISCOUNT come from the OC, keyed by full SKU. Build a style-base
  // fallback too, so a size the OC didn't price still inherits its style's
  // price (price/discount are per style on these orders).
  const pricing = parsed.ocPricing; // Map<sku,{price,disc}> | undefined
  const basePricing = new Map();
  if (pricing) {
    for (const [sku, v] of pricing) {
      const b = skuBase(sku);
      if (!basePricing.has(b)) basePricing.set(b, v);
    }
  }
  const priceOf = (sku) => {
    const v = (pricing && pricing.get(sku)) || basePricing.get(skuBase(sku));
    return [v && v.price != null ? v.price : '', v && v.disc != null ? v.disc : ''];
  };

  const globalIssues = parsed.globalIssues ?? [];

  // Excel caps tab names at 31 chars and forbids a few characters. PO numbers
  // are short and numeric, but disambiguate just in case two collide.
  const tabNameFor = (poNumber) => {
    const base = String(poNumber).replace(/[\\/?*[\]:]/g, '-').slice(0, 31) || 'PO';
    if (!usedTabs.has(base)) {
      usedTabs.add(base);
      return base;
    }
    for (let n = 2; ; n++) {
      const suffix = ` (${n})`;
      const candidate = base.slice(0, 31 - suffix.length) + suffix;
      if (!usedTabs.has(candidate)) {
        usedTabs.add(candidate);
        return candidate;
      }
    }
  };

  parsed.pos.forEach((po, idx) => {
    const aoa = [];
    aoa.push(['SKU', 'QUANTITY', 'PRICE', 'DISCOUNT', null, 'PO:', po.poNumber]);
    const warningParts = [
      ...(idx === 0 ? globalIssues.map((s) => `[Fichero] ${s}`) : []),
      ...(po.warnings ?? []),
    ];
    const warningText = warningParts.length ? 'Revisar: ' + warningParts.join(' | ') : null;

    const agg = new Map();
    for (const it of po.items) agg.set(it.sku, (agg.get(it.sku) ?? 0) + it.qty);
    const entries = [...agg].sort((a, b) => a[0].localeCompare(b[0]));

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
    XLSX.utils.book_append_sheet(wb, ws, tabNameFor(po.poNumber));
  });

  return wb;
}
