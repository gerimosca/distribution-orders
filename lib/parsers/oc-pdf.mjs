import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';

// Parses an ALOHAS (Sunset Ventures) order-confirmation PDF.
//
// Layout is reconstructed from glyph coordinates, not from flat text: tall
// rows make pdftotext interleave the SKU and its quantity across lines, but
// in coordinate space the "[STYLE-VVSS]" anchor and its quantity share the
// same baseline y. The quantity is the only token with exactly 3 decimals
// ("14.000"); price/discount/amount use 2 decimals ("78.00", "1,092.00"),
// so the decimal count alone identifies it (its x drifts with digit count).
const SKU_ANCHOR = /^\[([A-Z0-9]+-[A-Z0-9]+)\]/;
const QTY_TOKEN = /^\d[\d,]*\.\d{3}$/;
const MONEY_TOKEN = /^\d[\d,]*\.\d{2}$/; // price / discount / amount
const COL_X_TOL = 22; // a value belongs to a column if within this of its x

// Pick the money token whose x is closest to a column anchor (the amount
// column sits far right, so a tight tolerance keeps price/disc/amount apart).
function moneyNear(parts, colX) {
  if (colX == null) return null;
  let best = null;
  let bestD = COL_X_TOL;
  for (const q of parts) {
    if (!MONEY_TOKEN.test(q.s)) continue;
    const d = Math.abs(q.x - colX);
    if (d <= bestD) {
      bestD = d;
      best = q;
    }
  }
  return best ? parseFloat(best.s.replace(/,/g, '')) : null;
}

export async function parseOcPdf(buffer, fallbackLabel) {
  // pdf.js requires a plain Uint8Array and explicitly rejects Node Buffer
  // (which subclasses Uint8Array), so always coerce to a real Uint8Array.
  const data =
    buffer instanceof ArrayBuffer ? new Uint8Array(buffer) : Uint8Array.from(buffer);
  const doc = await getDocument({ data, useSystemFonts: true }).promise;

  let orderNumber = null;
  let priceColX = null; // "Unit Price" column x (both formats)
  let discColX = null; // "Disc.%" column x (KG format only)
  const items = new Map(); // sku -> qty (real product lines)
  const generic = new Map(); // sku -> qty (GEN* lines, not reconciled)
  const pricing = new Map(); // sku -> { price, disc }

  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p);
    const tc = await page.getTextContent();

    const lines = new Map(); // rounded y -> [{x, s}]
    for (const it of tc.items) {
      if (!it.str || !it.str.trim()) continue;
      const x = it.transform[4];
      const y = Math.round(it.transform[5]);
      if (!lines.has(y)) lines.set(y, []);
      lines.get(y).push({ x, s: it.str });
    }

    const rows = [...lines.values()].map((r) => r.sort((a, b) => a.x - b.x));

    // Header anchors (appear before the items; only on the first page).
    for (const parts of rows) {
      for (const q of parts) {
        const s = q.s.trim();
        if (priceColX == null && (s === 'Unit' || s === 'Price')) priceColX = q.x;
        // Header-driven, not client-specific: any OC (incl. Office Holdings)
        // gets DISCOUNT filled if it labels a discount column. Matches
        // "Disc", "Disc.", "Disc.%", "Disc %", "Discount", "Discount %"...
        if (discColX == null && /^disc(ount)?\.?\s*%?$/i.test(s)) discColX = q.x;
      }
    }

    for (const parts of rows) {
      if (!orderNumber) {
        const m = parts.map((q) => q.s).join(' ').match(/Order #\s*([A-Z0-9]+)/);
        if (m) orderNumber = m[1];
      }

      const skuTok = parts.find((q) => SKU_ANCHOR.test(q.s));
      if (!skuTok) continue;
      const sku = skuTok.s.match(SKU_ANCHOR)[1];

      // parts is x-sorted; the first 3-decimal token is the quantity
      // (price/amount come later and use 2 decimals).
      const qtyTok = parts.find((q) => QTY_TOKEN.test(q.s));
      if (!qtyTok) continue;
      const qty = Math.round(parseFloat(qtyTok.s.replace(/,/g, '')));
      if (!Number.isFinite(qty)) continue;

      const bucket = /^GEN/i.test(sku) ? generic : items;
      bucket.set(sku, (bucket.get(sku) ?? 0) + qty);

      if (!/^GEN/i.test(sku) && !pricing.has(sku)) {
        const price = moneyNear(parts, priceColX);
        const disc = moneyNear(parts, discColX);
        if (price != null || disc != null) pricing.set(sku, { price, disc });
      }
    }
  }

  return {
    orderNumber: orderNumber ?? fallbackLabel ?? 'OC',
    items, // Map<sku, qty>
    generic, // Map<sku, qty>
    pricing, // Map<sku, { price, disc }>
  };
}

// Merge per-SKU price/discount from every parsed OC into one lookup. First
// OC wins; a differing later value is flagged so the user can check.
export function ocPricingBySku(ocs) {
  const pricing = new Map();
  const conflicts = [];
  for (const oc of ocs) {
    for (const [sku, v] of oc.pricing) {
      const prev = pricing.get(sku);
      if (!prev) {
        pricing.set(sku, v);
      } else if (prev.price !== v.price || prev.disc !== v.disc) {
        conflicts.push(sku);
      }
    }
  }
  return { pricing, conflicts };
}

// Aggregate the Excel side: total qty per full SKU (base + integer size)
// across ALL warehouse blocks. Reuses parseKurtGeiger's warehouse items.
export function excelTotalsBySku(parsedExcel) {
  const total = new Map();
  for (const wh of parsedExcel.warehouses) {
    for (const it of wh.items) {
      total.set(it.sku, (total.get(it.sku) ?? 0) + it.qty);
    }
  }
  return total;
}

// Build the discrepancy report. One SKU is a discrepancy when the Excel total
// differs from the summed OC quantities, OR it appears on only one side.
// dif = Excel total - sum(OC quantities).
export function reconcile(excelTotals, ocs) {
  const columnLabels = ocs.map((o) => o.orderNumber);

  const skus = new Set();
  for (const s of excelTotals.keys()) skus.add(s);
  for (const oc of ocs) for (const s of oc.items.keys()) skus.add(s);

  const rows = [];
  let ok = 0;
  for (const sku of [...skus].sort()) {
    const cols = ocs.map((o) => (o.items.has(sku) ? o.items.get(sku) : null));
    const ocSum = cols.reduce((s, v) => s + (v ?? 0), 0);
    const reference = excelTotals.get(sku) ?? 0;
    const inAnyOc = cols.some((v) => v != null);
    if (!inAnyOc && reference === 0) continue;
    const dif = reference - ocSum;
    if (dif === 0 && inAnyOc) {
      ok++;
      continue;
    }
    // Classify so the sheet can say *why* each row is here, real
    // OC-vs-fichero mismatches first (rank 1), Excel-only noise last.
    let rank, motivo;
    if (!inAnyOc) {
      rank = 3;
      motivo = 'En fichero, sin OC';
    } else if (reference === 0) {
      rank = 2;
      motivo = `En OC, no repartido (faltan ${ocSum})`;
    } else {
      rank = 1;
      motivo =
        dif < 0
          ? `OC > Fichero (falta repartir ${-dif})`
          : `Fichero > OC (sobran ${dif})`;
    }
    rows.push({ sku, cols, reference, dif, motivo, rank });
  }
  rows.sort((a, b) => a.rank - b.rank || a.sku.localeCompare(b.sku));

  // GEN* lines: surfaced but never reconciled.
  const genericSkus = new Set();
  for (const o of ocs) for (const s of o.generic.keys()) genericSkus.add(s);
  const generic = [...genericSkus].sort().map((sku) => ({
    sku,
    perCol: ocs.map((o) => (o.generic.has(sku) ? o.generic.get(sku) : null)),
  }));

  return {
    rows,
    ok,
    generic,
    columnLabels,
    referenceLabel: 'Fichero',
    title: 'DISCREPANCIAS — Conciliación OC vs fichero',
    legend:
      'Diferencia = Fichero − (suma de OC).  Negativa: la OC pide más de lo repartido.  Positiva: el fichero reparte de más.',
    summaryLabels: {
      real: 'Reales OC≠fichero',
      refOnly: 'En OC no repartido',
      colOnly: 'En fichero sin OC',
    },
  };
}
