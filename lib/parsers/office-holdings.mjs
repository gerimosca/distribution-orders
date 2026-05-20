import * as XLSX from 'xlsx';
import { getDocument } from './pdfjs.mjs';
import { buildDiscrepancySheet } from './discrepancy-sheet.mjs';

// Office Holdings purchase-order PDF. One PDF = one PO.
// - PO number: header "ORDER NO: POxxxxxx".
// - SKU base: the FACTORY STYLE column (e.g. S100917-01).
// - Sizes: UK sizes in the column header, converted to EU.
// - Quantity: per size column, on the FACTORY STYLE row.
const UK_TO_EU = { '3.5': 36, '4': 37, '5': 38, '6': 39, '7': 40, '8': 41 };
// Office Holdings POs print the size header twice, in two adjacent y-rows:
// the UK label ("3.5 4 5 6 7 8") and a 3-digit pre-pack code ("035 040 050
// 060 070 080"). The pre-pack code is just the UK size × 10 zero-padded
// (035 = 3.5, 040 = 4, ...). On some runtimes (e.g. Vercel's Lambda without
// the full font set) pdfjs can drop the UK label line — fall back to the
// pre-pack codes so we still recover the column → EU mapping.
const PREPACK_TO_EU = { '035': 36, '040': 37, '050': 38, '060': 39, '070': 40, '080': 41 };
// Specific enough on its own; x-range checks made the parser fragile to any
// PDF coordinate-system shift between runtimes.
const FACTORY_STYLE = /^S\d{3,}-\d{1,3}[A-Z]?$/;
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

// Standard EU run for an Office Holdings ladies PO: 6 contiguous sizes
// 36..41. Used as a last-resort fallback when both header rows are absent
// from the textContent (some runtimes drop the size header glyphs).
const DEFAULT_EU_RUN = [36, 37, 38, 39, 40, 41];

// Find the factory style on a row, even when pdfjs splits "S100917-06" into
// multiple tokens (which can happen on Linux/serverless runtimes where the
// embedded font is treated differently than on Windows). Strategy:
//   1. Fast path — per-token regex match.
//   2. Slow path — pick any token starting with "S", concatenate the next few
//      tokens in x-order, look for the factory-style pattern as a prefix.
// Returns { x, s } where x is the leftmost token's x and s is the rebuilt
// SKU base ("S100917-06"), or null.
function findFactoryStyleAnchor(parts) {
  for (const q of parts) {
    if (FACTORY_STYLE.test(q.s)) return { x: q.x, s: q.s };
  }
  const sorted = [...parts].sort((a, b) => a.x - b.x);
  for (let i = 0; i < sorted.length; i++) {
    if (!sorted[i].s.startsWith('S')) continue;
    let joined = sorted[i].s;
    for (let j = i + 1; j < Math.min(sorted.length, i + 20); j++) {
      joined += sorted[j].s;
      const m = joined.match(/^(S\d{3,}-\d{1,3}[A-Z]?)(?![\dA-Z-])/);
      if (m) return { x: sorted[i].x, s: m[1] };
      // Stop early if the accumulator can no longer extend into a valid
      // factory style (saves work on long unrelated runs).
      if (!/^S\d*-?\d*[A-Z]?$/.test(joined)) break;
    }
  }
  return null;
}

// Infer the 6 size column x-positions by clustering integer tokens across
// every item row. Each item row contributes its non-zero quantities; the
// 6 most-frequent x-positions are the size columns. Robust to:
//   - missing size header row (UK + pre-pack both absent)
//   - rows where some sizes have qty=0 (those cells emit no token)
//   - shifted x coordinates between runtimes (we read positions, not match
//     them against hardcoded numbers)
function inferSizeColsFromItemRows(itemAnchors) {
  if (!itemAnchors.length) return null;
  const xCounts = new Map(); // rounded x → count
  for (const { parts, fs } of itemAnchors) {
    const ints = parts
      .filter((q) => q.x > fs.x && /^\d+$/.test(q.s))
      .sort((a, b) => a.x - b.x);
    if (ints.length < 2) continue;
    // Drop a trailing row-total if its gap is >3x the average inner gap.
    if (ints.length >= 3) {
      const last = ints[ints.length - 1];
      const prev = ints[ints.length - 2];
      const spans = [];
      for (let i = 1; i < ints.length - 1; i++) spans.push(ints[i].x - ints[i - 1].x);
      const avg = spans.reduce((s, v) => s + v, 0) / spans.length;
      if (last.x - prev.x > avg * 3) ints.pop();
    }
    for (const q of ints) {
      const rx = Math.round(q.x / 4) * 4; // 4-unit bucket absorbs jitter
      xCounts.set(rx, (xCounts.get(rx) ?? 0) + 1);
    }
  }
  if (xCounts.size < DEFAULT_EU_RUN.length) return null;
  // Pick the top-6 most-frequent x positions, then sort left-to-right.
  const top = [...xCounts.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, DEFAULT_EU_RUN.length)
    .map(([x]) => x)
    .sort((a, b) => a - b);
  return top.map((x, i) => ({ x, eu: DEFAULT_EU_RUN[i] }));
}

export async function parsePoPdf(buffer, fallbackLabel) {
  const data =
    buffer instanceof ArrayBuffer ? new Uint8Array(buffer) : Uint8Array.from(buffer);
  const doc = await getDocument({ data, useSystemFonts: true }).promise;

  let poNumber = null;
  let sizeCols = null; // [{ x, eu }]
  let inferredSizeCols = false;
  let grandTotal = null;
  const agg = new Map(); // sku -> qty
  let itemsSum = 0;

  // Collect every line across every page first so the size header can be
  // discovered regardless of where pdfjs places it in document order.
  const allRows = [];
  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p);
    const tc = await page.getTextContent();
    for (const parts of linesByY(tc.items).values()) {
      allRows.push(parts);
    }
  }

  // First pass: header metadata + size column detection + item-row anchors.
  const itemAnchors = []; // [{ parts, fs: { x, s } }]
  for (const parts of allRows) {
    const joined = parts.map((q) => q.s).join(' ');

    if (!poNumber) {
      const m = joined.match(/ORDER NO:\s*([A-Z0-9]+)/i);
      if (m) poNumber = m[1];
    }
    if (grandTotal == null && /GRAND TOTAL:/i.test(joined)) {
      const nums = parts.filter((q) => /^\d+$/.test(q.s));
      if (nums.length) grandTotal = Number(nums[nums.length - 1].s);
    }

    if (!sizeCols) {
      const ukHits = parts.filter((q) => Object.hasOwn(UK_TO_EU, q.s));
      if (ukHits.length >= 4) {
        sizeCols = ukHits.map((q) => ({ x: q.x, eu: UK_TO_EU[q.s] }));
      } else {
        const prepackHits = parts.filter((q) =>
          Object.hasOwn(PREPACK_TO_EU, q.s)
        );
        if (prepackHits.length >= 4) {
          sizeCols = prepackHits.map((q) => ({
            x: q.x,
            eu: PREPACK_TO_EU[q.s],
          }));
        }
      }
    }

    const fs = findFactoryStyleAnchor(parts);
    if (fs) itemAnchors.push({ parts, fs });
  }

  // Last resort: cluster qty x-positions from every item row to derive the
  // 6 size columns. Maps to the ladies template's standard EU 36..41 run.
  if (!sizeCols) {
    sizeCols = inferSizeColsFromItemRows(itemAnchors);
    if (sizeCols) inferredSizeCols = true;
  }

  // Second pass: extract quantities per (item row, size column).
  if (sizeCols) {
    for (const { parts, fs } of itemAnchors) {
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
    let diag;
    if (itemAnchors.length) {
      const fsSample = itemAnchors
        .slice(0, 3)
        .map(({ fs, parts }) => {
          const ints = parts
            .filter((q) => q.x > fs.x && /^\d+$/.test(q.s))
            .map((q) => `${q.s}@${Math.round(q.x)}`)
            .join(',');
          return `${fs.s}@x=${Math.round(fs.x)} ints=[${ints || '∅'}]`;
        })
        .join(' || ');
      diag = `factory styles=${itemAnchors.length}; ${fsSample}`;
    } else {
      // Surface document-wide candidates instead of just the first 6 rows
      // (which on Office Holdings PDFs is the invoicing address — useless
      // for diagnosing whether the SKUs are present further down).
      const flat = allRows.flat();
      const sLike = flat
        .filter((q) => /^S\d/.test(q.s) || /S\d{2,}/.test(q.s))
        .slice(0, 8)
        .map((q) => `${q.s}@${Math.round(q.x)}`);
      const dashed = flat
        .filter((q) => /\d-\d/.test(q.s))
        .slice(0, 6)
        .map((q) => q.s);
      const intRows = allRows
        .filter((parts) => parts.filter((q) => /^\d+$/.test(q.s)).length >= 3)
        .slice(0, 3)
        .map((parts) =>
          parts
            .slice(0, 12)
            .map((q) => `${q.s}@${Math.round(q.x)}`)
            .join(' ')
        );
      diag =
        `páginas=${doc.numPages}, filas=${allRows.length}, tokens=${flat.length}. ` +
        `Tokens con "S\\d": [${sLike.join(' | ') || '∅'}]. ` +
        `Tokens con guión+dígitos: [${dashed.join(' | ') || '∅'}]. ` +
        `Filas con ≥3 enteros: [${intRows.join(' || ') || '∅'}]`;
    }
    warnings.push(
      'No se detectó la fila de tallas (ni "3.5 4 5 6 7 8" ni "035 040 050 060 070 080") ' +
        `y no se pudo inferir de las cantidades. Diagnóstico: ${diag}.`
    );
  } else if (inferredSizeCols) {
    warnings.push(
      'Tallas inferidas (36-41) por la fila de cantidades — la cabecera de tallas no se leyó. ' +
        'Verificar que el PO sea formato Office Holdings señoras estándar.'
    );
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
  const usedTabs = new Set();

  if (parsed.discrepancies) {
    const ws = XLSX.utils.aoa_to_sheet(buildDiscrepancySheet(parsed.discrepancies));
    XLSX.utils.book_append_sheet(wb, ws, 'Discrepancias');
    usedTabs.add('Discrepancias');
  }

  const pricing = parsed.ocPricing; // Map<sku,{price,disc}> | undefined
  const priceOf = (sku) => {
    const v = pricing && pricing.get(sku);
    return [v && v.price != null ? v.price : '', v && v.disc != null ? v.disc : ''];
  };

  // uniquePoLabels already disambiguates duplicate PO numbers; usedTabs is
  // belt-and-braces in case a label collides with "Discrepancias".
  const tabLabels = uniquePoLabels(parsed.pos).map((label) => {
    if (!usedTabs.has(label)) {
      usedTabs.add(label);
      return label;
    }
    for (let n = 2; ; n++) {
      const suffix = ` (${n})`;
      const candidate = label.slice(0, 31 - suffix.length) + suffix;
      if (!usedTabs.has(candidate)) {
        usedTabs.add(candidate);
        return candidate;
      }
    }
  });
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
