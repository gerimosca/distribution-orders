import * as XLSX from 'xlsx';
import { buildDiscrepancySheet } from './discrepancy-sheet.mjs';

const KNOWN_UK = new Set([
  '001 UK Warehouse',
  'Liberty Regent St 025',
  'Selfridges Manchester 223',
  'Selfridges Manchester City Ladies 241',
  'Selfridges Birmingham Ladies 253',
  '911 Photography Unit',
]);

const KNOWN_VENLO = new Set([
  '122 Venlo Warehouse',
  'BT Dublin Ladies 457',
  'BT Cork 451',
  'BT Limerick 452',
  'BT Galway 453',
  'BT2 Dundrum 455',
  'BT2 Blanchardstown 456',
  'BT Online 499',
  'Arnotts Dublin 354',
]);

const IGNORE_WAREHOUSES = new Set([
  'UK Total',
  'Venlo Total',
  'Overall Total',
  'UK TOTAL',
  'VENLO TOTAL',
  'OVERALL TOTAL',
]);

const TAB_NAMES = {
  '001 UK Warehouse': '001 UK Warehouse',
  '122 Venlo Warehouse': '122 Venlo Warehouse',
  'Liberty Regent St 025': '025 Liberty Regent St',
  'Selfridges Manchester 223': '223 Selfridges Manchester',
  'Selfridges Manchester City Ladies 241': '241 Selfridges MCL',
  'Selfridges Birmingham Ladies 253': '253 Selfridges B Ladies',
  'BT Dublin Ladies 457': '457 BT Dublin Ladies',
  'BT Cork 451': '451 BT Cork',
  'BT Limerick 452': '452 BT Limerick',
  'BT Galway 453': '453 BT Galway',
  'BT2 Dundrum 455': '455 BT2 Dundrum',
  'BT2 Blanchardstown 456': '456 BT2 Blanchardstown',
  'BT Online 499': '499 BT Online',
  'Arnotts Dublin 354': '354 Arnotts Dublin',
  '911 Photography Unit': '911 Photography Unit',
};

function tabNameFor(warehouse) {
  if (TAB_NAMES[warehouse]) return TAB_NAMES[warehouse];
  const c = warehouseCode(warehouse);
  if (c && TAB_NAMES_BY_CODE[c]) return TAB_NAMES_BY_CODE[c];
  return warehouse.slice(0, 31);
}

const BLOCK_WIDTH = 20;

function normalizeName(s) {
  return String(s).replace(/\s+/g, ' ').trim();
}

// Block headers in newer exports carry a "SIZE CURVE" suffix and CR/LF; the
// warehouse name is the part before it.
function normalizeWarehouseName(s) {
  return normalizeName(s).replace(/\s*SIZE CURVE\s*$/i, '').trim();
}

// Each store has a stable 2-4 digit code. Exports differ only in whether the
// code leads ("025 Liberty Regent St") or trails ("Liberty Regent St 025"),
// so classify by code, order-independent. The lookarounds skip codes glued
// into a token like "BT2".
function warehouseCode(name) {
  const m = String(name).match(/(?<![A-Za-z0-9])(\d{2,4})(?![A-Za-z0-9])/);
  return m ? m[1] : null;
}

const KNOWN_UK_CODES = new Set(
  [...KNOWN_UK].map(warehouseCode).filter(Boolean)
);
const KNOWN_VENLO_CODES = new Set(
  [...KNOWN_VENLO].map(warehouseCode).filter(Boolean)
);

const TAB_NAMES_BY_CODE = {};
for (const [k, v] of Object.entries(TAB_NAMES)) {
  const c = warehouseCode(k);
  if (c) TAB_NAMES_BY_CODE[c] = v;
}

// Layout drifts between exports (column/row offsets shift), so detect it
// instead of hardcoding indices. Anchors: the header row is the one holding
// "Full Supplier Style Ref"; warehouse blocks are the columns that carry a
// numeric size in the header row and a block name in the row(s) just above.
function detectLayout(rows) {
  let headerIdx = -1;
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    if (row && row.some((v) => normalizeName(v) === 'Full Supplier Style Ref')) {
      headerIdx = r;
      break;
    }
  }
  if (headerIdx < 0) {
    throw new Error(
      'No reconozco el formato: falta la columna "Full Supplier Style Ref".'
    );
  }
  const header = rows[headerIdx];
  const colByName = (name) =>
    header.findIndex((v) => normalizeName(v) === name);

  const skuCol = colByName('Full Supplier Style Ref');
  const ukPoCol = colByName('UK PO Number');
  const venloPoCol = colByName('Venlo PO Number');
  const totalQtyCol = colByName('Total Qty');

  // Candidate block columns: numeric size in the header row, past the SKU col.
  const sizeRow = header;
  const numericCols = [];
  for (let c = skuCol + 1; c < sizeRow.length; c++) {
    if (typeof sizeRow[c] === 'number') numericCols.push(c);
  }

  // Warehouse-name row = the row above the header with the most string cells
  // sitting on those numeric columns (the gap to the header varies by export).
  let nameRowIdx = -1;
  let bestScore = -1;
  for (let r = headerIdx - 1; r >= Math.max(0, headerIdx - 4); r--) {
    const row = rows[r];
    if (!row) continue;
    let score = 0;
    for (const c of numericCols) {
      if (typeof row[c] === 'string' && row[c].trim()) score++;
    }
    if (score > bestScore) {
      bestScore = score;
      nameRowIdx = r;
    }
  }
  const nameRow = rows[nameRowIdx] ?? [];

  // Per-row subtotal columns (UK / Venlo totals) live in the header row,
  // before the block region, as plain strings.
  const firstBlockCol = numericCols.find(
    (c) => typeof nameRow[c] === 'string' && nameRow[c].trim()
  );
  const subtotalCol = (label) => {
    const idx = header.findIndex(
      (v, i) => i < (firstBlockCol ?? Infinity) && normalizeName(v) === label
    );
    return idx;
  };

  return {
    headerIdx,
    skuCol,
    ukPoCol,
    venloPoCol,
    totalQtyCol,
    ukTotalCol: subtotalCol('001 UK Warehouse'),
    venloTotalCol: subtotalCol('122 Venlo Warehouse'),
    nameRowIdx,
    sizeRowIdx: headerIdx,
    dataStartIdx: headerIdx + 1,
    firstBlockCol: firstBlockCol ?? -1,
  };
}

function findWarehouseBlocks(nameRow, sizeRow, firstBlockCol) {
  const blocks = [];
  for (let c = firstBlockCol; c < nameRow.length; c++) {
    const v = nameRow[c];
    // Real block headers are text; numeric stray cells (0, 911) are not.
    if (typeof v !== 'string') continue;
    const name = normalizeWarehouseName(v);
    if (!name) continue;
    if (!/[a-z]/i.test(name)) continue;
    if (IGNORE_WAREHOUSES.has(name)) continue;
    if (typeof sizeRow[c] !== 'number') continue;
    blocks.push({ name, startCol: c });
  }
  return blocks;
}

function readSizesForBlock(sizeRow, startCol) {
  const sizes = [];
  for (let i = 0; i < BLOCK_WIDTH - 1; i++) {
    const v = sizeRow[startCol + i];
    if (v == null || v === '') {
      sizes.push(null);
    } else {
      sizes.push(Number(v));
    }
  }
  return sizes;
}

function isEmpty(v) {
  return v == null || v === '' || v === 0;
}

function classifyWarehouses(blocks, sizesByBlock, rows, L) {
  const signals = new Map();
  for (const b of blocks) signals.set(b.name, { uk: 0, venlo: 0, both: 0 });

  for (let r = L.dataStartIdx; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    if (row[L.skuCol] == null || row[L.skuCol] === '') continue;
    const hasUkPo = !isEmpty(row[L.ukPoCol]);
    const hasVenloPo = !isEmpty(row[L.venloPoCol]);
    if (!hasUkPo && !hasVenloPo) continue;

    for (let bi = 0; bi < blocks.length; bi++) {
      const b = blocks[bi];
      const sizes = sizesByBlock[bi];
      let hasQty = false;
      for (let i = 0; i < sizes.length; i++) {
        const qty = row[b.startCol + i];
        if (typeof qty === 'number' && qty > 0) {
          hasQty = true;
          break;
        }
      }
      if (!hasQty) continue;
      const s = signals.get(b.name);
      if (hasUkPo && !hasVenloPo) s.uk += 1;
      else if (hasVenloPo && !hasUkPo) s.venlo += 1;
      else s.both += 1;
    }
  }

  const classification = new Map();
  for (const b of blocks) {
    const s = signals.get(b.name);
    let kind;
    if (s.uk > 0 && s.venlo === 0) kind = 'UK';
    else if (s.venlo > 0 && s.uk === 0) kind = 'Venlo';
    else if (s.uk === 0 && s.venlo === 0) kind = 'unknown';
    else kind = 'conflict';
    classification.set(b.name, { kind, signals: s });
  }
  return classification;
}

export function parseKurtGeiger(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, {
    header: 1,
    defval: null,
    blankrows: true,
  });

  const L = detectLayout(rows);

  const blocks = findWarehouseBlocks(
    rows[L.nameRowIdx],
    rows[L.sizeRowIdx],
    L.firstBlockCol
  );
  const sizesByBlock = blocks.map((b) =>
    readSizesForBlock(rows[L.sizeRowIdx], b.startCol)
  );

  const classification = classifyWarehouses(blocks, sizesByBlock, rows, L);

  const warehouseData = new Map();
  for (const b of blocks) {
    warehouseData.set(b.name, { items: [], pos: new Set(), warnings: [] });
  }

  for (const b of blocks) {
    const c = classification.get(b.name);
    const warns = warehouseData.get(b.name).warnings;
    const code = warehouseCode(b.name);
    const expectedUk =
      KNOWN_UK.has(b.name) || (code != null && KNOWN_UK_CODES.has(code));
    const expectedVenlo =
      KNOWN_VENLO.has(b.name) || (code != null && KNOWN_VENLO_CODES.has(code));
    const known = expectedUk || expectedVenlo;

    if (c.kind === 'unknown' && known) {
      c.kind = expectedUk ? 'UK' : 'Venlo';
      classification.set(b.name, c);
    } else if (c.kind === 'unknown') {
      warns.push(
        `Almacén "${b.name}" no clasificable (todas las filas con qty tienen ambas POs) y no está en lista conocida. PO en blanco.`
      );
    } else if (c.kind === 'conflict') {
      warns.push(
        `Almacén "${b.name}": señales mezcladas UK=${c.signals.uk} filas, Venlo=${c.signals.venlo} filas. PO en blanco.`
      );
    } else if (known && expectedUk && c.kind === 'Venlo') {
      warns.push(
        `Conflicto: "${b.name}" está en lista UK pero el fichero lo trata como Venlo (${c.signals.venlo} filas con solo Venlo PO).`
      );
    } else if (known && expectedVenlo && c.kind === 'UK') {
      warns.push(
        `Conflicto: "${b.name}" está en lista Venlo pero el fichero lo trata como UK (${c.signals.uk} filas con solo UK PO).`
      );
    } else if (!known) {
      warns.push(
        `Almacén NUEVO "${b.name}" clasificado como ${c.kind} por inferencia (UK=${c.signals.uk}, Venlo=${c.signals.venlo} filas). Verificar.`
      );
    }
  }

  const globalIssues = [];
  const totalChecks = [];
  const ukChecks = [];
  const venloChecks = [];

  for (let r = L.dataStartIdx; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const baseSku = row[L.skuCol];
    if (baseSku == null || baseSku === '') continue;

    const ukPo = row[L.ukPoCol];
    const venloPo = row[L.venloPoCol];

    let hasAnyQty = false;
    for (const b of blocks) {
      for (let i = 0; i < BLOCK_WIDTH - 1; i++) {
        const v = row[b.startCol + i];
        if (typeof v === 'number' && v > 0) {
          hasAnyQty = true;
          break;
        }
      }
      if (hasAnyQty) break;
    }
    if (!hasAnyQty && isEmpty(ukPo) && isEmpty(venloPo)) continue;

    let rowAllSizesSum = 0;
    let rowUkInt = 0;
    let rowVenloInt = 0;

    for (let bi = 0; bi < blocks.length; bi++) {
      const b = blocks[bi];
      const sizes = sizesByBlock[bi];
      const kind = classification.get(b.name).kind;
      const po = kind === 'UK' ? ukPo : kind === 'Venlo' ? venloPo : null;

      let warehouseHasQty = false;
      for (let i = 0; i < sizes.length; i++) {
        const size = sizes[i];
        const qty = row[b.startCol + i];
        if (size == null) continue;
        if (typeof qty !== 'number') continue;
        rowAllSizesSum += qty;
        if (!Number.isInteger(size)) {
          if (qty > 0) {
            const msg = `${baseSku}: talla ${size} tiene qty=${qty} (se ignoró)`;
            warehouseData.get(b.name).warnings.push(msg);
          }
          continue;
        }
        if (b.name === '001 UK Warehouse') rowUkInt += qty;
        if (b.name === '122 Venlo Warehouse') rowVenloInt += qty;
        if (qty <= 0) continue;
        warehouseHasQty = true;
        const sku = `${baseSku}${size}`;
        warehouseData.get(b.name).items.push({ sku, qty });
      }

      if (warehouseHasQty && po != null && po !== '') {
        warehouseData.get(b.name).pos.add(String(po));
      }
    }

    const totalQty =
      L.totalQtyCol >= 0 && typeof row[L.totalQtyCol] === 'number'
        ? row[L.totalQtyCol]
        : null;
    const ukTotal =
      L.ukTotalCol >= 0 && typeof row[L.ukTotalCol] === 'number'
        ? row[L.ukTotalCol]
        : null;
    const venloTotal =
      L.venloTotalCol >= 0 && typeof row[L.venloTotalCol] === 'number'
        ? row[L.venloTotalCol]
        : null;
    if (totalQty !== null) {
      totalChecks.push({ r, baseSku, expected: totalQty, detail: rowAllSizesSum });
    }
    if (ukTotal !== null) {
      ukChecks.push({ baseSku, expected: ukTotal, detail: rowUkInt });
    }
    if (venloTotal !== null) {
      venloChecks.push({ baseSku, expected: venloTotal, detail: rowVenloInt });
    }
  }

  // A total/subtotal column is only trustworthy if it equals the computed
  // detail on at least one row; newer exports put a size-curve code in those
  // columns instead of a quantity, so skip the check rather than flood false
  // warnings.
  const columnIsTotal = (arr) => arr.some((x) => x.expected === x.detail);
  if (columnIsTotal(totalChecks)) {
    for (const x of totalChecks) {
      if (x.expected !== x.detail) {
        globalIssues.push(
          `Fila ${x.r + 1} (${x.baseSku}): Total Qty=${x.expected} ≠ suma total detalle=${x.detail}`
        );
      }
    }
  }
  if (warehouseData.has('001 UK Warehouse') && columnIsTotal(ukChecks)) {
    for (const x of ukChecks) {
      if (x.expected !== x.detail) {
        warehouseData
          .get('001 UK Warehouse')
          .warnings.push(
            `${x.baseSku}: subtotal en fichero=${x.expected}, detalle suma=${x.detail}`
          );
      }
    }
  }
  if (warehouseData.has('122 Venlo Warehouse') && columnIsTotal(venloChecks)) {
    for (const x of venloChecks) {
      if (x.expected !== x.detail) {
        warehouseData
          .get('122 Venlo Warehouse')
          .warnings.push(
            `${x.baseSku}: subtotal en fichero=${x.expected}, detalle suma=${x.detail}`
          );
      }
    }
  }

  const warehouses = [];
  let grandTotal = 0;
  for (const b of blocks) {
    const d = warehouseData.get(b.name);
    const totalQty = d.items.reduce((s, it) => s + it.qty, 0);
    grandTotal += totalQty;
    warehouses.push({
      warehouse: b.name,
      kind: classification.get(b.name).kind,
      pos: [...d.pos],
      items: d.items,
      warnings: d.warnings,
      totalQty,
    });
  }

  return { warehouses, globalIssues, grandTotal };
}

export function buildWorkbook(parsed) {
  const wb = XLSX.utils.book_new();

  // Optional reconciliation sheet first, so it is the first tab the user sees.
  if (parsed.discrepancies) {
    const aoa = buildDiscrepancySheet(parsed.discrepancies);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    XLSX.utils.book_append_sheet(wb, ws, 'Discrepancias');
  }

  const pricing = parsed.ocPricing; // Map<sku,{price,disc}> | undefined
  const priceOf = (sku) => {
    const v = pricing && pricing.get(sku);
    return [v && v.price != null ? v.price : '', v && v.disc != null ? v.disc : ''];
  };

  for (const wh of parsed.warehouses) {
    const aoa = [];
    aoa.push(['SKU', 'QUANTITY', 'PRICE', 'DISCOUNT', null, 'PO Numbers:', wh.pos.join(', ')]);
    const warningText = wh.warnings.length
      ? 'Revisar: ' + wh.warnings.join(' | ')
      : null;
    const agg = new Map();
    for (const it of wh.items) {
      agg.set(it.sku, (agg.get(it.sku) ?? 0) + it.qty);
    }
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
    XLSX.utils.book_append_sheet(wb, ws, tabNameFor(wh.warehouse));
  }
  return wb;
}
