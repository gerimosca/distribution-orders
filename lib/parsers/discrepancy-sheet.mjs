// Shared "Discrepancias" sheet builder, used by every client's reconciliation.
//
// d = {
//   rows: [{ sku, cols:[n|null,...], reference, dif, motivo, rank }],
//   ok: number,
//   columnLabels: string[],          // per-column headers (OC#s or PO#s)
//   referenceLabel?: string,         // single reference column header
//   title?: string,
//   legend?: string,
//   summaryLabels?: { real, refOnly, colOnly },
//   generic?: [{ sku, perCol:[n|null,...] }],
//   genericLabel?: string,
// }
//
// rank: 1 = real mismatch (present both sides, differ), 2 = only on the
// reference side, 3 = only on the columns side. Rows arrive pre-sorted.
export function buildDiscrepancySheet(d) {
  const columnLabels = d.columnLabels ?? [];
  const referenceLabel = d.referenceLabel ?? 'Fichero';
  const title = d.title ?? 'DISCREPANCIAS — Conciliación';
  const s = d.summaryLabels ?? {
    real: 'Reales (no cuadran)',
    refOnly: `Solo en ${referenceLabel}`,
    colOnly: 'Solo en columnas',
  };

  const real = d.rows.filter((r) => r.rank === 1).length;
  const refOnly = d.rows.filter((r) => r.rank === 2).length;
  const colOnly = d.rows.filter((r) => r.rank === 3).length;

  const aoa = [];
  aoa.push([title]);
  aoa.push([`SKU que cuadran: ${d.ok}`, `Discrepancias: ${d.rows.length}`]);
  aoa.push([
    `${s.real}: ${real}`,
    `${s.refOnly}: ${refOnly}`,
    `${s.colOnly}: ${colOnly}`,
  ]);
  if (d.legend) aoa.push([d.legend]);
  aoa.push([]);

  if (d.rows.length === 0) {
    aoa.push(['Sin discrepancias: todo cuadra ✔']);
  } else {
    aoa.push(['SKU', ...columnLabels, referenceLabel, 'Diferencia', 'Motivo']);
    for (const row of d.rows) {
      aoa.push([
        row.sku,
        ...row.cols.map((v) => (v == null ? '--' : v)),
        row.reference === 0 ? '--' : row.reference,
        row.dif,
        row.motivo,
      ]);
    }
  }

  if (d.generic && d.generic.length) {
    aoa.push([]);
    aoa.push([d.genericLabel ?? 'Líneas genéricas (no conciliadas):']);
    aoa.push(['SKU', ...columnLabels]);
    for (const g of d.generic) {
      aoa.push([g.sku, ...g.perCol.map((v) => (v == null ? '--' : v))]);
    }
  }
  return aoa;
}
