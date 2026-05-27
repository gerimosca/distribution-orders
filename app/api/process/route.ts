import { NextResponse } from 'next/server';
import * as XLSX from 'xlsx';
import * as parsers from '@/lib/parsers/index.mjs';
import { createZip, safeFileName } from '@/lib/zip.mjs';

export const runtime = 'nodejs';
// PDF parsing is sequential and CPU-bound; default 10s (Hobby) cuts off mid-
// batch when several PDFs are uploaded. 60s covers realistic batches without
// changing tier behavior (Vercel clamps to the plan's max).
export const maxDuration = 60;

// The parsers are untyped .mjs; treat the boundary as dynamic.
const getClient = (parsers as { getClient: (id: string) => any }).getClient;

export async function POST(req: Request) {
  try {
    const form = await req.formData();
    const clientId = form.get('client');
    if (typeof clientId !== 'string' || !clientId) {
      return new NextResponse('Falta el campo "client"', { status: 400 });
    }
    const client = getClient(clientId);
    if (!client) {
      return new NextResponse(`Cliente desconocido: ${clientId}`, { status: 400 });
    }

    const ocFiles = form
      .getAll('ocFiles')
      .filter((f): f is File => f instanceof File);

    // Output format: the default Excel workbook, or a ZIP bundling each tab as
    // its own .csv (requested via the "Descargar CSVs (ZIP)" button).
    const output = form.get('output') === 'zip' ? 'zip' : 'xlsx';

    let parsed: any;
    let downloadBase: string;

    if (client.inputKind === 'pdfs') {
      // Office Holdings: several PO PDFs are the main input.
      const poFiles = form
        .getAll('poFiles')
        .filter((f): f is File => f instanceof File);
      if (!poFiles.length) {
        return new NextResponse('Sube al menos un PDF de PO', { status: 400 });
      }
      const pos = [];
      for (const f of poFiles) {
        const buf = Buffer.from(await f.arrayBuffer());
        pos.push(await client.parsePo(buf, f.name.replace(/\.[^.]+$/, '')));
      }
      parsed = { pos };
      if (ocFiles.length && client.parseOc) {
        const ocs = [];
        for (const oc of ocFiles) {
          const buf = Buffer.from(await oc.arrayBuffer());
          ocs.push(await client.parseOc(buf, oc.name.replace(/\.[^.]+$/, '')));
        }
        parsed.discrepancies = client.reconcile(ocs, pos);
        if (client.ocPricingBySku) parsed.ocPricing = client.ocPricingBySku(ocs).pricing;
        if (client.allocateOcsToPos) parsed.ocByPo = client.allocateOcsToPos(ocs, pos);
      }
      downloadBase = `${client.name} - POs`;
    } else {
      // Kurt Geiger: one client Excel + optional OC PDFs.
      const file = form.get('file');
      if (!(file instanceof File)) {
        return new NextResponse('Falta el fichero', { status: 400 });
      }
      const buffer = Buffer.from(await file.arrayBuffer());
      parsed = client.parse(buffer);
      if (ocFiles.length && client.parseOc) {
        const ocs = [];
        for (const oc of ocFiles) {
          const buf = Buffer.from(await oc.arrayBuffer());
          ocs.push(await client.parseOc(buf, oc.name.replace(/\.[^.]+$/, '')));
        }
        if (client.reconcileMode === 'per-po') {
          // John Lewis: the Excel is already split per PO, so reconcile the
          // parsed POs against the OC (per-PO columns).
          parsed.discrepancies = client.reconcile(ocs, parsed.pos);
        } else {
          // Kurt Geiger: reconcile a single Excel total per SKU against the OC.
          const excelTotals = client.excelTotalsBySku(parsed);
          parsed.discrepancies = client.reconcile(excelTotals, ocs);
        }
        if (client.ocPricingBySku) parsed.ocPricing = client.ocPricingBySku(ocs).pricing;
      }
      downloadBase = file.name.replace(/\.[^.]+$/, '');
    }

    const wb = client.build(parsed);

    const discrepancyCount = parsed.discrepancies
      ? parsed.discrepancies.rows.length
      : 0;
    const warnings =
      (parsed.warehouses ?? parsed.pos ?? []).reduce(
        (s: number, w: { warnings?: string[] }) => s + (w.warnings?.length ?? 0),
        0
      ) + (parsed.globalIssues?.length ?? 0);

    const commonHeaders = {
      'x-warnings': String(warnings),
      'x-discrepancies': String(discrepancyCount),
    };

    if (output === 'zip') {
      // One CSV per workbook tab. A UTF-8 BOM makes Excel open the CSV with the
      // right encoding (accents, ✔, ≠) instead of mojibake.
      const files = wb.SheetNames.map((name: string) => ({
        name: `${safeFileName(name)}.csv`,
        data: '﻿' + XLSX.utils.sheet_to_csv(wb.Sheets[name]),
      }));
      const zip = createZip(files);
      return new NextResponse(zip, {
        status: 200,
        headers: {
          ...commonHeaders,
          'content-type': 'application/zip',
          'content-disposition': `attachment; filename="${downloadBase} - distribución (CSV).zip"`,
        },
      });
    }

    const out = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    return new NextResponse(out, {
      status: 200,
      headers: {
        ...commonHeaders,
        'content-type':
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'content-disposition': `attachment; filename="${downloadBase} - distribución.xlsx"`,
      },
    });
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    return new NextResponse(`Error procesando: ${msg}`, { status: 500 });
  }
}
