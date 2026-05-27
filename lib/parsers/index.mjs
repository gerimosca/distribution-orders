import { parseKurtGeiger, buildWorkbook as buildKurtGeigerWorkbook } from './kurt-geiger.mjs';
import { parseOcPdf, excelTotalsBySku, reconcile, ocPricingBySku } from './oc-pdf.mjs';
import {
  parsePoPdf,
  build as buildOfficeHoldings,
  reconcile as reconcileOffice,
  allocateOcsToPos as allocateOcsToPosOffice,
} from './office-holdings.mjs';
import {
  parseJohnLewis,
  build as buildJohnLewis,
  reconcile as reconcileJohnLewis,
} from './john-lewis.mjs';

export const CLIENTS = [
  {
    id: 'kurt-geiger',
    name: 'Kurt Geiger',
    // Main input: one client Excel. Optional OC PDFs reconcile against it.
    inputKind: 'excel',
    parse: parseKurtGeiger,
    build: buildKurtGeigerWorkbook,
    parseOc: parseOcPdf,
    excelTotalsBySku,
    reconcile,
    ocPricingBySku,
  },
  {
    id: 'office-holdings',
    name: 'Office Holdings',
    // Main input: several PO PDFs (one per PO). Optional OC PDFs validate them.
    inputKind: 'pdfs',
    parsePo: parsePoPdf,
    build: buildOfficeHoldings,
    parseOc: parseOcPdf,
    reconcile: reconcileOffice,
    ocPricingBySku,
    allocateOcsToPos: allocateOcsToPosOffice,
  },
  {
    id: 'john-lewis',
    name: 'John Lewis',
    // Main input: one client Excel that already splits the order by PO. The
    // optional Aloha OC PDFs supply PRICE/DISCOUNT and validate quantities.
    inputKind: 'excel',
    // The Excel is pre-split per PO, so reconcile against the parsed POs
    // (per-PO columns, like Office Holdings) rather than a single Excel total.
    reconcileMode: 'per-po',
    parse: parseJohnLewis,
    build: buildJohnLewis,
    parseOc: parseOcPdf,
    reconcile: reconcileJohnLewis,
    ocPricingBySku,
  },
];

export function getClient(id) {
  return CLIENTS.find((c) => c.id === id);
}
