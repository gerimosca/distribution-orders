import { parseKurtGeiger, buildWorkbook as buildKurtGeigerWorkbook } from './kurt-geiger.mjs';
import { parseOcPdf, excelTotalsBySku, reconcile, ocPricingBySku } from './oc-pdf.mjs';
import {
  parsePoPdf,
  build as buildOfficeHoldings,
  reconcile as reconcileOffice,
} from './office-holdings.mjs';

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
  },
];

export function getClient(id) {
  return CLIENTS.find((c) => c.id === id);
}
