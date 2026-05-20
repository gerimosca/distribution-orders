'use client';

import { useRef, useState } from 'react';

type InputKind = 'excel' | 'pdfs';
const CLIENTS: { id: string; name: string; inputKind: InputKind }[] = [
  { id: 'kurt-geiger', name: 'Kurt Geiger', inputKind: 'excel' },
  { id: 'office-holdings', name: 'Office Holdings', inputKind: 'pdfs' },
];

type Status =
  | { kind: 'idle' }
  | { kind: 'loading' }
  | { kind: 'error'; msg: string }
  | { kind: 'success'; msg: string };

// Vercel's serverless functions cap request bodies around 4.5 MB. Multipart
// adds a few % overhead, so warn a bit earlier than the hard limit.
const UPLOAD_SOFT_LIMIT = 4 * 1024 * 1024;

// Heuristic file classification by name. Used to flag obvious mis-uploads
// (a PO07 file dropped in the OC field, or vice versa) so the user catches
// it before submitting. "unknown" doesn't trigger any warning.
function classifyFile(name: string): 'po' | 'oc' | 'unknown' {
  if (/^PO\d/i.test(name)) return 'po';
  if (/^Order\s*-?\s*S\d/i.test(name)) return 'oc';
  return 'unknown';
}

// Merge new files into an existing list without duplicates (by name+size).
// Lets the user click the input several times to add files in batches —
// the previous single-shot behavior silently dropped earlier selections,
// which caused the "Excel only has 4 OC tabs" mishap.
function mergeFiles(existing: File[], added: File[]): File[] {
  const seen = new Set(existing.map((f) => `${f.name}:${f.size}`));
  const out = [...existing];
  for (const f of added) {
    const key = `${f.name}:${f.size}`;
    if (!seen.has(key)) {
      seen.add(key);
      out.push(f);
    }
  }
  return out;
}

export default function Page() {
  const [clientId, setClientId] = useState(CLIENTS[0].id);
  const [file, setFile] = useState<File | null>(null);
  const [poFiles, setPoFiles] = useState<File[]>([]);
  const [ocFiles, setOcFiles] = useState<File[]>([]);
  const [status, setStatus] = useState<Status>({ kind: 'idle' });

  const poInputRef = useRef<HTMLInputElement>(null);
  const ocInputRef = useRef<HTMLInputElement>(null);

  const client = CLIENTS.find((c) => c.id === clientId)!;
  const ready = client.inputKind === 'pdfs' ? poFiles.length > 0 : !!file;

  const totalBytes =
    (file?.size ?? 0) +
    poFiles.reduce((s, f) => s + f.size, 0) +
    ocFiles.reduce((s, f) => s + f.size, 0);
  const overLimit = totalBytes > UPLOAD_SOFT_LIMIT;

  const misfiledInPo = poFiles.filter((f) => classifyFile(f.name) === 'oc');
  const misfiledInOc = ocFiles.filter((f) => classifyFile(f.name) === 'po');

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    if (!ready) return;
    setStatus({ kind: 'loading' });
    try {
      const fd = new FormData();
      fd.append('client', clientId);
      if (client.inputKind === 'pdfs') {
        for (const p of poFiles) fd.append('poFiles', p);
      } else if (file) {
        fd.append('file', file);
      }
      for (const oc of ocFiles) fd.append('ocFiles', oc);

      const res = await fetch('/api/process', { method: 'POST', body: fd });
      if (!res.ok) {
        // 413 doesn't always carry a body — explain it ourselves so the user
        // knows it's a size issue, not a parser failure.
        if (res.status === 413) {
          throw new Error(
            'Los archivos superan el límite del servidor (~4,5 MB). Sube los PDFs en lotes más pequeños.'
          );
        }
        throw new Error((await res.text()) || 'Error procesando');
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      const cd = res.headers.get('content-disposition') ?? '';
      const match = cd.match(/filename="(.+?)"/);
      a.download = match ? match[1] : 'output.xlsx';
      a.click();
      // Revoking immediately cancels the download in some Safari versions;
      // give the browser a beat to start the transfer first.
      setTimeout(() => URL.revokeObjectURL(url), 1000);

      const warnings = res.headers.get('x-warnings');
      const discrepancies = res.headers.get('x-discrepancies');
      const parts: string[] = ['Descargado.'];
      if (discrepancies && discrepancies !== '0') {
        parts.push(
          `${discrepancies} discrepancia(s) — revisa la pestaña "Discrepancias".`
        );
      } else if (ocFiles.length) {
        parts.push('Las cantidades cuadran con la OC ✔');
      }
      if (warnings && warnings !== '0') {
        parts.push(`${warnings} aviso(s) en las pestañas.`);
      }
      setStatus({ kind: 'success', msg: parts.join(' ') });
    } catch (err) {
      setStatus({ kind: 'error', msg: err instanceof Error ? err.message : String(err) });
    }
  }

  return (
    <main className="container">
      <h1>Distribution Orders</h1>
      <form onSubmit={handleSubmit}>
        <label htmlFor="client">Cliente</label>
        <select
          id="client"
          value={clientId}
          onChange={(e) => {
            setClientId(e.target.value);
            setFile(null);
            setPoFiles([]);
            setOcFiles([]);
            setStatus({ kind: 'idle' });
          }}
        >
          {CLIENTS.map((c) => (
            <option key={c.id} value={c.id}>
              {c.name}
            </option>
          ))}
        </select>

        {client.inputKind === 'pdfs' ? (
          <>
            <label htmlFor="poFiles">PDF de las PO del cliente Office Holdings (uno o varios)</label>
            <input
              id="poFiles"
              ref={poInputRef}
              type="file"
              accept=".pdf"
              multiple
              onChange={(e) => {
                const added = e.target.files ? Array.from(e.target.files) : [];
                setPoFiles((prev) => mergeFiles(prev, added));
                // Clear the input so re-selecting the same file fires onChange.
                if (poInputRef.current) poInputRef.current.value = '';
              }}
            />
            {poFiles.length > 0 && (
              <p className="hint">
                {poFiles.length} PO seleccionado(s){' '}
                <button
                  type="button"
                  onClick={() => setPoFiles([])}
                  style={{ background: 'none', border: 'none', color: '#0070f3', cursor: 'pointer', padding: 0, fontSize: 'inherit' }}
                >
                  (limpiar)
                </button>
                : {poFiles.map((f) => f.name).join(', ')}
              </p>
            )}
            {misfiledInPo.length > 0 && (
              <p className="hint" style={{ color: '#b45309' }}>
                Aviso: {misfiledInPo.length} archivo(s) en este campo
                parecen Order Confirmations ({misfiledInPo.map((f) => f.name).join(', ')}).
                Muévelos al campo "OC" de abajo.
              </p>
            )}
          </>
        ) : (
          <>
            <label htmlFor="file">Fichero del cliente (Excel)</label>
            <input
              id="file"
              type="file"
              accept=".xlsx,.xls"
              onChange={(e) => setFile(e.target.files?.[0] ?? null)}
            />
          </>
        )}

        <label htmlFor="ocFiles">
          OC — Order Confirmation (PDF Alohas — uno o varios)
        </label>
        <input
          id="ocFiles"
          ref={ocInputRef}
          type="file"
          accept=".pdf"
          multiple
          onChange={(e) => {
            const added = e.target.files ? Array.from(e.target.files) : [];
            setOcFiles((prev) => mergeFiles(prev, added));
            if (ocInputRef.current) ocInputRef.current.value = '';
          }}
        />
        {ocFiles.length > 0 && (
          <p className="hint">
            {ocFiles.length} OC seleccionado(s){' '}
            <button
              type="button"
              onClick={() => setOcFiles([])}
              style={{ background: 'none', border: 'none', color: '#0070f3', cursor: 'pointer', padding: 0, fontSize: 'inherit' }}
            >
              (limpiar)
            </button>
            : {ocFiles.map((f) => f.name).join(', ')}
          </p>
        )}
        {misfiledInOc.length > 0 && (
          <p className="hint" style={{ color: '#b45309' }}>
            Aviso: {misfiledInOc.length} archivo(s) en este campo parecen
            POs ({misfiledInOc.map((f) => f.name).join(', ')}).
            Muévelos al campo "PO" de arriba.
          </p>
        )}

        {overLimit && (
          <p className="hint" style={{ color: '#b45309' }}>
            Aviso: {(totalBytes / (1024 * 1024)).toFixed(1)} MB superan el
            límite recomendado (~4 MB). El servidor puede rechazar el envío;
            sube los PDFs en lotes más pequeños si falla.
          </p>
        )}

        <button type="submit" disabled={!ready || status.kind === 'loading'}>
          {status.kind === 'loading' ? 'Procesando...' : 'Generar Excel'}
        </button>
      </form>

      {status.kind === 'error' && <div className="status error">{status.msg}</div>}
      {status.kind === 'success' && <div className="status success">{status.msg}</div>}
    </main>
  );
}
