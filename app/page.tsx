'use client';

import { useState } from 'react';

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

export default function Page() {
  const [clientId, setClientId] = useState(CLIENTS[0].id);
  const [file, setFile] = useState<File | null>(null);
  const [poFiles, setPoFiles] = useState<File[]>([]);
  const [ocFiles, setOcFiles] = useState<File[]>([]);
  const [status, setStatus] = useState<Status>({ kind: 'idle' });

  const client = CLIENTS.find((c) => c.id === clientId)!;
  const ready = client.inputKind === 'pdfs' ? poFiles.length > 0 : !!file;

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
      URL.revokeObjectURL(url);

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
            <label htmlFor="poFiles">PDF de las PO (uno por PO, varios)</label>
            <input
              id="poFiles"
              type="file"
              accept=".pdf"
              multiple
              onChange={(e) => setPoFiles(e.target.files ? Array.from(e.target.files) : [])}
            />
            {poFiles.length > 0 && (
              <p className="hint">
                {poFiles.length} PO: {poFiles.map((f) => f.name).join(', ')}
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
          OC — Order Confirmation (PDF, opcional, una o varias)
        </label>
        <input
          id="ocFiles"
          type="file"
          accept=".pdf"
          multiple
          onChange={(e) => setOcFiles(e.target.files ? Array.from(e.target.files) : [])}
        />
        {ocFiles.length > 0 && (
          <p className="hint">
            {ocFiles.length} OC: {ocFiles.map((f) => f.name).join(', ')}
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
