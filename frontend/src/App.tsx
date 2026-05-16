import { useEffect, useMemo, useState } from 'react';
import { MallSheet } from './components/MallSheet';
import { buildGridFromDump } from './lib/gridModel';
import type { WorkbookDump } from './types/workbook';

export default function App() {
  const [dump, setDump] = useState<WorkbookDump | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let ok = true;
    fetch('/workbook_dump.json')
      .then((r) => {
        if (!r.ok) throw new Error(`${r.status} ${r.statusText}`);
        return r.json();
      })
      .then((data: WorkbookDump) => {
        if (!ok) return;
        setDump(data);
      })
      .catch((e: unknown) => {
        if (!ok) return;
        setError(e instanceof Error ? e.message : String(e));
      });
    return () => {
      ok = false;
    };
  }, []);

  const model = useMemo(() => (dump ? buildGridFromDump(dump) : null), [dump]);

  if (error) {
    return (
      <div className="excel-app">
        <p className="excel-error">Kunde inte ladda workbook_dump.json: {error}</p>
        <p className="excel-sub">Kör: <code>python tools/parse_workbook.py ... --json tools/workbook_dump.json</code> och kopiera till <code>frontend/public/</code>.</p>
      </div>
    );
  }

  if (!model || !dump) {
    return (
      <div className="excel-app">
        <p className="excel-loading">Laddar mall…</p>
      </div>
    );
  }

  const sheetName = dump.sheets[0]?.name ?? 'VM-tipset';

  return (
    <div className="excel-app">
      <MallSheet title={sheetName} grid={model.grid} rowCount={model.rowCount} colCount={model.colCount} />
    </div>
  );
}
