import { useEffect, useMemo, useState } from 'react';
import { SpreadsheetView } from './components/SpreadsheetView';
import { buildGridFromDump } from './lib/gridModel';
import type { WorkbookDump } from './types/workbook';

export default function App() {
  const [dump, setDump] = useState<WorkbookDump | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let stillMounted = true;
    fetch('/workbook_dump.json')
      .then((r) => {
        if (!r.ok) throw new Error(`${r.status} ${r.statusText}`);
        return r.json();
      })
      .then((data: WorkbookDump) => {
        if (!stillMounted) return;
        setDump(data);
      })
      .catch((e: unknown) => {
        if (!stillMounted) return;
        setError(e instanceof Error ? e.message : String(e));
      });
    return () => {
      stillMounted = false;
    };
  }, []);

  const gridModel = useMemo(() => (dump ? buildGridFromDump(dump) : null), [dump]);

  if (error) {
    return (
      <div className="excel-app">
        <p className="excel-error">Failed to load workbook_dump.json: {error}</p>
        <p className="excel-sub">
          Run{' '}
          <code>python tools/parse_workbook.py ... --json tools/workbook_dump.json</code> and copy it to{' '}
          <code>frontend/public/</code>.
        </p>
      </div>
    );
  }

  if (!gridModel || !dump) {
    return (
      <div className="excel-app">
        <p className="excel-loading">Loading workbook…</p>
      </div>
    );
  }

  const sheetTitle = dump.sheets[0]?.name ?? 'VM-tipset';

  return (
    <div className="excel-app">
      <SpreadsheetView
        title={sheetTitle}
        grid={gridModel.grid}
        rowCount={gridModel.rowCount}
        colCount={gridModel.colCount}
      />
    </div>
  );
}
