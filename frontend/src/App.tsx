import { useEffect, useMemo, useState } from 'react';
import { useTranslation } from 'react-i18next';
import { SpreadsheetView } from './components/SpreadsheetView';
import { buildGridFromDump } from './lib/gridModel';
import type { WorkbookDump } from './types/workbook';

export default function App() {
  const { t } = useTranslation('app');
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
        <p className="excel-error">{t('loadError', { error })}</p>
        <p className="excel-sub">{t('loadErrorHint')}</p>
      </div>
    );
  }

  if (!gridModel || !dump) {
    return (
      <div className="excel-app">
        <p className="excel-loading">{t('loading')}</p>
      </div>
    );
  }

  const sheetTitle = dump.sheets[0]?.name ?? t('meta.defaultSheetTitle');

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
