import { useEffect, useMemo, useState } from 'react';
import { useTranslation } from 'react-i18next';
import { SpreadsheetView } from './components/SpreadsheetView';
import { translateWorkbookString } from './i18n/workbookStrings';
import { buildGridFromDump } from './lib/gridModel';
import type { WorkbookDump } from './types/workbook';

export default function App() {
  const { t, i18n } = useTranslation('app');
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

  const rawSheetTitle = dump.sheets[0]?.name ?? t('meta.defaultSheetTitle');
  const sheetTitle = translateWorkbookString(rawSheetTitle, i18n);

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
