import { useEffect, useMemo, useState } from 'react';
import { useTranslation } from 'react-i18next';
import { SpreadsheetView } from './components/SpreadsheetView';
import { translateWorkbookString } from './i18n/workbookStrings';
import { buildFixturesModel } from './lib/fixturesModel';
import { buildGridFromDump } from './lib/gridModel';
import type { FixturesDump } from './types/fixtures';
import type { WorkbookDump } from './types/workbook';

export default function App() {
  const { t, i18n } = useTranslation('app');
  const [dump, setDump] = useState<WorkbookDump | null>(null);
  const [fixturesDump, setFixturesDump] = useState<FixturesDump | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let stillMounted = true;
    Promise.all([
      fetch('/workbook_dump.json').then((r) => {
        if (!r.ok) throw new Error(`workbook_dump.json: ${r.status} ${r.statusText}`);
        return r.json() as Promise<WorkbookDump>;
      }),
      fetch('/fixtures.json').then((r) => {
        if (!r.ok) throw new Error(`fixtures.json: ${r.status} ${r.statusText}`);
        return r.json() as Promise<FixturesDump>;
      }),
    ])
      .then(([workbook, fixtures]) => {
        if (!stillMounted) return;
        setDump(workbook);
        setFixturesDump(fixtures);
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
  const fixtures = useMemo(
    () => (fixturesDump ? buildFixturesModel(fixturesDump) : null),
    [fixturesDump],
  );

  if (error) {
    return (
      <div className="tipset-app">
        <p className="tipset-error">{t('loadError', { error })}</p>
        <p className="tipset-error-hint">{t('loadErrorHint')}</p>
      </div>
    );
  }

  if (!gridModel || !dump || !fixtures) {
    return (
      <div className="tipset-app">
        <p className="tipset-loading">{t('loading')}</p>
      </div>
    );
  }

  const rawSheetTitle = dump.sheets[0]?.name ?? t('meta.defaultSheetTitle');
  const sheetTitle = translateWorkbookString(rawSheetTitle, i18n);

  return (
    <div className="tipset-app">
      <SpreadsheetView
        title={sheetTitle}
        grid={gridModel.grid}
        rowCount={gridModel.rowCount}
        colCount={gridModel.colCount}
        cellByAddress={gridModel.cellByAddress}
        fixtures={fixtures}
      />
    </div>
  );
}
