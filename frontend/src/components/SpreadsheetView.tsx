import { useCallback, useState } from 'react';
import { useTranslation } from 'react-i18next';
import type { WorkbookCell } from '../types/workbook';
import { colIndexToLetters } from '../lib/excelAddress';
import type { GridCell } from '../lib/gridModel';
import {
  getScoreSelectValue,
  isScoreDigitListValidation,
  SCORE_OPTION_VALUES,
} from '../lib/scoreCells';
import { translateWorkbookString } from '../i18n/workbookStrings';

export type SpreadsheetViewProps = {
  title: string;
  grid: GridCell[][];
  rowCount: number;
  colCount: number;
};

type CellBodyProps = {
  entry?: WorkbookCell;
  scoreDrafts: Record<string, string>;
  onScoreDraft: (address: string, value: string) => void;
};

function CellBody({ entry, scoreDrafts, onScoreDraft }: CellBodyProps) {
  const { t, i18n } = useTranslation('app');

  if (!entry) return null;

  if (entry.kind === 'formula') {
    const cached = entry.cached_value;
    const hasCached = cached !== undefined && cached !== null;
    if (hasCached) {
      const isNum = typeof cached === 'number';
      const text = isNum ? String(cached) : translateWorkbookString(String(cached), i18n);
      return (
        <span
          className={isNum ? 'excel-cell-num' : 'excel-cell-text'}
          title={entry.formula ? String(entry.formula) : undefined}
        >
          {text}
        </span>
      );
    }
    return (
      <span
        className="excel-cell-text"
        title={
          entry.formula ? t('formula.tooltipNotCached', { formula: entry.formula }) : undefined
        }
      >
        {'\u00a0'}
      </span>
    );
  }

  if (isScoreDigitListValidation(entry.validation)) {
    const value = getScoreSelectValue(entry.address, entry, scoreDrafts);
    return (
      <select
        className="excel-dropdown excel-score-select"
        aria-label={t('aria.scorePicker', { address: entry.address })}
        value={value}
        onChange={(e) => onScoreDraft(entry.address, e.target.value)}
      >
        <option value="">{t('choices.blankScore')}</option>
        {SCORE_OPTION_VALUES.map((n) => (
          <option key={n} value={String(n)}>
            {n}
          </option>
        ))}
      </select>
    );
  }

  if (entry.kind === 'empty') {
    const v = entry.validation;
    if (v?.type === 'list' && Array.isArray(v.list_values) && v.list_values.length > 0) {
      return (
        <select
          className="excel-dropdown"
          aria-label={t('aria.dropdownCell', { address: entry.address })}
          defaultValue=""
        >
          <option value=""> </option>
          {v.list_values.map((opt) => {
            const s = String(opt);
            return (
              <option key={s} value={s}>
                {translateWorkbookString(s, i18n)}
              </option>
            );
          })}
        </select>
      );
    }
    return null;
  }

  if (entry.value != null && entry.value !== '') {
    const raw = String(entry.value);
    const text = entry.kind === 'number' ? raw : translateWorkbookString(raw, i18n);
    return (
      <span className={entry.kind === 'number' ? 'excel-cell-num' : 'excel-cell-text'}>{text}</span>
    );
  }

  return null;
}

export function SpreadsheetView({ title, grid, rowCount, colCount }: SpreadsheetViewProps) {
  const { t } = useTranslation('app');
  const [scoreDrafts, setScoreDrafts] = useState<Record<string, string>>({});
  const onScoreDraft = useCallback((address: string, value: string) => {
    setScoreDrafts((d) => ({ ...d, [address]: value }));
  }, []);

  const columnLabels = Array.from({ length: colCount }, (_, i) => colIndexToLetters(i + 1));

  return (
    <div className="excel-shell">
      <header className="excel-docbar">
        <h1 className="excel-title">{title}</h1>
        <p className="excel-sub">{t('spreadsheet.subtitle')}</p>
      </header>

      <div className="excel-scroll">
        <table className="excel-grid" role="grid" aria-rowcount={rowCount + 1} aria-colcount={colCount + 1}>
          <thead>
            <tr>
              <th className="excel-gutter excel-gutter-corner" scope="col" />
              {columnLabels.map((letter) => (
                <th key={letter} className="excel-gutter excel-col-head" scope="col">
                  {letter}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {grid.map((row, rowIndex) => (
              <tr key={rowIndex} aria-rowindex={rowIndex + 2}>
                <th className="excel-gutter excel-row-head" scope="row">
                  {rowIndex + 1}
                </th>
                {row.map((cell: GridCell, columnIndex) => {
                  void columnIndex;
                  if (cell.kind === 'skip') return null;
                  const entry = cell.entry;
                  const hasListValidation = entry?.validation?.type === 'list';
                  const isFormula = entry?.kind === 'formula';
                  const className = [
                    'excel-cell',
                    entry?.kind === 'number' || (isFormula && typeof entry?.cached_value === 'number')
                      ? 'excel-align-right'
                      : 'excel-align-left',
                    hasListValidation ? 'excel-cell-validated' : '',
                  ]
                    .filter(Boolean)
                    .join(' ');
                  return (
                    <td
                      key={cell.address}
                      className={className}
                      rowSpan={cell.rowSpan}
                      colSpan={cell.colSpan}
                      data-address={cell.address}
                    >
                      <CellBody
                        entry={entry}
                        scoreDrafts={scoreDrafts}
                        onScoreDraft={onScoreDraft}
                      />
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
