import { useCallback, useMemo, useState } from 'react';
import { useTranslation } from 'react-i18next';
import type { WorkbookCell } from '../types/workbook';
import type { GridCell } from '../lib/gridModel';
import {
  computeGroupStandings,
  resolveStandingsCellAtAddress,
  type StandingRow,
} from '../lib/groupStandings';
import { GROUP_STAGE_BLOCKS } from '../config/groupStageBlocks';
import {
  getScoreSelectValue,
  isScoreDigitListValidation,
  SCORE_OPTION_VALUES,
} from '../lib/scoreCells';
import { lettersToColIndex } from '../lib/excelAddress';
import { translateWorkbookString } from '../i18n/workbookStrings';

export type SpreadsheetViewProps = {
  title: string;
  grid: GridCell[][];
  rowCount: number;
  colCount: number;
  cellByAddress: Map<string, WorkbookCell>;
};

/** 1-based inclusive column range K–S (standings block in the mall). */
const STANDINGS_COL_START = lettersToColIndex('K');
const STANDINGS_COL_END = lettersToColIndex('S');

type CellBodyProps = {
  entry?: WorkbookCell;
  scoreDrafts: Record<string, string>;
  onScoreDraft: (address: string, value: string) => void;
  standingTables: Map<string, StandingRow[] | null>;
};

function CellBody({ entry, scoreDrafts, onScoreDraft, standingTables }: CellBodyProps) {
  const { t, i18n } = useTranslation('app');

  if (!entry) return null;

  if (entry.kind === 'formula') {
    const standing = resolveStandingsCellAtAddress(entry.address, standingTables, GROUP_STAGE_BLOCKS);
    if (standing.kind === 'value') {
      const v = standing.value;
      const isNum = typeof v === 'number';
      const text = isNum ? String(v) : translateWorkbookString(String(v), i18n);
      return (
        <span
          className={isNum ? 'tipset-cell-num' : 'tipset-cell-text'}
          title={entry.formula ? String(entry.formula) : undefined}
        >
          {text}
        </span>
      );
    }

    const cached = entry.cached_value;
    const hasCached = cached !== undefined && cached !== null;
    if (hasCached) {
      const isNum = typeof cached === 'number';
      const text = isNum ? String(cached) : translateWorkbookString(String(cached), i18n);
      return (
        <span
          className={isNum ? 'tipset-cell-num' : 'tipset-cell-text'}
          title={entry.formula ? String(entry.formula) : undefined}
        >
          {text}
        </span>
      );
    }
    return (
      <span
        className="tipset-cell-text"
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
        className="tipset-select tipset-select--in-cell"
        aria-label={t('aria.scorePicker', { address: entry.address })}
        value={value}
        onChange={(e) => onScoreDraft(entry.address, e.target.value)}
      >
        <option value="" />
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
          className="tipset-select tipset-select--in-cell"
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
      <span className={entry.kind === 'number' ? 'tipset-cell-num' : 'tipset-cell-text'}>{text}</span>
    );
  }

  return null;
}

export function SpreadsheetView({ title, grid, rowCount, colCount, cellByAddress }: SpreadsheetViewProps) {
  const { t } = useTranslation('app');
  const [scoreDrafts, setScoreDrafts] = useState<Record<string, string>>({});
  const onScoreDraft = useCallback((address: string, value: string) => {
    setScoreDrafts((d) => ({ ...d, [address]: value }));
  }, []);

  const standingTables = useMemo(
    () => computeGroupStandings(cellByAddress, scoreDrafts),
    [cellByAddress, scoreDrafts],
  );

  return (
    <div className="tipset-shell">
      <header className="tipset-header">
        <h1 className="tipset-title">{title}</h1>
        <p className="tipset-sub">{t('spreadsheet.subtitle')}</p>
      </header>

      <div className="tipset-scroll">
        <div className="tipset-table-wrap">
          <table
            className="tipset-table"
            role="grid"
            aria-label={title}
            aria-rowcount={rowCount}
            aria-colcount={colCount}
          >
            <tbody>
              {grid.map((row, rowIndex) => (
                <tr key={rowIndex} aria-rowindex={rowIndex + 1}>
                  {row.map((cell: GridCell, columnIndex) => {
                    if (cell.kind === 'skip') return null;
                    const entry = cell.entry;
                    const gridCol = columnIndex + 1;
                    const isStandingsColumn =
                      gridCol >= STANDINGS_COL_START && gridCol <= STANDINGS_COL_END;
                    const hasListValidation = entry?.validation?.type === 'list';
                    const isFormula = entry?.kind === 'formula';
                    const className = [
                      'tipset-cell',
                      entry?.kind === 'number' || (isFormula && typeof entry?.cached_value === 'number')
                        ? 'tipset-align-right'
                        : 'tipset-align-left',
                      isStandingsColumn ? 'tipset-cell--standings' : '',
                      hasListValidation ? 'tipset-cell--input' : '',
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
                        aria-colindex={gridCol}
                      >
                        <CellBody
                          entry={entry}
                          scoreDrafts={scoreDrafts}
                          onScoreDraft={onScoreDraft}
                          standingTables={standingTables}
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
    </div>
  );
}
