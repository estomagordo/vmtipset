import type { WorkbookCell } from '../types/workbook';
import { colIndexToLetters } from '../lib/excelAddress';
import type { GridCell } from '../lib/gridModel';

export type SpreadsheetViewProps = {
  title: string;
  grid: GridCell[][];
  rowCount: number;
  colCount: number;
};

function CellBody({ entry }: { entry?: WorkbookCell }) {
  if (!entry) return null;

  if (entry.kind === 'formula') {
    const cached = entry.cached_value;
    const hasCached = cached !== undefined && cached !== null;
    if (hasCached) {
      const isNum = typeof cached === 'number';
      return (
        <span
          className={isNum ? 'excel-cell-num' : 'excel-cell-text'}
          title={entry.formula ? String(entry.formula) : undefined}
        >
          {String(cached)}
        </span>
      );
    }
    // Excel often stores no cached value when IF(...,"",...) yields blank (e.g. unset bracket).
    // Normal Excel UI shows an empty cell, not the formula text.
    return (
      <span
        className="excel-cell-text"
        title={entry.formula ? `Formula (not cached in file): ${entry.formula}` : undefined}
      >
        {'\u00a0'}
      </span>
    );
  }

  if (entry.kind === 'empty') {
    const v = entry.validation;
    if (v?.type === 'list' && Array.isArray(v.list_values) && v.list_values.length > 0) {
      return (
        <select className="excel-dropdown" aria-label={entry.address} defaultValue="">
          <option value=""> </option>
          {v.list_values.map((opt) => (
            <option key={String(opt)} value={String(opt)}>
              {String(opt)}
            </option>
          ))}
        </select>
      );
    }
    return null;
  }

  if (entry.value != null && entry.value !== '') {
    return (
      <span className={entry.kind === 'number' ? 'excel-cell-num' : 'excel-cell-text'}>{String(entry.value)}</span>
    );
  }

  return null;
}

export function SpreadsheetView({ title, grid, rowCount, colCount }: SpreadsheetViewProps) {
  const columnLabels = Array.from({ length: colCount }, (_, i) => colIndexToLetters(i + 1));

  return (
    <div className="excel-shell">
      <header className="excel-docbar">
        <h1 className="excel-title">{title}</h1>
        <p className="excel-sub">
          Columns A–T. Formula cells without a saved value in the workbook render empty (like Excel); hover for the
          formula.
        </p>
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
                  const hasValidation = Boolean(entry?.validation?.type === 'list');
                  const isFormula = entry?.kind === 'formula';
                  const className = [
                    'excel-cell',
                    entry?.kind === 'number' || (isFormula && typeof entry?.cached_value === 'number')
                      ? 'excel-align-right'
                      : 'excel-align-left',
                    hasValidation ? 'excel-cell-validated' : '',
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
                      <CellBody entry={entry} />
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
