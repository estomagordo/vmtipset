import type { WorkbookCell } from '../types/workbook';
import { colIndexToLetters } from '../lib/excelAddress';
import type { GridCell } from '../lib/gridModel';

function truncate(s: string, max: number): string {
  if (s.length <= max) return s;
  return `${s.slice(0, max - 1)}…`;
}

function formatFormulaDisplay(f?: string): string {
  if (!f) return '';
  const t = f.startsWith('=') ? f : `=${f}`;
  return truncate(t, 48);
}

export type MallSheetProps = {
  title: string;
  grid: GridCell[][];
  rowCount: number;
  colCount: number;
};

function CellBody({ entry }: { entry?: WorkbookCell }) {
  if (!entry) return null;

  if (entry.kind === 'formula') {
    return (
      <span className="excel-cell-formula" title={entry.formula}>
        {formatFormulaDisplay(entry.formula)}
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

export function MallSheet({ title, grid, rowCount, colCount }: MallSheetProps) {
  const letters = Array.from({ length: colCount }, (_, i) => colIndexToLetters(i + 1));

  return (
    <div className="excel-shell">
      <header className="excel-docbar">
        <h1 className="excel-title">{title}</h1>
        <p className="excel-sub">Klientgränssnitt inspirerat av Excel-mallen (kolumner A–T).</p>
      </header>

      <div className="excel-scroll">
        <table className="excel-grid" role="grid" aria-rowcount={rowCount + 1} aria-colcount={colCount + 1}>
          <thead>
            <tr>
              <th className="excel-gutter excel-gutter-corner" scope="col" />
              {letters.map((L) => (
                <th key={L} className="excel-gutter excel-col-head" scope="col">
                  {L}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {grid.map((row, ri) => (
              <tr key={ri} aria-rowindex={ri + 2}>
                <th className="excel-gutter excel-row-head" scope="row">
                  {ri + 1}
                </th>
                {row.map((cell: GridCell, ci) => {
                  void ci;
                  if (cell.kind === 'skip') return null;
                  const entry = cell.entry;
                  const hasValidation = Boolean(entry?.validation?.type === 'list');
                  const isFormula = entry?.kind === 'formula';
                  const cls = [
                    'excel-cell',
                    entry?.kind === 'number' ? 'excel-align-right' : 'excel-align-left',
                    hasValidation ? 'excel-cell-validated' : '',
                    isFormula ? 'excel-cell-is-formula' : '',
                  ]
                    .filter(Boolean)
                    .join(' ');
                  return (
                    <td
                      key={cell.address}
                      className={cls}
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
