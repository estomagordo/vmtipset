import type { WorkbookCell, WorkbookDump, WorkbookSheet } from '../types/workbook';
import { clipRangeBounds, formatA1, lettersToColIndex, parseA1, rangeBounds, splitRangeSpecs } from './excelAddress';

export type GridAnchor = {
  kind: 'anchor';
  row: number;
  col: number;
  address: string;
  rowSpan: number;
  colSpan: number;
  entry?: WorkbookCell;
};

export type GridSkip = { kind: 'skip' };

export type GridCell = GridAnchor | GridSkip;

function cellMapFromSheet(sheet: WorkbookSheet): Map<string, WorkbookCell> {
  const m = new Map<string, WorkbookCell>();
  for (const c of sheet.cells) {
    m.set(c.address.toUpperCase().replace(/\$/g, ''), c);
  }
  return m;
}

function inferRowCount(sheet: WorkbookSheet, colCount: number, regionMaxRowHint: number): number {
  let maxRow = 1;

  for (const c of sheet.cells) {
    maxRow = Math.max(maxRow, parseA1(c.address).row);
  }

  for (const raw of sheet.merged_ranges) {
    for (const spec of splitRangeSpecs(raw)) {
      try {
        const b = rangeBounds(spec);
        const clipped = clipRangeBounds(b, 99999, colCount);
        if (clipped) maxRow = Math.max(maxRow, clipped.r1);
      } catch {
        /* ignore */
      }
    }
  }

  if (regionMaxRowHint > 0) maxRow = Math.max(maxRow, regionMaxRowHint);
  return maxRow;
}

/** 0-based grid[row][col] */
export function buildGridFromDump(data: WorkbookDump): {
  rowCount: number;
  colCount: number;
  grid: GridCell[][];
  cellByAddress: Map<string, WorkbookCell>;
} {
  const sheet = data.sheets[0];
  const colCount = lettersToColIndex(String(data.parse?.region_max_column_letter ?? 'T'));
  const rowCountHint = Number(data.parse?.region_max_row);
  const rowCount = inferRowCount(sheet, colCount, Number.isFinite(rowCountHint) ? rowCountHint : 0);
  const cellByAddress = cellMapFromSheet(sheet);

  const grid: GridCell[][] = [];
  for (let r = 0; r < rowCount; r++) {
    const row: GridCell[] = [];
    for (let c = 0; c < colCount; c++) {
      const addr = formatA1(r + 1, c + 1);
      row.push({
        kind: 'anchor',
        row: r + 1,
        col: c + 1,
        address: addr,
        rowSpan: 1,
        colSpan: 1,
        entry: cellByAddress.get(addr),
      });
    }
    grid.push(row);
  }

  for (const raw of sheet.merged_ranges) {
    for (const spec of splitRangeSpecs(raw)) {
      let b: ReturnType<typeof rangeBounds>;
      try {
        b = rangeBounds(spec);
      } catch {
        continue;
      }
      const clipped = clipRangeBounds(b, rowCount, colCount);
      if (!clipped) continue;

      const rowSpan = clipped.r1 - clipped.r0 + 1;
      const colSpan = clipped.c1 - clipped.c0 + 1;
      const anchorRow = clipped.r0 - 1;
      const anchorCol = clipped.c0 - 1;

      const anchor = grid[anchorRow][anchorCol];
      if (anchor.kind !== 'anchor') continue;
      anchor.rowSpan = rowSpan;
      anchor.colSpan = colSpan;

      for (let rowIdx = clipped.r0 - 1; rowIdx < clipped.r1; rowIdx++) {
        for (let colIdx = clipped.c0 - 1; colIdx < clipped.c1; colIdx++) {
          if (rowIdx === anchorRow && colIdx === anchorCol) continue;
          grid[rowIdx][colIdx] = { kind: 'skip' };
        }
      }
    }
  }

  return { rowCount, colCount, grid, cellByAddress };
}
