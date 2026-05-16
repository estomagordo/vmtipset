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
  let maxR = 1;

  for (const c of sheet.cells) {
    maxR = Math.max(maxR, parseA1(c.address).row);
  }

  for (const raw of sheet.merged_ranges) {
    for (const spec of splitRangeSpecs(raw)) {
      try {
        const b = rangeBounds(spec);
        const clipped = clipRangeBounds(b, 99999, colCount);
        if (clipped) maxR = Math.max(maxR, clipped.r1);
      } catch {
        /* ignore */
      }
    }
  }

  if (regionMaxRowHint > 0) maxR = Math.max(maxR, regionMaxRowHint);
  return maxR;
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
  const hint = Number(data.parse?.region_max_row);
  const rowCount = inferRowCount(sheet, colCount, Number.isFinite(hint) ? hint : 0);
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
      const r0 = clipped.r0 - 1;
      const c0 = clipped.c0 - 1;

      const anchor = grid[r0][c0];
      if (anchor.kind !== 'anchor') continue;
      anchor.rowSpan = rowSpan;
      anchor.colSpan = colSpan;

      for (let rr = clipped.r0 - 1; rr < clipped.r1; rr++) {
        for (let cc = clipped.c0 - 1; cc < clipped.c1; cc++) {
          if (rr === r0 && cc === c0) continue;
          grid[rr][cc] = { kind: 'skip' };
        }
      }
    }
  }

  return { rowCount, colCount, grid, cellByAddress };
}
