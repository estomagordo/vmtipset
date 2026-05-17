import { GROUP_STAGE_BLOCKS, type GroupStageBlock } from '../config/groupStageBlocks';
import type { WorkbookCell } from '../types/workbook';
import { lettersToColIndex } from './excelAddress';
import { type StandingRow, compareStandings, readStandingRowFromWorkbook } from './groupStandings';

/** Mall sheet row with column labels S, V, O, … for the third-place ranking table. */
export const BEST_THIRD_PLACE_HEADER_ROW = 127;
/** Data rows J128:J139 (12 teams). */
export const BEST_THIRD_PLACE_FIRST_DATA_ROW = 128;
export const BEST_THIRD_PLACE_LAST_DATA_ROW = 139;

/** 0-based grid row for Excel {@link BEST_THIRD_PLACE_HEADER_ROW}. */
export const BEST_THIRD_PLACE_FIRST_GRID_ROW_INDEX = BEST_THIRD_PLACE_HEADER_ROW - 1;
/** Exclusive end for `grid.slice(BEST_THIRD_PLACE_FIRST_GRID_ROW_INDEX, …)` through row 139. */
export const BEST_THIRD_PLACE_GRID_SLICE_END_EXCLUSIVE = BEST_THIRD_PLACE_LAST_DATA_ROW;

const COL_J = lettersToColIndex('J');
const COL_K = lettersToColIndex('K');
const COL_L = lettersToColIndex('L');
const COL_M = lettersToColIndex('M');
const COL_N = lettersToColIndex('N');
const COL_O = lettersToColIndex('O');
const COL_P = lettersToColIndex('P');
const COL_Q = lettersToColIndex('Q');
const COL_R = lettersToColIndex('R');
const COL_S = lettersToColIndex('S');
const COL_T = lettersToColIndex('T');

export type BestThirdPlaceRow = StandingRow & { rank: number; groupLetter: string };

function thirdPlaceRowForGroup(
  block: GroupStageBlock,
  tables: Map<string, StandingRow[] | null>,
  cellByAddress: Map<string, WorkbookCell>,
): (StandingRow & { groupLetter: string }) | null {
  const live = tables.get(block.id);
  const letter = block.id.replace(/^Grupp /, '');
  if (live && live.length >= 3) return { ...live[2], groupLetter: letter };
  const excelRow = block.startRow + 2;
  const fromMall = readStandingRowFromWorkbook(cellByAddress, excelRow);
  return fromMall ? { ...fromMall, groupLetter: letter } : null;
}

/**
 * 3rd-placed team per group: live from predictions when available, otherwise the mall base row.
 */
export function computeBestThirdPlaceTable(
  tables: Map<string, StandingRow[] | null>,
  cellByAddress: Map<string, WorkbookCell>,
): BestThirdPlaceRow[] | null {
  const thirds: (StandingRow & { groupLetter: string })[] = [];
  for (const block of GROUP_STAGE_BLOCKS) {
    const row = thirdPlaceRowForGroup(block, tables, cellByAddress);
    if (row) thirds.push(row);
  }
  if (thirds.length === 0) return null;
  const sorted = [...thirds].sort(compareStandings);
  return sorted.map((row, i) => ({ ...row, rank: i + 1 }));
}

export function bestThirdPlaceDataRowIndex(excelRow: number): number | null {
  if (excelRow < BEST_THIRD_PLACE_FIRST_DATA_ROW || excelRow > BEST_THIRD_PLACE_LAST_DATA_ROW) {
    return null;
  }
  return excelRow - BEST_THIRD_PLACE_FIRST_DATA_ROW;
}

export function isBestThirdPlaceHeaderCell(excelRow: number, gridCol: number): boolean {
  return excelRow === BEST_THIRD_PLACE_HEADER_ROW && gridCol >= COL_J && gridCol <= COL_T;
}

export function isBestThirdPlaceDataCell(excelRow: number, gridCol: number): boolean {
  return (
    bestThirdPlaceDataRowIndex(excelRow) !== null && gridCol >= COL_J && gridCol <= COL_T
  );
}

function fieldForBestThirdGridCol(gridCol: number): keyof StandingRow | 'rank' | null {
  switch (gridCol) {
    case COL_J:
      return 'rank';
    case COL_K:
      return 'name';
    case COL_L:
      return 'played';
    case COL_M:
      return 'wins';
    case COL_N:
      return 'draws';
    case COL_O:
      return 'losses';
    case COL_P:
      return 'gf';
    case COL_Q:
      return 'ga';
    case COL_R:
      return 'gd';
    case COL_S:
      return 'pts';
    default:
      return null;
  }
}

export function getBestThirdPlaceCellValue(
  excelRow: number,
  gridCol: number,
  rows: BestThirdPlaceRow[] | null,
): string | number | null {
  if (rows === null) return null;
  const idx = bestThirdPlaceDataRowIndex(excelRow);
  if (idx === null) return null;
  const row = rows[idx];
  const field = fieldForBestThirdGridCol(gridCol);
  if (field === null) return null;
  if (!row) return null;
  if (field === 'rank') return row.rank;
  const v = row[field];
  return v as string | number;
}
