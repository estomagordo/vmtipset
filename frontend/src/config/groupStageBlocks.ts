export type GroupStageBlock = {
  /** Label used in the mall (e.g. Grupp A); keys the computed standings map. */
  id: string;
  /**
   * First 1-based row of this block on the mall sheet: rows `startRow`…`startRow + 3` show K–S standings;
   * rows `startRow`…`startRow + 5` hold the six round-robin fixtures (B/C home, E/F away).
   */
  startRow: number;
};

/** Row index of the first group block (Grupp A) in the VM-tipset 2026 mall export. */
export const GROUP_STAGE_FIRST_ROW = 14;

/** Rows between each group's `startRow` (section headers / spacing rows in the template). */
export const GROUP_STAGE_ROW_STRIDE = 8;

/** Group letters in schedule order for this template (A–L). */
export const GROUP_STAGE_LETTERS = [
  'A',
  'B',
  'C',
  'D',
  'E',
  'F',
  'G',
  'H',
  'I',
  'J',
  'K',
  'L',
] as const;

export const GROUP_STAGE_STANDING_ROWS = 4;
export const GROUP_STAGE_FIXTURE_ROWS = 6;

/** Excel row of the "Match nr" / "Grupp X" header line for each block (row before first fixture). */
export const GROUP_STAGE_HEADER_EXCEL_ROW = GROUP_STAGE_FIRST_ROW - 1;

/** 0-based grid index of the first group header row (Excel row 13 in the mall). */
export const GROUP_STAGE_FIRST_HEADER_GRID_INDEX = GROUP_STAGE_HEADER_EXCEL_ROW - 1;

/** Header row + six fixture rows (excludes the blank gap row between groups in the sheet). */
export const GROUP_STAGE_BLOCK_DATA_ROWS = 7;

export const GROUP_STAGE_PREAMBLE_ROW_COUNT = GROUP_STAGE_FIRST_HEADER_GRID_INDEX;

/** First grid index after all group blocks and their gap rows (footer / playoff area). */
export function groupStageFooterStartGridIndex(): number {
  return GROUP_STAGE_FIRST_HEADER_GRID_INDEX + GROUP_STAGE_LETTERS.length * GROUP_STAGE_ROW_STRIDE;
}

/**
 * Maps a data row inside the group stage to its block, or `null` for preamble, gaps, or footer.
 */
export function tryParseGroupBlockRow(
  rowIndex0Based: number,
): { groupIndex: number; rowInBlock: number } | null {
  const o = rowIndex0Based - GROUP_STAGE_FIRST_HEADER_GRID_INDEX;
  if (o < 0) return null;
  const mod = o % GROUP_STAGE_ROW_STRIDE;
  if (mod >= GROUP_STAGE_BLOCK_DATA_ROWS) return null;
  const groupIndex = Math.floor(o / GROUP_STAGE_ROW_STRIDE);
  if (groupIndex >= GROUP_STAGE_LETTERS.length) return null;
  return { groupIndex, rowInBlock: mod };
}

export const GROUP_STAGE_BLOCKS: readonly GroupStageBlock[] = GROUP_STAGE_LETTERS.map((letter, i) => ({
  id: `Grupp ${letter}`,
  startRow: GROUP_STAGE_FIRST_ROW + i * GROUP_STAGE_ROW_STRIDE,
}));
