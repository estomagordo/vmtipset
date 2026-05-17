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

/** Group letters in schedule order for this template (A–K). */
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
] as const;

export const GROUP_STAGE_STANDING_ROWS = 4;
export const GROUP_STAGE_FIXTURE_ROWS = 6;

export const GROUP_STAGE_BLOCKS: readonly GroupStageBlock[] = GROUP_STAGE_LETTERS.map((letter, i) => ({
  id: `Grupp ${letter}`,
  startRow: GROUP_STAGE_FIRST_ROW + i * GROUP_STAGE_ROW_STRIDE,
}));
