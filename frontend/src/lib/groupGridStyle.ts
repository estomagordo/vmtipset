import { lettersToColIndex } from './excelAddress';

const COL_D = lettersToColIndex('D');
const COL_J = lettersToColIndex('J');
const COL_B = lettersToColIndex('B');
const COL_F = lettersToColIndex('F');
const COL_K = lettersToColIndex('K');
const COL_T = lettersToColIndex('T');

/** Visual weight classes for cells inside a group-stage card (see index.css). */
export function classifyGroupGridCell(
  gridCol: number,
  excelRow: number,
  blockHeaderExcelRow: number,
  rowInBlock: number,
): string[] {
  const tier: string[] = [];

  if (rowInBlock === 0) {
    tier.push('tipset-cell--grp-h');
    return tier;
  }

  if (gridCol === COL_D || gridCol === COL_J) {
    tier.push('tipset-cell--grp-dj');
    return tier;
  }

  const fixtureFirstExcelRow = blockHeaderExcelRow + 1;
  const rank =
    excelRow >= fixtureFirstExcelRow && excelRow <= fixtureFirstExcelRow + 3
      ? excelRow - fixtureFirstExcelRow
      : -1;
  const inStatsBand = gridCol >= COL_K && gridCol <= COL_T;

  if (inStatsBand && rank >= 0) {
    if (rank <= 1) tier.push('tipset-cell--grp-t2');
    else if (rank === 2) tier.push('tipset-cell--grp-t3');
    else tier.push('tipset-cell--grp-t4');
    return tier;
  }

  if (rowInBlock >= 1 && (gridCol === COL_B || gridCol === COL_F)) {
    tier.push('tipset-cell--grp-t2');
  }

  return tier;
}

/** Column width hints for compact match / prediction layout (A–T). */
export const GROUP_TABLE_COL_CLASSES: readonly string[] = [
  'tipset-col--match',
  'tipset-col--pred-team',
  'tipset-col--score',
  'tipset-col--sep',
  'tipset-col--score',
  'tipset-col--pred-team',
  'tipset-col--x',
  'tipset-col--x',
  'tipset-col--x',
  'tipset-col--idx',
  'tipset-col--stat-team',
  ...Array.from({ length: 9 }, () => 'tipset-col--stat'),
] as const;
