import { lettersToColIndex } from './excelAddress';

const COL_D = lettersToColIndex('D');
const COL_J = lettersToColIndex('J');
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

  /* Hyphen column stays accent on every row; match index (J) only on header row. */
  if (gridCol === COL_D) {
    tier.push('tipset-cell--grp-dj');
    /* Stacking: keep full separator band above neighbouring score <select> painting. */
    tier.push('tipset-cell--grp-sep');
    return tier;
  }

  /* Match index (J): rows showing 1–4 in the mall, not the two fixture rows below. */
  if (gridCol === COL_J && rowInBlock >= 1 && rowInBlock <= 4) {
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

  return tier;
}

/** Percent widths for group-stage columns (A–T); sum = 100 (used with `table-layout: fixed`). */
/** A=wider match nr, C/E=readable scores, D=wider ― band; trims elsewhere to keep sum 100. */
export const GROUP_COL_WIDTH_PCT: readonly number[] = [
  7, 9, 7, 4, 7, 9, 3, 3, 3, 3, 8, 4, 4, 4, 4, 4, 4, 4, 4, 5,
] as const;

const sumPct = (from: number, toExclusive: number) =>
  GROUP_COL_WIDTH_PCT.slice(from, toExclusive).reduce((a, b) => a + b, 0);

/** Sums of {@link GROUP_COL_WIDTH_PCT} for CSS grid tracks: A–F, G–I, J–T (each sums with others to 100). */
export const GROUP_STAGE_GRID_MATCH_FR = sumPct(0, 6);
export const GROUP_STAGE_GRID_X_FR = sumPct(6, 9);
export const GROUP_STAGE_GRID_STANDINGS_FR = sumPct(9, GROUP_COL_WIDTH_PCT.length);

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
