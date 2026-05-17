import type { WorkbookCell } from '../types/workbook';
import { getScoreSelectValue, isScoreDigitListValidation } from './scoreCells';

export type MatchPick = 'home' | 'draw' | 'away';

/**
 * When both C and E scores are filled (0–9), return the 1×2 outcome; otherwise `null`.
 */
export function matchPickFromRow(
  matchExcelRow: number,
  cellByAddress: Map<string, WorkbookCell>,
  scoreDrafts: Record<string, string>,
): MatchPick | null {
  const cAddr = `C${matchExcelRow}`;
  const eAddr = `E${matchExcelRow}`;
  const c = cellByAddress.get(cAddr);
  const e = cellByAddress.get(eAddr);
  if (
    !isScoreDigitListValidation(c?.validation) ||
    !isScoreDigitListValidation(e?.validation)
  ) {
    return null;
  }

  const hs = readGoals(cAddr, c, scoreDrafts);
  const as = readGoals(eAddr, e, scoreDrafts);
  if (hs === null || as === null) return null;
  if (hs > as) return 'home';
  if (hs < as) return 'away';
  return 'draw';
}

function readGoals(
  address: string,
  entry: WorkbookCell | undefined,
  drafts: Record<string, string>,
): number | null {
  if (!entry) return null;
  const s = getScoreSelectValue(address, entry, drafts);
  if (s === '') return null;
  const n = parseInt(s, 10);
  if (Number.isNaN(n) || n < 0 || n > 9) return null;
  return n;
}
