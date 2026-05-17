import type { DataValidation } from '../types/workbook';

/** Digits 0–9 as score picks (same as Excel list on C/E group-stage columns in the mall file). */
export const SCORE_OPTION_VALUES = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9] as const;

export function isScoreDigitListValidation(v: DataValidation | undefined): boolean {
  if (!v || v.type !== 'list' || !Array.isArray(v.list_values)) {
    return false;
  }
  const nums = v.list_values.map((x) => (typeof x === 'number' ? x : parseInt(String(x), 10)));
  if (nums.length !== 10 || nums.some((n) => Number.isNaN(n))) {
    return false;
  }
  const set = new Set(nums);
  for (let i = 0; i <= 9; i++) {
    if (!set.has(i)) {
      return false;
    }
  }
  return true;
}

/** Effective score text for the select value attribute: '' or '0'..'9'. */
export function getWorkbookScoreBase(entry: { value?: unknown }): string {
  const raw = entry.value;
  if (raw == null || raw === '') {
    return '';
  }
  const s = String(raw).trim();
  if (s === '') {
    return '';
  }
  const n = parseInt(s, 10);
  if (!Number.isNaN(n) && n >= 0 && n <= 9 && String(n) === s) {
    return String(n);
  }
  return s;
}

export function getScoreSelectValue(
  address: string,
  entry: { value?: unknown },
  drafts: Record<string, string>,
): string {
  if (Object.prototype.hasOwnProperty.call(drafts, address)) {
    return drafts[address];
  }
  return getWorkbookScoreBase(entry);
}
