import type { WorkbookCell } from '../types/workbook';
import type { R32Slot } from '../types/fixtures';
import { parseA1 } from './excelAddress';
import { readStandingRowFromWorkbook, type StandingRow } from './groupStandings';
import { annexCAssignmentsForAdvancingThirds } from '../data/annexC2026';
import type { BestThirdPlaceRow } from './bestThirdPlace';
import {
  groupBlockForLetter,
  isBestThirdSlot,
  isGroupRankSlot,
  r32MatchById,
  type FixturesModel,
} from './fixturesModel';

/** Row with "16 delsfinal" header (left). */
export const R32_HEADER_ROW = 109;
/** M73 … M88 — team columns B/F use live standings + Annex C. */
export const R32_FIRST_MATCH_ROW = 110;
export const R32_LAST_MATCH_ROW = 125;

/** 0-based grid row index for Excel {@link R32_HEADER_ROW}. */
export const R32_FIRST_GRID_ROW_INDEX = R32_HEADER_ROW - 1;
/** Exclusive end index for `grid.slice(R32_FIRST_GRID_ROW_INDEX, …)`. */
export const R32_GRID_SLICE_END_EXCLUSIVE = R32_LAST_MATCH_ROW;

const COL_B = 2;
const COL_F = 6;

function matchIdForExcelRow(excelRow: number, fixtures: FixturesModel): string | null {
  const m = fixtures.r32ByExcelRow.get(excelRow);
  return m?.match_id ?? null;
}

function teamAtGroupRank(
  groupLetter: string,
  rank1Based: number,
  tables: Map<string, StandingRow[] | null>,
  cellByAddress: Map<string, WorkbookCell>,
  fixtures: FixturesModel,
): string | null {
  const rank0 = rank1Based - 1;
  if (rank0 < 0) return null;
  const id = `Grupp ${groupLetter.toUpperCase()}`;
  const live = tables.get(id);
  if (live && live.length > rank0) {
    const n = live[rank0]?.name?.trim();
    return n || null;
  }
  const block = groupBlockForLetter(fixtures, groupLetter);
  if (!block) return null;
  const fromFixture = block.teams[rank0]?.trim();
  if (fromFixture) return fromFixture;
  return readStandingRowFromWorkbook(cellByAddress, block.start_row + rank0)?.name?.trim() ?? null;
}

function thirdPlacedTeamFromGroup(
  groupLetter: string,
  tables: Map<string, StandingRow[] | null>,
  cellByAddress: Map<string, WorkbookCell>,
  fixtures: FixturesModel,
): string | null {
  return teamAtGroupRank(groupLetter, 3, tables, cellByAddress, fixtures);
}

function codeToThirdGroup(code: string): string | null {
  const m = code.trim().toUpperCase().match(/^3([A-L])$/);
  return m ? m[1] : null;
}

function thirdTeamForAnnexIndex(
  annexIndex: number,
  bestThird: BestThirdPlaceRow[] | null,
  tables: Map<string, StandingRow[] | null>,
  cellByAddress: Map<string, WorkbookCell>,
  fixtures: FixturesModel,
): string | null {
  if (!bestThird || bestThird.length < 8) return null;
  const top8 = bestThird.filter((r) => r.rank <= 8);
  if (top8.length !== 8) return null;
  const adv = top8.map((r) => r.groupLetter.toUpperCase());
  if (new Set(adv).size !== 8) return null;
  const assign = annexCAssignmentsForAdvancingThirds(adv);
  if (!assign) return null;
  const g = codeToThirdGroup(assign[annexIndex]);
  if (!g) return null;
  return thirdPlacedTeamFromGroup(g, tables, cellByAddress, fixtures);
}

function resolveR32Slot(
  slot: R32Slot,
  tables: Map<string, StandingRow[] | null>,
  cellByAddress: Map<string, WorkbookCell>,
  bestThird: BestThirdPlaceRow[] | null,
  fixtures: FixturesModel,
): string | null {
  if (isGroupRankSlot(slot)) {
    return teamAtGroupRank(slot.group, slot.rank, tables, cellByAddress, fixtures);
  }
  if (isBestThirdSlot(slot)) {
    return thirdTeamForAnnexIndex(slot.annex_index, bestThird, tables, cellByAddress, fixtures);
  }
  return null;
}

/** `side`: home = column B, away = column F. */
export function roundOf32TeamName(
  matchId: string,
  side: 'home' | 'away',
  tables: Map<string, StandingRow[] | null>,
  cellByAddress: Map<string, WorkbookCell>,
  bestThird: BestThirdPlaceRow[] | null,
  fixtures: FixturesModel,
): string | null {
  const m = r32MatchById(fixtures, matchId);
  if (!m) return null;
  const slot = side === 'home' ? m.home : m.away;
  return resolveR32Slot(slot, tables, cellByAddress, bestThird, fixtures);
}

/** Non-null if this address should show computed R32 team instead of mall formula. */
export function tryRoundOf32TeamDisplay(
  normalizedAddress: string,
  gridCol: number,
  tables: Map<string, StandingRow[] | null>,
  cellByAddress: Map<string, WorkbookCell>,
  bestThird: BestThirdPlaceRow[] | null,
  fixtures: FixturesModel,
): string | undefined {
  let row: number;
  try {
    row = parseA1(normalizedAddress).row;
  } catch {
    return undefined;
  }
  if (gridCol !== COL_B && gridCol !== COL_F) return undefined;
  const matchId = matchIdForExcelRow(row, fixtures);
  if (!matchId) return undefined;
  const side = gridCol === COL_B ? 'home' : 'away';
  const name = roundOf32TeamName(matchId, side, tables, cellByAddress, bestThird, fixtures);
  if (name === null) return '';
  return name;
}
