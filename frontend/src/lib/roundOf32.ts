import type { WorkbookCell } from '../types/workbook';
import { GROUP_STAGE_BLOCKS } from '../config/groupStageBlocks';
import { readStandingRowFromWorkbook, type StandingRow } from './groupStandings';
import { parseA1 } from './excelAddress';
import { annexCAssignmentsForAdvancingThirds } from '../data/annexC2026';
import type { BestThirdPlaceRow } from './bestThirdPlace';

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

/** Wikipedia Annex column order: 1Avs … 1Lvs → which Round-of-32 match gets that third-placed team. */
const MATCH_THIRD_ANNEX_INDEX: Record<string, number> = {
  M79: 0,
  M85: 1,
  M81: 2,
  M74: 3,
  M82: 4,
  M77: 5,
  M87: 6,
  M80: 7,
};

function matchIdForExcelRow(excelRow: number): string | null {
  if (excelRow < R32_FIRST_MATCH_ROW || excelRow > R32_LAST_MATCH_ROW) return null;
  return `M${73 + (excelRow - R32_FIRST_MATCH_ROW)}`;
}

function blockStartRow(groupLetter: string): number | null {
  const block = GROUP_STAGE_BLOCKS.find((b) => b.id === `Grupp ${groupLetter.toUpperCase()}`);
  return block ? block.startRow : null;
}

function teamAtGroupRank(
  groupLetter: string,
  rank0: number,
  tables: Map<string, StandingRow[] | null>,
  cellByAddress: Map<string, WorkbookCell>,
): string | null {
  const id = `Grupp ${groupLetter.toUpperCase()}`;
  const live = tables.get(id);
  if (live && live.length > rank0) {
    const n = live[rank0]?.name?.trim();
    return n || null;
  }
  const sr = blockStartRow(groupLetter);
  if (sr === null) return null;
  return readStandingRowFromWorkbook(cellByAddress, sr + rank0)?.name?.trim() ?? null;
}

function thirdPlacedTeamFromGroup(
  groupLetter: string,
  tables: Map<string, StandingRow[] | null>,
  cellByAddress: Map<string, WorkbookCell>,
): string | null {
  return teamAtGroupRank(groupLetter, 2, tables, cellByAddress);
}

function codeToThirdGroup(code: string): string | null {
  const m = code.trim().toUpperCase().match(/^3([A-L])$/);
  return m ? m[1] : null;
}

/**
 * When the eight advancing third-place groups are known, resolve the third-placed team
 * name for a given R32 match (M74, M77, …).
 */
function thirdTeamForMatch(
  matchId: string,
  bestThird: BestThirdPlaceRow[] | null,
  tables: Map<string, StandingRow[] | null>,
  cellByAddress: Map<string, WorkbookCell>,
): string | null {
  const idx = MATCH_THIRD_ANNEX_INDEX[matchId];
  if (idx === undefined) return null;
  if (!bestThird || bestThird.length < 8) return null;
  const top8 = bestThird.filter((r) => r.rank <= 8);
  if (top8.length !== 8) return null;
  const adv = top8.map((r) => r.groupLetter.toUpperCase());
  if (new Set(adv).size !== 8) return null;
  const assign = annexCAssignmentsForAdvancingThirds(adv);
  if (!assign) return null;
  const g = codeToThirdGroup(assign[idx]);
  if (!g) return null;
  return thirdPlacedTeamFromGroup(g, tables, cellByAddress);
}

/** `side`: home = column B, away = column F. */
export function roundOf32TeamName(
  matchId: string,
  side: 'home' | 'away',
  tables: Map<string, StandingRow[] | null>,
  cellByAddress: Map<string, WorkbookCell>,
  bestThird: BestThirdPlaceRow[] | null,
): string | null {
  switch (matchId) {
    case 'M73':
      return side === 'home' ? teamAtGroupRank('A', 1, tables, cellByAddress) : teamAtGroupRank('B', 1, tables, cellByAddress);
    case 'M74':
      return side === 'home'
        ? teamAtGroupRank('E', 0, tables, cellByAddress)
        : thirdTeamForMatch(matchId, bestThird, tables, cellByAddress);
    case 'M75':
      return side === 'home' ? teamAtGroupRank('F', 0, tables, cellByAddress) : teamAtGroupRank('C', 1, tables, cellByAddress);
    case 'M76':
      return side === 'home' ? teamAtGroupRank('C', 0, tables, cellByAddress) : teamAtGroupRank('F', 1, tables, cellByAddress);
    case 'M77':
      return side === 'home'
        ? teamAtGroupRank('I', 0, tables, cellByAddress)
        : thirdTeamForMatch(matchId, bestThird, tables, cellByAddress);
    case 'M78':
      return side === 'home' ? teamAtGroupRank('E', 1, tables, cellByAddress) : teamAtGroupRank('I', 1, tables, cellByAddress);
    case 'M79':
      return side === 'home'
        ? teamAtGroupRank('A', 0, tables, cellByAddress)
        : thirdTeamForMatch(matchId, bestThird, tables, cellByAddress);
    case 'M80':
      return side === 'home'
        ? teamAtGroupRank('L', 0, tables, cellByAddress)
        : thirdTeamForMatch(matchId, bestThird, tables, cellByAddress);
    case 'M81':
      return side === 'home'
        ? teamAtGroupRank('D', 0, tables, cellByAddress)
        : thirdTeamForMatch(matchId, bestThird, tables, cellByAddress);
    case 'M82':
      return side === 'home'
        ? teamAtGroupRank('G', 0, tables, cellByAddress)
        : thirdTeamForMatch(matchId, bestThird, tables, cellByAddress);
    case 'M83':
      return side === 'home' ? teamAtGroupRank('K', 1, tables, cellByAddress) : teamAtGroupRank('L', 1, tables, cellByAddress);
    case 'M84':
      return side === 'home' ? teamAtGroupRank('H', 0, tables, cellByAddress) : teamAtGroupRank('J', 1, tables, cellByAddress);
    case 'M85':
      return side === 'home'
        ? teamAtGroupRank('B', 0, tables, cellByAddress)
        : thirdTeamForMatch(matchId, bestThird, tables, cellByAddress);
    case 'M86':
      return side === 'home' ? teamAtGroupRank('J', 0, tables, cellByAddress) : teamAtGroupRank('H', 1, tables, cellByAddress);
    case 'M87':
      return side === 'home'
        ? teamAtGroupRank('K', 0, tables, cellByAddress)
        : thirdTeamForMatch(matchId, bestThird, tables, cellByAddress);
    case 'M88':
      return side === 'home' ? teamAtGroupRank('D', 1, tables, cellByAddress) : teamAtGroupRank('G', 1, tables, cellByAddress);
    default:
      return null;
  }
}

/** Non-null if this address should show computed R32 team instead of mall formula. */
export function tryRoundOf32TeamDisplay(
  normalizedAddress: string,
  gridCol: number,
  tables: Map<string, StandingRow[] | null>,
  cellByAddress: Map<string, WorkbookCell>,
  bestThird: BestThirdPlaceRow[] | null,
): string | undefined {
  let row: number;
  try {
    row = parseA1(normalizedAddress).row;
  } catch {
    return undefined;
  }
  if (gridCol !== COL_B && gridCol !== COL_F) return undefined;
  const matchId = matchIdForExcelRow(row);
  if (!matchId) return undefined;
  const side = gridCol === COL_B ? 'home' : 'away';
  const name = roundOf32TeamName(matchId, side, tables, cellByAddress, bestThird);
  if (name === null) return '';
  return name;
}
