import { GROUP_STAGE_STANDING_ROWS } from '../config/groupStageBlocks';
import type {
  FixturesDump,
  GroupBlock,
  GroupFixture,
  KnockoutBracketSpec,
  R32Match,
  R32Slot,
} from '../types/fixtures';
import { lettersToColIndex } from './excelAddress';

const COL_A = lettersToColIndex('A');
const COL_B = lettersToColIndex('B');
const COL_F = lettersToColIndex('F');
const COL_K = lettersToColIndex('K');

export type FixturesModel = {
  dump: FixturesDump;
  groupsById: Map<string, GroupBlock>;
  groupsByLetter: Map<string, GroupBlock>;
  groupFixtureByExcelRow: Map<number, GroupFixture>;
  groupBlockByStandingRow: Map<number, GroupBlock>;
  r32ByMatchId: Map<string, R32Match>;
  r32ByExcelRow: Map<number, R32Match>;
  bracket: KnockoutBracketSpec;
};

export function buildFixturesModel(dump: FixturesDump): FixturesModel {
  const groupsById = new Map<string, GroupBlock>();
  const groupsByLetter = new Map<string, GroupBlock>();
  const groupFixtureByExcelRow = new Map<number, GroupFixture>();
  const groupBlockByStandingRow = new Map<number, GroupBlock>();

  for (const g of dump.groups) {
    groupsById.set(g.id, g);
    groupsByLetter.set(g.letter.toUpperCase(), g);
    for (const fx of g.fixtures) {
      groupFixtureByExcelRow.set(fx.excel_row, fx);
    }
    for (let i = 0; i < GROUP_STAGE_STANDING_ROWS; i++) {
      groupBlockByStandingRow.set(g.start_row + i, g);
    }
  }

  const r32ByMatchId = new Map<string, R32Match>();
  const r32ByExcelRow = new Map<number, R32Match>();
  for (const m of dump.round_of_32) {
    r32ByMatchId.set(m.match_id.toUpperCase(), m);
    r32ByExcelRow.set(m.excel_row, m);
  }

  return {
    dump,
    groupsById,
    groupsByLetter,
    groupFixtureByExcelRow,
    groupBlockByStandingRow,
    r32ByMatchId,
    r32ByExcelRow,
    bracket: dump.knockout_bracket,
  };
}

export function groupBlockForLetter(model: FixturesModel, letter: string): GroupBlock | undefined {
  return model.groupsByLetter.get(letter.trim().toUpperCase());
}

export function groupFixtures(block: GroupBlock): readonly GroupFixture[] {
  return block.fixtures;
}

export function r32MatchForExcelRow(model: FixturesModel, excelRow: number): R32Match | undefined {
  return model.r32ByExcelRow.get(excelRow);
}

export function r32MatchById(model: FixturesModel, matchId: string): R32Match | undefined {
  return model.r32ByMatchId.get(matchId.trim().toUpperCase());
}

export function isGroupRankSlot(slot: R32Slot): slot is Extract<R32Slot, { type: 'group_rank' }> {
  return slot.type === 'group_rank';
}

export function isBestThirdSlot(slot: R32Slot): slot is Extract<R32Slot, { type: 'best_third' }> {
  return slot.type === 'best_third';
}

/** Group-stage fixture home (B) or away (F) team name; `undefined` if not a fixture team cell. */
export function groupFixtureTeamName(
  model: FixturesModel,
  excelRow: number,
  gridCol: number,
): string | undefined {
  if (gridCol !== COL_B && gridCol !== COL_F) return undefined;
  const fx = model.groupFixtureByExcelRow.get(excelRow);
  if (!fx) return undefined;
  const name = gridCol === COL_B ? fx.home_team : fx.away_team;
  const trimmed = name?.trim();
  return trimmed || undefined;
}

/** Match id label in column A on a fixture row. */
export function groupFixtureMatchId(model: FixturesModel, excelRow: number): string | undefined {
  const fx = model.groupFixtureByExcelRow.get(excelRow);
  return fx?.match_id?.trim() || undefined;
}

/** Standing-table team name (column K) before live standings exist. */
export function standingTeamNameAtRow(model: FixturesModel, excelRow: number): string | undefined {
  const block = model.groupBlockByStandingRow.get(excelRow);
  if (!block) return undefined;
  const rank0 = excelRow - block.start_row;
  if (rank0 < 0 || rank0 >= block.teams.length) return undefined;
  return block.teams[rank0]?.trim() || undefined;
}

export function isGroupStandingNameCell(model: FixturesModel, excelRow: number, gridCol: number): boolean {
  return gridCol === COL_K && model.groupBlockByStandingRow.has(excelRow);
}

export function isGroupFixtureTeamCell(model: FixturesModel, excelRow: number, gridCol: number): boolean {
  return groupFixtureTeamName(model, excelRow, gridCol) !== undefined;
}

export function isGroupFixtureMatchIdCell(model: FixturesModel, excelRow: number, gridCol: number): boolean {
  return gridCol === COL_A && model.groupFixtureByExcelRow.has(excelRow);
}
