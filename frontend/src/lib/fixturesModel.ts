import type {
  FixturesDump,
  GroupBlock,
  GroupFixture,
  KnockoutBracketSpec,
  R32Match,
  R32Slot,
} from '../types/fixtures';

export type FixturesModel = {
  dump: FixturesDump;
  groupsById: Map<string, GroupBlock>;
  groupsByLetter: Map<string, GroupBlock>;
  r32ByMatchId: Map<string, R32Match>;
  r32ByExcelRow: Map<number, R32Match>;
  bracket: KnockoutBracketSpec;
};

export function buildFixturesModel(dump: FixturesDump): FixturesModel {
  const groupsById = new Map<string, GroupBlock>();
  const groupsByLetter = new Map<string, GroupBlock>();
  for (const g of dump.groups) {
    groupsById.set(g.id, g);
    groupsByLetter.set(g.letter.toUpperCase(), g);
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
