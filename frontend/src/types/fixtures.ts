/** Slot for a team in round of 32 (from fixtures.json). */
export type GroupRankSlot = {
  type: 'group_rank';
  group: string;
  /** 1 = group winner, 2 = runner-up */
  rank: number;
};

export type BestThirdSlot = {
  type: 'best_third';
  /** Index into Annex C assignment for this match (0–7). */
  annex_index: number;
};

export type R32Slot = GroupRankSlot | BestThirdSlot;

export type GroupFixture = {
  match_id: string;
  excel_row: number;
  home_team: string | null;
  away_team: string | null;
  score_cells: { home: string; away: string };
};

export type GroupBlock = {
  id: string;
  letter: string;
  header_row: number;
  start_row: number;
  teams: string[];
  fixtures: GroupFixture[];
};

export type R32Match = {
  match_id: string;
  excel_row: number;
  slot_hint: string | null;
  home: R32Slot;
  away: R32Slot;
};

export type KnockoutBracketSpec = {
  r16: Record<string, [string, string]>;
  qf: Record<string, [string, string]>;
  sf: Record<string, [string, string]>;
  finalist_left_semis: [string, string];
  finalist_right_semis: [string, string];
  finalist_picks: { left: string; right: string };
  champion: string;
};

export type FixturesDump = {
  version: number;
  source_workbook?: string;
  layout: {
    group_stage: {
      first_fixture_row: number;
      row_stride: number;
      fixture_rows_per_group: number;
      standing_rows_per_group: number;
      letters: string[];
    };
    round_of_32: {
      header_row: number;
      first_match_row: number;
      last_match_row: number;
    };
  };
  groups: GroupBlock[];
  round_of_32: R32Match[];
  knockout_bracket: KnockoutBracketSpec;
};
