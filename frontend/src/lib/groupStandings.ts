import {
  GROUP_STAGE_BLOCKS,
  GROUP_STAGE_FIXTURE_ROWS,
  GROUP_STAGE_STANDING_ROWS,
  type GroupStageBlock,
} from '../config/groupStageBlocks';
import type { WorkbookCell } from '../types/workbook';
import { lettersToColIndex, parseA1 } from './excelAddress';
import { getScoreSelectValue } from './scoreCells';

/** One row in a group standings table. */
export type StandingRow = {
  name: string;
  played: number;
  wins: number;
  draws: number;
  losses: number;
  gf: number;
  ga: number;
  gd: number;
  pts: number;
};

export type StandingResolve =
  | { kind: 'none' }
  | { kind: 'value'; value: string | number };

type StandingStatKey = keyof StandingRow;

/** Mall columns K–S (11–19) → field on `StandingRow`. */
const MALL_COL_TO_STAT: Record<number, StandingStatKey> = {
  [lettersToColIndex('K')]: 'name',
  [lettersToColIndex('L')]: 'played',
  [lettersToColIndex('M')]: 'wins',
  [lettersToColIndex('N')]: 'draws',
  [lettersToColIndex('O')]: 'losses',
  [lettersToColIndex('P')]: 'gf',
  [lettersToColIndex('Q')]: 'ga',
  [lettersToColIndex('R')]: 'gd',
  [lettersToColIndex('S')]: 'pts',
};

function blockAndRankForStandingsCell(
  row: number,
  col: number,
  blocks: readonly GroupStageBlock[],
): { block: GroupStageBlock; rank: number; statKey: StandingStatKey } | null {
  const statKey = MALL_COL_TO_STAT[col];
  if (statKey === undefined) return null;

  for (const block of blocks) {
    const r0 = block.startRow;
    if (row >= r0 && row < r0 + GROUP_STAGE_STANDING_ROWS) {
      return { block, rank: row - r0, statKey };
    }
  }
  return null;
}

function teamNameFromFixtureCell(entry: WorkbookCell | undefined): string | null {
  if (!entry) return null;
  if (entry.kind === 'formula' && entry.cached_value != null && entry.cached_value !== '') {
    return String(entry.cached_value).trim();
  }
  if (entry.value != null && entry.value !== '') {
    return String(entry.value).trim();
  }
  return null;
}

/**
 * User prediction for goals: `null` if the score is blank (not played / not entered).
 * Zero is a valid score; do not treat empty string as 0.
 */
function readPredictedGoals(
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

type TeamAgg = { p: number; w: number; d: number; l: number; gf: number; ga: number };

function compareStandings(a: StandingRow, b: StandingRow): number {
  if (b.pts !== a.pts) return b.pts - a.pts;
  if (b.gd !== a.gd) return b.gd - a.gd;
  if (b.gf !== a.gf) return b.gf - a.gf;
  return a.name.localeCompare(b.name, 'sv');
}

/**
 * Build standings for every group from predicted scores. A fixture counts only when **both**
 * home and away predictions are non-null; otherwise it does not affect the table.
 *
 * Map value is `null` when there is no qualifying fixture yet — the UI should show the dump’s
 * static `cached_value` for that block.
 */
export function computeGroupStandings(
  cellByAddress: Map<string, WorkbookCell>,
  scoreDrafts: Record<string, string>,
  blocks: readonly GroupStageBlock[] = GROUP_STAGE_BLOCKS,
): Map<string, StandingRow[] | null> {
  const out = new Map<string, StandingRow[] | null>();

  for (const { id: groupId, startRow } of blocks) {
    const fixtureRows: number[] = [];
    for (let i = 0; i < GROUP_STAGE_FIXTURE_ROWS; i++) {
      fixtureRows.push(startRow + i);
    }

    const teams = new Set<string>();
    const pairings: { home: string; away: string; cAddr: string; eAddr: string }[] = [];

    for (const r of fixtureRows) {
      const cAddr = `C${r}`;
      const eAddr = `E${r}`;
      const b = cellByAddress.get(`B${r}`);
      const f = cellByAddress.get(`F${r}`);
      const home = teamNameFromFixtureCell(b);
      const away = teamNameFromFixtureCell(f);
      if (!home || !away) continue;
      teams.add(home);
      teams.add(away);
      pairings.push({ home, away, cAddr, eAddr });
    }

    let finishedFixtures = 0;
    const stats = new Map<string, TeamAgg>();
    for (const t of teams) {
      stats.set(t, { p: 0, w: 0, d: 0, l: 0, gf: 0, ga: 0 });
    }

    for (const m of pairings) {
      const homeGoals = readPredictedGoals(m.cAddr, cellByAddress.get(m.cAddr), scoreDrafts);
      const awayGoals = readPredictedGoals(m.eAddr, cellByAddress.get(m.eAddr), scoreDrafts);
      if (homeGoals === null || awayGoals === null) continue;

      finishedFixtures += 1;
      const stH = stats.get(m.home);
      const stA = stats.get(m.away);
      if (!stH || !stA) continue;

      stH.gf += homeGoals;
      stH.ga += awayGoals;
      stA.gf += awayGoals;
      stA.ga += homeGoals;
      stH.p += 1;
      stA.p += 1;

      if (homeGoals > awayGoals) {
        stH.w += 1;
        stA.l += 1;
      } else if (homeGoals < awayGoals) {
        stA.w += 1;
        stH.l += 1;
      } else {
        stH.d += 1;
        stA.d += 1;
      }
    }

    if (finishedFixtures === 0) {
      out.set(groupId, null);
      continue;
    }

    const rows: StandingRow[] = [...teams].map((name) => {
      const s = stats.get(name) ?? { p: 0, w: 0, d: 0, l: 0, gf: 0, ga: 0 };
      const pts = s.w * 3 + s.d;
      const gd = s.gf - s.ga;
      return {
        name,
        played: s.p,
        wins: s.w,
        draws: s.d,
        losses: s.l,
        gf: s.gf,
        ga: s.ga,
        gd,
        pts,
      };
    });

    rows.sort(compareStandings);
    out.set(groupId, rows);
  }

  return out;
}

/** If `address` is a K–S standings cell in a configured group block, return the live value. */
export function resolveStandingsCellAtAddress(
  address: string,
  tables: Map<string, StandingRow[] | null>,
  blocks: readonly GroupStageBlock[] = GROUP_STAGE_BLOCKS,
): StandingResolve {
  let row: number;
  let col: number;
  try {
    const p = parseA1(address.replace(/\$/g, ''));
    row = p.row;
    col = p.col;
  } catch {
    return { kind: 'none' };
  }

  const located = blockAndRankForStandingsCell(row, col, blocks);
  if (!located) return { kind: 'none' };

  const table = tables.get(located.block.id);
  if (table === null || table === undefined) return { kind: 'none' };

  const rank = located.rank;
  if (rank < 0 || rank >= table.length) return { kind: 'none' };

  const value = table[rank][located.statKey];
  return { kind: 'value', value };
}
