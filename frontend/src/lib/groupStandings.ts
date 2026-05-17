import type { WorkbookCell } from '../types/workbook';
import { parseA1 } from './excelAddress';
import { getScoreSelectValue } from './scoreCells';

/**
 * Group-stage standings row (one team). Independent of Excel; values are what we show in the mall grid
 * columns K–S, which were authored as cross-sheet links in the template file.
 */
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

/** Parsed from template formula text only to learn *where* each visible cell maps in (group, rank, stat). */
const STANDINGS_LINK_RE = /^'Grupp ([A-Za-z]+)'!([A-Z]{1,3})(\d+)$/i;

export type StandingsLayoutRef = {
  groupId: string;
  /** Template column on the (conceptual) standings table: L = name, M = played, … T = points. */
  statColumn: string;
  /** Template row (7–10): encodes rank when we fix base row 7 = 1st place. */
  templateRow: number;
};

export function parseStandingsLayoutRef(formula: string | undefined): StandingsLayoutRef | null {
  if (!formula) return null;
  const inner = formula.trim().replace(/^=/, '').trim();
  const m = inner.match(STANDINGS_LINK_RE);
  if (!m) return null;
  const groupId = `Grupp ${m[1].toUpperCase()}`;
  const statColumn = m[2].toUpperCase();
  const templateRow = parseInt(m[3], 10);
  if (!Number.isFinite(templateRow)) return null;
  return { groupId, statColumn, templateRow };
}

/**
 * From the dumped grid, find the first fixture row of each group's 6-game block by reading
 * standings “link” cells in column K (layout metadata from the export, not runtime Excel).
 */
export function inferFixtureBlockStartRowsPerGroup(cellByAddress: Map<string, WorkbookCell>): Map<string, number> {
  const anchors = new Map<string, number>();

  for (const cell of cellByAddress.values()) {
    let col: number;
    let row: number;
    try {
      const p = parseA1(cell.address.replace(/\$/g, ''));
      row = p.row;
      col = p.col;
    } catch {
      continue;
    }
    if (col !== 11) continue;

    const ref = parseStandingsLayoutRef(cell.formula);
    if (!ref || ref.statColumn !== 'L' || ref.templateRow < 7) continue;

    const firstFixtureRow = row - (ref.templateRow - 7);
    if (firstFixtureRow < 1) continue;

    const prev = anchors.get(ref.groupId);
    if (prev === undefined || firstFixtureRow < prev) {
      anchors.set(ref.groupId, firstFixtureRow);
    }
  }

  return anchors;
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

function formatStatForGrid(row: StandingRow, statColumn: string): string | number {
  switch (statColumn) {
    case 'L':
      return row.name;
    case 'M':
      return row.played;
    case 'N':
      return row.wins;
    case 'O':
      return row.draws;
    case 'P':
      return row.losses;
    case 'Q':
      return row.gf;
    case 'R':
      return row.ga;
    case 'S':
      return row.gd;
    case 'T':
      return row.pts;
    default:
      return '';
  }
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
): Map<string, StandingRow[] | null> {
  const anchors = inferFixtureBlockStartRowsPerGroup(cellByAddress);
  const out = new Map<string, StandingRow[] | null>();

  for (const [groupId, firstRow] of anchors) {
    const fixtureRows = [firstRow, firstRow + 1, firstRow + 2, firstRow + 3, firstRow + 4, firstRow + 5];
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

/** Map a template “standings link” formula to a live computed value if that cell is part of a live table. */
export function resolveGroupStandingDisplay(
  formula: string | undefined,
  tables: Map<string, StandingRow[] | null>,
): StandingResolve {
  const ref = parseStandingsLayoutRef(formula);
  if (!ref) return { kind: 'none' };

  const rows = tables.get(ref.groupId);
  if (rows === null || rows === undefined) {
    return { kind: 'none' };
  }

  const rank = ref.templateRow - 7;
  if (rank < 0 || rank >= rows.length) {
    return { kind: 'none' };
  }

  const value = formatStatForGrid(rows[rank], ref.statColumn);
  return { kind: 'value', value };
}
