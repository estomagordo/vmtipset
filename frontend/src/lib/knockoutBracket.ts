import type { WorkbookCell } from '../types/workbook';
import type { StandingRow } from './groupStandings';
import type { BestThirdPlaceRow } from './bestThirdPlace';
import type { FixturesModel } from './fixturesModel';
import { r32MatchById } from './fixturesModel';
import { roundOf32TeamName } from './roundOf32';
import { matchPickFromRow } from './match1x2';

export function normalizeBracketAddress(address: string): string {
  return address.trim().toUpperCase().replace(/\$/g, '');
}

/** Excel row for round-of-32 match M73 … M88 (VM-tipset mall). */
export function excelRowForR32Match(matchId: string, fixtures: FixturesModel): number | null {
  return r32MatchById(fixtures, matchId)?.excel_row ?? null;
}

export type BracketResolveCtx = {
  fixtures: FixturesModel;
  cellByAddress: Map<string, WorkbookCell>;
  scoreDrafts: Record<string, string>;
  standingTables: Map<string, StandingRow[] | null>;
  bestThird: BestThirdPlaceRow[] | null;
  bracketDrafts: Record<string, string>;
};

function draftNorm(ctx: BracketResolveCtx, address: string): string {
  const n = normalizeBracketAddress(address);
  return ctx.bracketDrafts[n]?.trim() ?? '';
}

function participantsFromR32(
  matchId: string,
  ctx: BracketResolveCtx,
): { home: string | null; away: string | null } {
  const h = roundOf32TeamName(
    matchId,
    'home',
    ctx.standingTables,
    ctx.cellByAddress,
    ctx.bestThird,
    ctx.fixtures,
  );
  const a = roundOf32TeamName(
    matchId,
    'away',
    ctx.standingTables,
    ctx.cellByAddress,
    ctx.bestThird,
    ctx.fixtures,
  );
  return { home: h, away: a };
}

function standingRowForTeamName(
  tables: Map<string, StandingRow[] | null>,
  teamName: string,
): StandingRow | null {
  const needle = teamName.trim();
  if (!needle) return null;
  for (const rows of tables.values()) {
    if (!rows) continue;
    const hit = rows.find((r) => r.name.trim() === needle);
    if (hit) return hit;
  }
  return null;
}

/**
 * When R32 score cells are not used (mall shows "―" in C), pick the winner from predicted
 * group-stage points, then goal difference, then goals scored; stable name order as last tiebreak.
 */
function r32WinnerFromStandingsWhenNoScorePick(
  homeName: string | null,
  awayName: string | null,
  tables: Map<string, StandingRow[] | null>,
): string | null {
  if (!homeName || !awayName) return null;
  const h = standingRowForTeamName(tables, homeName);
  const a = standingRowForTeamName(tables, awayName);
  if (!h || !a) return null;
  if (h.pts !== a.pts) return h.pts > a.pts ? homeName : awayName;
  if (h.gd !== a.gd) return h.gd > a.gd ? homeName : awayName;
  if (h.gf !== a.gf) return h.gf > a.gf ? homeName : awayName;
  return homeName.localeCompare(awayName, 'sv') <= 0 ? homeName : awayName;
}

/**
 * Predicted R32 winner: explicit C/E goals when the mall exposes 0–9 lists on that row;
 * otherwise inferred from group-stage tables (VM-tipset mall uses "―" in column C for R32).
 */
export function predictedR32Winner(matchId: string, ctx: BracketResolveCtx): string | null {
  const row = excelRowForR32Match(matchId, ctx.fixtures);
  if (row === null) return null;
  const { home, away } = participantsFromR32(matchId, ctx);

  const pick = matchPickFromRow(row, ctx.cellByAddress, ctx.scoreDrafts);
  if (pick === 'home') return home;
  if (pick === 'away') return away;
  if (pick === 'draw') return null;

  return r32WinnerFromStandingsWhenNoScorePick(home, away, ctx.standingTables);
}

/**
 * Åttondelsfinal: each cell is winner(Ma) vs winner(Mb). Options are the two predicted R32
 * winners (scores on the R32 row if present, else inferred from group-stage tables).
 */
export function r16SelectOptions(address: string, ctx: BracketResolveCtx): string[] {
  const n = normalizeBracketAddress(address);
  const pair = ctx.fixtures.bracket.r16[n];
  if (!pair) return [];
  const [m1, m2] = pair;
  const w1 = predictedR32Winner(m1, ctx);
  const w2 = predictedR32Winner(m2, ctx);
  if (!w1 || !w2) {
    return [];
  }
  if (w1 === w2) {
    return [w1];
  }
  return [w1, w2];
}

function qfSelectOptions(address: string, ctx: BracketResolveCtx, memo: Map<string, string[]>): string[] {
  const n = normalizeBracketAddress(address);
  if (memo.has(n)) return memo.get(n)!;

  const feeds = ctx.fixtures.bracket.qf[n];
  if (!feeds) {
    memo.set(n, []);
    return [];
  }

  const out = new Set<string>();
  for (const f of feeds) {
    const d = draftNorm(ctx, f);
    if (d) {
      out.add(d);
      continue;
    }
    for (const t of r16SelectOptions(f, ctx)) {
      out.add(t);
    }
  }
  const arr = [...out];
  memo.set(n, arr);
  return arr;
}

function sfSelectOptions(address: string, ctx: BracketResolveCtx): string[] {
  const n = normalizeBracketAddress(address);
  const sf = ctx.fixtures.bracket.sf[n];
  if (!sf) return [];
  const out = new Set<string>();
  for (const f of sf) {
    const d = draftNorm(ctx, f);
    if (d) {
      out.add(d);
      continue;
    }
    const memo = new Map<string, string[]>();
    for (const t of qfSelectOptions(f, ctx, memo)) {
      out.add(t);
    }
  }
  return [...out];
}

function finalistOptions(side: 'left' | 'right', ctx: BracketResolveCtx): string[] {
  const sf =
    side === 'left' ? ctx.fixtures.bracket.finalist_left_semis : ctx.fixtures.bracket.finalist_right_semis;
  const out = new Set<string>();
  for (const f of sf) {
    const d = draftNorm(ctx, f);
    if (d) {
      out.add(d);
      continue;
    }
    for (const t of sfSelectOptions(f, ctx)) {
      out.add(t);
    }
  }
  return [...out];
}

function championOptions(ctx: BracketResolveCtx): string[] {
  const { left, right } = ctx.fixtures.bracket.finalist_picks;
  const l = pickOrResolvedBracket(ctx, left);
  const r = pickOrResolvedBracket(ctx, right);
  if (!l || !r) {
    return [];
  }
  return l === r ? [l] : [l, r];
}

function pickOrResolvedBracket(ctx: BracketResolveCtx, addr: string): string | null {
  const d = draftNorm(ctx, addr);
  if (d) return d;
  return resolvedBracketPick(addr, ctx);
}

/** All non-empty pick options for a knockout list cell. */
export function bracketSelectOptions(address: string, ctx: BracketResolveCtx): string[] {
  const n = normalizeBracketAddress(address);

  const b = ctx.fixtures.bracket;
  if (b.r16[n]) {
    return r16SelectOptions(n, ctx);
  }
  if (b.qf[n]) {
    return qfSelectOptions(n, ctx, new Map());
  }
  if (b.sf[n]) {
    return sfSelectOptions(n, ctx);
  }
  if (n === b.finalist_picks.left) {
    return finalistOptions('left', ctx);
  }
  if (n === b.finalist_picks.right) {
    return finalistOptions('right', ctx);
  }
  if (n === b.champion) {
    return championOptions(ctx);
  }
  return [];
}

/** When the user has not picked, auto-select if exactly one option remains. */
export function resolvedBracketPick(address: string, ctx: BracketResolveCtx): string | null {
  const n = normalizeBracketAddress(address);
  const d = draftNorm(ctx, n);
  if (d) return d;
  const opts = bracketSelectOptions(n, ctx);
  if (opts.length === 1) return opts[0] ?? null;
  return null;
}

/**
 * Bronze opponent from one bracket half: the semifinal winner (F148/F156 or N148/N156) that was
 * not picked as finalist in G152/L152. Requires finalist pick and both semi picks (or unique resolution).
 */
export function bronzeTeamFromFinalistHalf(
  finalistCell: string,
  sfA: string,
  sfB: string,
  ctx: BracketResolveCtx,
): string | null {
  const w = draftNorm(ctx, finalistCell);
  if (!w) return null;
  const a = pickOrResolvedBracket(ctx, sfA);
  const b = pickOrResolvedBracket(ctx, sfB);
  if (!a || !b) return null;
  if (w === a) return b;
  if (w === b) return a;
  return null;
}

export function knockoutBracketPickAddresses(fixtures: FixturesModel): readonly string[] {
  const b = fixtures.bracket;
  return [
    ...Object.keys(b.r16),
    ...Object.keys(b.qf),
    ...Object.keys(b.sf),
    b.finalist_picks.left,
    b.finalist_picks.right,
    b.champion,
  ];
}
