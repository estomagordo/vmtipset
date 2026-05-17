import type { WorkbookCell } from '../types/workbook';
import type { StandingRow } from './groupStandings';
import type { BestThirdPlaceRow } from './bestThirdPlace';
import { R32_FIRST_MATCH_ROW, roundOf32TeamName } from './roundOf32';
import { matchPickFromRow } from './match1x2';

export function normalizeBracketAddress(address: string): string {
  return address.trim().toUpperCase().replace(/\$/g, '');
}

/** Excel row for round-of-32 match M73 … M88 (VM-tipset mall). */
export function excelRowForR32Match(matchId: string): number | null {
  const n = /^M(\d+)$/.exec(matchId.trim().toUpperCase());
  if (!n) return null;
  const m = parseInt(n[1], 10);
  if (m < 73 || m > 88) return null;
  return R32_FIRST_MATCH_ROW + (m - 73);
}

/**
 * R16 feeder pairings per 2026 FIFA bracket (see Wikipedia "Round of 16" table).
 * Column B lists M97–M98 path first (rows 145–151), then M99–M100 path (153–159).
 * Column R mirrors the other side of the draw for the symmetric mall layout.
 */
export const R16_PAIR_BY_ADDRESS: Record<string, readonly [string, string]> = {
  B145: ['M73', 'M75'],
  B147: ['M74', 'M77'],
  B149: ['M83', 'M84'],
  B151: ['M81', 'M82'],
  B153: ['M76', 'M78'],
  B155: ['M79', 'M80'],
  B157: ['M86', 'M88'],
  B159: ['M85', 'M87'],
  R145: ['M76', 'M78'],
  R147: ['M79', 'M80'],
  R149: ['M86', 'M88'],
  R151: ['M85', 'M87'],
  R153: ['M73', 'M75'],
  R155: ['M74', 'M77'],
  R157: ['M83', 'M84'],
  R159: ['M81', 'M82'],
};

export type BracketResolveCtx = {
  cellByAddress: Map<string, WorkbookCell>;
  scoreDrafts: Record<string, string>;
  standingTables: Map<string, StandingRow[] | null>;
  bestThird: BestThirdPlaceRow[] | null;
  bracketDrafts: Record<string, string>;
};

const QF_LEFT: Record<string, readonly [string, string]> = {
  C146: ['B145', 'B147'],
  C150: ['B149', 'B151'],
  C154: ['B153', 'B155'],
  C158: ['B157', 'B159'],
};

const QF_RIGHT: Record<string, readonly [string, string]> = {
  P146: ['R145', 'R147'],
  P150: ['R149', 'R151'],
  P154: ['R153', 'R155'],
  P158: ['R157', 'R159'],
};

const SF_LEFT: Record<string, readonly [string, string]> = {
  F148: ['C146', 'C150'],
  F156: ['C154', 'C158'],
};

const SF_RIGHT: Record<string, readonly [string, string]> = {
  N148: ['P146', 'P150'],
  N156: ['P154', 'P158'],
};

/** Semifinal winners that feed the left/right finalist pick (G152 / L152). */
const FINALIST_LEFT_SEMIS: readonly [string, string] = ['F148', 'F156'];
const FINALIST_RIGHT_SEMIS: readonly [string, string] = ['N148', 'N156'];

const FINALIST_LEFT = 'G152';
const FINALIST_RIGHT = 'L152';
const CHAMPION = 'J152';

function draftNorm(ctx: BracketResolveCtx, address: string): string {
  const n = normalizeBracketAddress(address);
  return ctx.bracketDrafts[n]?.trim() ?? '';
}

function participantsFromR32(
  matchId: string,
  ctx: BracketResolveCtx,
): { home: string | null; away: string | null } {
  const h = roundOf32TeamName(matchId, 'home', ctx.standingTables, ctx.cellByAddress, ctx.bestThird);
  const a = roundOf32TeamName(matchId, 'away', ctx.standingTables, ctx.cellByAddress, ctx.bestThird);
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
  const row = excelRowForR32Match(matchId);
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
  const pair = R16_PAIR_BY_ADDRESS[n];
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

  const feeds = QF_LEFT[n] ?? QF_RIGHT[n];
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
  const sf = SF_LEFT[n] ?? SF_RIGHT[n];
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
  const sf = side === 'left' ? FINALIST_LEFT_SEMIS : FINALIST_RIGHT_SEMIS;
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
  const l = pickOrResolvedBracket(ctx, FINALIST_LEFT);
  const r = pickOrResolvedBracket(ctx, FINALIST_RIGHT);
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

  if (R16_PAIR_BY_ADDRESS[n]) {
    return r16SelectOptions(n, ctx);
  }
  if (QF_LEFT[n] || QF_RIGHT[n]) {
    return qfSelectOptions(n, ctx, new Map());
  }
  if (SF_LEFT[n] || SF_RIGHT[n]) {
    return sfSelectOptions(n, ctx);
  }
  if (n === FINALIST_LEFT) {
    return finalistOptions('left', ctx);
  }
  if (n === FINALIST_RIGHT) {
    return finalistOptions('right', ctx);
  }
  if (n === CHAMPION) {
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

export const KNOCKOUT_BRACKET_PICK_ADDRESSES: readonly string[] = [
  ...Object.keys(R16_PAIR_BY_ADDRESS),
  ...Object.keys(QF_LEFT),
  ...Object.keys(QF_RIGHT),
  ...Object.keys(SF_LEFT),
  ...Object.keys(SF_RIGHT),
  FINALIST_LEFT,
  FINALIST_RIGHT,
  CHAMPION,
];
