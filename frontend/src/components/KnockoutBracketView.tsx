import { useMemo } from 'react';
import { useTranslation } from 'react-i18next';
import type { i18n as I18nInstance } from 'i18next';
import type { WorkbookCell } from '../types/workbook';
import type { StandingRow } from '../lib/groupStandings';
import type { BestThirdPlaceRow } from '../lib/bestThirdPlace';
import { translateWorkbookString } from '../i18n/workbookStrings';
import {
  type BracketResolveCtx,
  bracketSelectOptions,
  bronzeTeamFromFinalistHalf,
  normalizeBracketAddress,
} from '../lib/knockoutBracket';
import type { FixturesModel } from '../lib/fixturesModel';

export type KnockoutBracketViewProps = {
  fixtures: FixturesModel;
  cellByAddress: Map<string, WorkbookCell>;
  scoreDrafts: Record<string, string>;
  standingTables: Map<string, StandingRow[] | null>;
  bestThirdPlaceRows: BestThirdPlaceRow[] | null;
  bracketDrafts: Record<string, string>;
  onBracketDraft: (normalizedAddress: string, value: string) => void;
};

function cellLabel(cellByAddress: Map<string, WorkbookCell>, addr: string): string {
  const v = cellByAddress.get(addr)?.value;
  if (v == null || v === '') return '';
  return String(v);
}

type GridItem = {
  col: number;
  row: number;
  address: string;
  align?: 'left' | 'right';
};

const LEFT_R16: GridItem[] = [
  { col: 0, row: 0, address: 'B145', align: 'left' },
  { col: 0, row: 2, address: 'B147', align: 'left' },
  { col: 0, row: 4, address: 'B149', align: 'left' },
  { col: 0, row: 6, address: 'B151', align: 'left' },
  { col: 0, row: 8, address: 'B153', align: 'left' },
  { col: 0, row: 10, address: 'B155', align: 'left' },
  { col: 0, row: 12, address: 'B157', align: 'left' },
  { col: 0, row: 14, address: 'B159', align: 'left' },
];

const LEFT_QF: GridItem[] = [
  { col: 1, row: 1, address: 'C146', align: 'left' },
  { col: 1, row: 5, address: 'C150', align: 'left' },
  { col: 1, row: 9, address: 'C154', align: 'left' },
  { col: 1, row: 13, address: 'C158', align: 'left' },
];

const LEFT_SF: GridItem[] = [
  { col: 2, row: 3, address: 'F148', align: 'left' },
  { col: 2, row: 11, address: 'F156', align: 'left' },
];

const RIGHT_R16: GridItem[] = [
  { col: 8, row: 0, address: 'R145', align: 'right' },
  { col: 8, row: 2, address: 'R147', align: 'right' },
  { col: 8, row: 4, address: 'R149', align: 'right' },
  { col: 8, row: 6, address: 'R151', align: 'right' },
  { col: 8, row: 8, address: 'R153', align: 'right' },
  { col: 8, row: 10, address: 'R155', align: 'right' },
  { col: 8, row: 12, address: 'R157', align: 'right' },
  { col: 8, row: 14, address: 'R159', align: 'right' },
];

const RIGHT_QF: GridItem[] = [
  { col: 7, row: 1, address: 'P146', align: 'right' },
  { col: 7, row: 5, address: 'P150', align: 'right' },
  { col: 7, row: 9, address: 'P154', align: 'right' },
  { col: 7, row: 13, address: 'P158', align: 'right' },
];

const RIGHT_SF: GridItem[] = [
  { col: 6, row: 3, address: 'N148', align: 'right' },
  { col: 6, row: 11, address: 'N156', align: 'right' },
];

const CENTER_ROW = 7;

function BracketPick({
  address,
  ctx,
  align,
  i18n,
  onBracketDraft,
}: {
  address: string;
  ctx: BracketResolveCtx;
  align: 'left' | 'right' | 'center';
  i18n: I18nInstance;
  onBracketDraft: (normalizedAddress: string, value: string) => void;
}) {
  const n = normalizeBracketAddress(address);
  const opts = useMemo(() => bracketSelectOptions(n, ctx), [
    n,
    ctx.fixtures,
    ctx.cellByAddress,
    ctx.scoreDrafts,
    ctx.standingTables,
    ctx.bestThird,
    ctx.bracketDrafts,
  ]);
  const raw = ctx.bracketDrafts[n]?.trim() ?? '';
  const value = raw && opts.includes(raw) ? raw : '';

  return (
    <select
      className={[
        'tipset-knockout__select',
        'tipset-knockout__select--playoff',
        align === 'right' ? 'tipset-knockout__select--right' : '',
        align === 'center' ? 'tipset-knockout__select--center' : '',
      ]
        .filter(Boolean)
        .join(' ')}
      aria-label={address}
      value={value}
      onChange={(e) => onBracketDraft(n, e.target.value)}
    >
      <option value=""> </option>
      {opts.map((o) => (
        <option key={o} value={o}>
          {translateWorkbookString(o, i18n)}
        </option>
      ))}
    </select>
  );
}

export function KnockoutBracketView({
  fixtures,
  cellByAddress,
  scoreDrafts,
  standingTables,
  bestThirdPlaceRows,
  bracketDrafts,
  onBracketDraft,
}: KnockoutBracketViewProps) {
  const { i18n, t } = useTranslation('app');

  const ctx: BracketResolveCtx = useMemo(
    () => ({
      fixtures,
      cellByAddress,
      scoreDrafts,
      standingTables,
      bestThird: bestThirdPlaceRows,
      bracketDrafts,
    }),
    [fixtures, cellByAddress, scoreDrafts, standingTables, bestThirdPlaceRows, bracketDrafts],
  );

  const headerAddrs = [
    ['B143', 'B144'],
    ['C143', 'C144'],
    ['F143', 'F144'],
    ['G143', 'G144'],
    ['J143', 'J144'],
    ['L143', 'L144'],
    ['N143', 'N144'],
    ['P143', 'P144'],
    ['R143', 'R144'],
  ] as const;

  const b169Entry = cellByAddress.get('B169');
  const goalKingOptions =
    b169Entry?.validation?.type === 'list' && Array.isArray(b169Entry.validation.list_values)
      ? b169Entry.validation.list_values.map((s) => String(s)).filter((s) => !s.startsWith('='))
      : [];

  const b169Val = bracketDrafts.B169?.trim() ?? '';
  const b169Select = b169Val && goalKingOptions.includes(b169Val) ? b169Val : '';

  const k169Raw = bracketDrafts.K169?.trim() ?? '';

  const penaltiesEntry = cellByAddress.get('B174');
  const penaltiesOptions =
    penaltiesEntry?.validation?.type === 'list' && Array.isArray(penaltiesEntry.validation.list_values)
      ? penaltiesEntry.validation.list_values
          .map((s) => String(s))
          .filter((s) => !s.startsWith('='))
      : ['Ja', 'Nej'];
  const b174Raw = bracketDrafts.B174?.trim() ?? '';
  const b174Select = b174Raw && penaltiesOptions.includes(b174Raw) ? b174Raw : '';

  const bronzeLeft = useMemo(() => {
    const b = fixtures.bracket;
    const [sfA, sfB] = b.finalist_left_semis;
    return bronzeTeamFromFinalistHalf(b.finalist_picks.left, sfA, sfB, ctx);
  }, [ctx, fixtures.bracket]);

  const bronzeRight = useMemo(() => {
    const b = fixtures.bracket;
    const [sfA, sfB] = b.finalist_right_semis;
    return bronzeTeamFromFinalistHalf(b.finalist_picks.right, sfA, sfB, ctx);
  }, [ctx, fixtures.bracket]);

  return (
    <section className="tipset-knockout" aria-label={t('knockout.sectionAria')}>
      <div className="tipset-knockout__banner">
        <h2 className="tipset-knockout__title">
          {translateWorkbookString(cellLabel(cellByAddress, 'B143') || 'Slutspel', i18n)}
        </h2>
        <p className="tipset-knockout__hint">{t('knockout.r16Hint')}</p>
      </div>

      <div className="tipset-knockout__grid">
        {headerAddrs.map(([titleAddr, ptsAddr], i) => (
          <div
            key={titleAddr}
            className="tipset-knockout__head"
            style={{ gridColumn: i + 1, gridRow: 1 }}
          >
            <div className="tipset-knockout__head-title">
              {translateWorkbookString(cellLabel(cellByAddress, titleAddr), i18n)}
            </div>
            <div className="tipset-knockout__head-pts">
              {translateWorkbookString(cellLabel(cellByAddress, ptsAddr), i18n)}
            </div>
          </div>
        ))}

        {LEFT_R16.map((it) => (
          <div
            key={it.address}
            className="tipset-knockout__cell"
            style={{ gridColumn: it.col + 1, gridRow: it.row + 2, justifySelf: 'stretch' }}
          >
            <BracketPick
              address={it.address}
              ctx={ctx}
              align={it.align === 'right' ? 'right' : 'left'}
              i18n={i18n}
              onBracketDraft={onBracketDraft}
            />
          </div>
        ))}
        {LEFT_QF.map((it) => (
          <div
            key={it.address}
            className="tipset-knockout__cell"
            style={{ gridColumn: it.col + 1, gridRow: it.row + 2, justifySelf: 'stretch' }}
          >
            <BracketPick address={it.address} ctx={ctx} align="left" i18n={i18n} onBracketDraft={onBracketDraft} />
          </div>
        ))}
        {LEFT_SF.map((it) => (
          <div
            key={it.address}
            className="tipset-knockout__cell"
            style={{ gridColumn: it.col + 1, gridRow: it.row + 2, justifySelf: 'stretch' }}
          >
            <BracketPick address={it.address} ctx={ctx} align="left" i18n={i18n} onBracketDraft={onBracketDraft} />
          </div>
        ))}

        <div
          className="tipset-knockout__cell tipset-knockout__cell--finalist"
          style={{ gridColumn: 4, gridRow: CENTER_ROW + 2, justifySelf: 'stretch' }}
        >
          <BracketPick address="G152" ctx={ctx} align="center" i18n={i18n} onBracketDraft={onBracketDraft} />
        </div>
        <div
          className="tipset-knockout__cell tipset-knockout__cell--champion"
          style={{ gridColumn: 5, gridRow: CENTER_ROW + 2, justifySelf: 'stretch' }}
        >
          <BracketPick address="J152" ctx={ctx} align="center" i18n={i18n} onBracketDraft={onBracketDraft} />
        </div>
        <div
          className="tipset-knockout__cell tipset-knockout__cell--finalist"
          style={{ gridColumn: 6, gridRow: CENTER_ROW + 2, justifySelf: 'stretch' }}
        >
          <BracketPick address="L152" ctx={ctx} align="center" i18n={i18n} onBracketDraft={onBracketDraft} />
        </div>

        {RIGHT_SF.map((it) => (
          <div
            key={it.address}
            className="tipset-knockout__cell"
            style={{ gridColumn: it.col + 1, gridRow: it.row + 2, justifySelf: 'stretch' }}
          >
            <BracketPick address={it.address} ctx={ctx} align="right" i18n={i18n} onBracketDraft={onBracketDraft} />
          </div>
        ))}
        {RIGHT_QF.map((it) => (
          <div
            key={it.address}
            className="tipset-knockout__cell"
            style={{ gridColumn: it.col + 1, gridRow: it.row + 2, justifySelf: 'stretch' }}
          >
            <BracketPick address={it.address} ctx={ctx} align="right" i18n={i18n} onBracketDraft={onBracketDraft} />
          </div>
        ))}
        {RIGHT_R16.map((it) => (
          <div
            key={it.address}
            className="tipset-knockout__cell"
            style={{ gridColumn: it.col + 1, gridRow: it.row + 2, justifySelf: 'stretch' }}
          >
            <BracketPick address={it.address} ctx={ctx} align="right" i18n={i18n} onBracketDraft={onBracketDraft} />
          </div>
        ))}
      </div>

      <div className="tipset-knockout__bronze">
        <div className="tipset-knockout__bronze-head">
          <div className="tipset-knockout__bronze-title">
            {translateWorkbookString(cellLabel(cellByAddress, 'G159'), i18n)}
          </div>
          <div className="tipset-knockout__bronze-pts">
            {translateWorkbookString(cellLabel(cellByAddress, 'G160'), i18n)}
          </div>
        </div>
        <div className="tipset-knockout__bronze-picks">
          <div className="tipset-knockout__bronze-team" aria-label="G161">
            {bronzeLeft ? translateWorkbookString(bronzeLeft, i18n) : '\u00a0'}
          </div>
          <div className="tipset-knockout__bronze-team" aria-label="L161">
            {bronzeRight ? translateWorkbookString(bronzeRight, i18n) : '\u00a0'}
          </div>
        </div>
      </div>

      <div className="tipset-knockout__extras">
        <div className="tipset-knockout__extra">
          <div className="tipset-knockout__extra-pts">
            {translateWorkbookString(cellLabel(cellByAddress, 'B167'), i18n)}
          </div>
          <div className="tipset-knockout__extra-label">
            {translateWorkbookString(cellLabel(cellByAddress, 'B168'), i18n)}
          </div>
          <select
            className="tipset-knockout__select tipset-knockout__select--playoff tipset-knockout__select--wide"
            aria-label="B169"
            value={b169Select}
            onChange={(e) => onBracketDraft('B169', e.target.value)}
          >
            <option value=""> </option>
            {goalKingOptions.map((o) => (
              <option key={o} value={o}>
                {translateWorkbookString(o, i18n)}
              </option>
            ))}
          </select>
        </div>
        <div className="tipset-knockout__extra">
          <div className="tipset-knockout__extra-pts">
            {translateWorkbookString(cellLabel(cellByAddress, 'K167'), i18n)}
          </div>
          <div className="tipset-knockout__extra-label">
            {translateWorkbookString(cellLabel(cellByAddress, 'K168'), i18n)}
          </div>
          <input
            type="text"
            className="tipset-knockout__text"
            aria-label="K169"
            autoComplete="off"
            value={k169Raw}
            onChange={(e) => onBracketDraft('K169', e.target.value)}
          />
        </div>
        <div className="tipset-knockout__extra">
          <div className="tipset-knockout__extra-pts">
            {translateWorkbookString(cellLabel(cellByAddress, 'B172'), i18n)}
          </div>
          <div className="tipset-knockout__extra-label">
            {translateWorkbookString(cellLabel(cellByAddress, 'B173'), i18n)}
          </div>
          <select
            className="tipset-knockout__select tipset-knockout__select--playoff tipset-knockout__select--wide"
            aria-label="B174"
            value={b174Select}
            onChange={(e) => onBracketDraft('B174', e.target.value)}
          >
            <option value=""> </option>
            {penaltiesOptions.map((o) => (
              <option key={o} value={o}>
                {translateWorkbookString(o, i18n)}
              </option>
            ))}
          </select>
        </div>
      </div>
    </section>
  );
}
