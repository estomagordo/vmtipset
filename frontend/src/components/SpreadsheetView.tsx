import { Fragment, useCallback, useMemo, useState } from 'react';
import { useTranslation } from 'react-i18next';
import type { WorkbookCell } from '../types/workbook';
import type { GridCell } from '../lib/gridModel';
import {
  computeGroupStandings,
  resolveStandingsCellAtAddress,
  type StandingRow,
} from '../lib/groupStandings';
import {
  GROUP_STAGE_BLOCKS,
  GROUP_STAGE_BLOCK_DATA_ROWS,
  GROUP_STAGE_FIRST_HEADER_GRID_INDEX,
  GROUP_STAGE_ROW_STRIDE,
} from '../config/groupStageBlocks';
import {
  getScoreSelectValue,
  isScoreDigitListValidation,
  SCORE_OPTION_VALUES,
} from '../lib/scoreCells';
import { translateWorkbookString } from '../i18n/workbookStrings';
import { isMetaFreeTextCell } from '../lib/metaTextFields';
import {
  bestThirdPlaceDataRowIndex,
  computeBestThirdPlaceTable,
  getBestThirdPlaceCellValue,
  isBestThirdPlaceDataCell,
  isBestThirdPlaceHeaderCell,
  BEST_THIRD_PLACE_FIRST_GRID_ROW_INDEX,
  BEST_THIRD_PLACE_GRID_SLICE_END_EXCLUSIVE,
  type BestThirdPlaceRow,
} from '../lib/bestThirdPlace';
import { classifyGroupGridCell, GROUP_COL_WIDTH_PCT, GROUP_TABLE_COL_CLASSES, GROUP_STAGE_GRID_MATCH_FR, GROUP_STAGE_GRID_X_FR, GROUP_STAGE_GRID_STANDINGS_FR } from '../lib/groupGridStyle';
import { matchPickFromRow } from '../lib/match1x2';
import { lettersToColIndex } from '../lib/excelAddress';
import { tryRoundOf32TeamDisplay, R32_FIRST_MATCH_ROW, R32_LAST_MATCH_ROW, R32_HEADER_ROW, R32_FIRST_GRID_ROW_INDEX, R32_GRID_SLICE_END_EXCLUSIVE } from '../lib/roundOf32';
import { KnockoutBracketView } from './KnockoutBracketView';
import { normalizeBracketAddress } from '../lib/knockoutBracket';

export type SpreadsheetViewProps = {
  title: string;
  grid: GridCell[][];
  rowCount: number;
  colCount: number;
  cellByAddress: Map<string, WorkbookCell>;
};

type GroupRowCtx = { blockHeaderExcelRow: number; rowInBlock: number };

function normalizeCellAddress(address: string): string {
  return address.trim().toUpperCase().replace(/\$/g, '');
}

const GROUP_COLGROUP = (
  <colgroup>
    {GROUP_TABLE_COL_CLASSES.map((cls, i) => (
      <col key={i} className={cls} style={{ width: `${GROUP_COL_WIDTH_PCT[i]}%` }} />
    ))}
  </colgroup>
);

const COL_G = lettersToColIndex('G');
const COL_I = lettersToColIndex('I');

type CellBodyProps = {
  address: string;
  entry?: WorkbookCell;
  excelRow: number;
  gridCol: number;
  groupCtx: GroupRowCtx | null;
  cellByAddress: Map<string, WorkbookCell>;
  scoreDrafts: Record<string, string>;
  metaDrafts: Record<string, string>;
  onMetaDraft: (normalizedAddress: string, value: string) => void;
  onScoreDraft: (address: string, value: string) => void;
  standingTables: Map<string, StandingRow[] | null>;
  bestThirdPlaceRows: BestThirdPlaceRow[] | null;
};

function CellBody({
  address,
  entry,
  excelRow,
  gridCol,
  groupCtx,
  cellByAddress,
  scoreDrafts,
  metaDrafts,
  onMetaDraft,
  onScoreDraft,
  standingTables,
  bestThirdPlaceRows,
}: CellBodyProps) {
  const { t, i18n } = useTranslation('app');
  const key = normalizeCellAddress(address);

  const r32Display = tryRoundOf32TeamDisplay(
    key,
    gridCol,
    standingTables,
    cellByAddress,
    bestThirdPlaceRows,
  );
  if (r32Display !== undefined) {
    if (r32Display === '') {
      return <span className="tipset-cell-text">{'\u00a0'}</span>;
    }
    const text = translateWorkbookString(r32Display, i18n);
    return <span className="tipset-cell-text">{text}</span>;
  }

  if (
    groupCtx != null &&
    groupCtx.rowInBlock >= 1 &&
    gridCol >= COL_G &&
    gridCol <= COL_I &&
    entry?.kind === 'formula'
  ) {
    const pick = matchPickFromRow(excelRow, cellByAddress, scoreDrafts);
    const slot = pick === 'home' ? 0 : pick === 'draw' ? 1 : pick === 'away' ? 2 : -1;
    const colSlot = gridCol - COL_G;
    if (slot >= 0 && colSlot === slot) {
      return <span className="tipset-x-marker">X</span>;
    }
    return <span className="tipset-cell-1x2-empty">{'\u00a0'}</span>;
  }

  if (bestThirdPlaceRows && isBestThirdPlaceDataCell(excelRow, gridCol)) {
    const raw = getBestThirdPlaceCellValue(excelRow, gridCol, bestThirdPlaceRows);
    if (raw === null) {
      return <span className="tipset-cell-text">{'\u00a0'}</span>;
    }
    if (typeof raw === 'number') {
      return <span className="tipset-cell-num">{String(raw)}</span>;
    }
    const text = translateWorkbookString(raw, i18n);
    return <span className="tipset-cell-text">{text}</span>;
  }

  if (isMetaFreeTextCell(key)) {
    const fromDraft = Object.prototype.hasOwnProperty.call(metaDrafts, key);
    const initial =
      entry?.value != null && entry.value !== '' ? String(entry.value) : '';
    const value = fromDraft ? metaDrafts[key] : initial;
    const aria = key === 'G4' ? t('aria.playerName') : t('aria.playerEmail');
    return (
      <input
        type={key === 'G6' ? 'email' : 'text'}
        name={key === 'G4' ? 'playerName' : 'playerEmail'}
        className="tipset-text-input"
        value={value}
        onChange={(e) => onMetaDraft(key, e.target.value)}
        aria-label={aria}
        autoComplete={key === 'G6' ? 'email' : 'name'}
      />
    );
  }

  if (!entry) return null;

  if (entry.kind === 'formula') {
    const standing = resolveStandingsCellAtAddress(entry.address, standingTables, GROUP_STAGE_BLOCKS);
    if (standing.kind === 'value') {
      const v = standing.value;
      const isNum = typeof v === 'number';
      const text = isNum ? String(v) : translateWorkbookString(String(v), i18n);
      return (
        <span
          className={isNum ? 'tipset-cell-num' : 'tipset-cell-text'}
          title={entry.formula ? String(entry.formula) : undefined}
        >
          {text}
        </span>
      );
    }

    const cached = entry.cached_value;
    const hasCached = cached !== undefined && cached !== null;
    if (hasCached) {
      const isNum = typeof cached === 'number';
      const text = isNum ? String(cached) : translateWorkbookString(String(cached), i18n);
      return (
        <span
          className={isNum ? 'tipset-cell-num' : 'tipset-cell-text'}
          title={entry.formula ? String(entry.formula) : undefined}
        >
          {text}
        </span>
      );
    }
    return (
      <span
        className="tipset-cell-text"
        title={
          entry.formula ? t('formula.tooltipNotCached', { formula: entry.formula }) : undefined
        }
      >
        {'\u00a0'}
      </span>
    );
  }

  if (isScoreDigitListValidation(entry.validation)) {
    const value = getScoreSelectValue(entry.address, entry, scoreDrafts);
    return (
      <select
        className="tipset-select tipset-select--in-cell"
        aria-label={t('aria.scorePicker', { address: entry.address })}
        value={value}
        onChange={(e) => onScoreDraft(entry.address, e.target.value)}
      >
        <option value="" />
        {SCORE_OPTION_VALUES.map((n) => (
          <option key={n} value={String(n)}>
            {n}
          </option>
        ))}
      </select>
    );
  }

  if (entry.kind === 'empty') {
    const v = entry.validation;
    if (v?.type === 'list' && Array.isArray(v.list_values) && v.list_values.length > 0) {
      return (
        <select
          className="tipset-select tipset-select--in-cell"
          aria-label={t('aria.dropdownCell', { address: entry.address })}
          defaultValue=""
        >
          <option value=""> </option>
          {v.list_values.map((opt) => {
            const s = String(opt);
            return (
              <option key={s} value={s}>
                {translateWorkbookString(s, i18n)}
              </option>
            );
          })}
        </select>
      );
    }
    return null;
  }

  if (entry.value != null && entry.value !== '') {
    const raw = String(entry.value);
    const text = entry.kind === 'number' ? raw : translateWorkbookString(raw, i18n);
    return (
      <span className={entry.kind === 'number' ? 'tipset-cell-num' : 'tipset-cell-text'}>{text}</span>
    );
  }

  return null;
}

const COL_B = lettersToColIndex('B');
const COL_C = lettersToColIndex('C');
const COL_D = lettersToColIndex('D');
const COL_E = lettersToColIndex('E');
const COL_F = lettersToColIndex('F');
const COL_K = lettersToColIndex('K');
const COL_J = lettersToColIndex('J');
const COL_L = lettersToColIndex('L');
const COL_S = lettersToColIndex('S');
const COL_T = lettersToColIndex('T');

type SheetRowVariant = 'default' | 'bestThird' | 'r32';

type SheetRowProps = {
  row: GridCell[];
  rowIndex0: number;
  groupCtx: GroupRowCtx | null;
  rowVariant?: SheetRowVariant;
  /** Translated matchup blurb from mall column P (e.g. "2A vs 2B"); hover bubble on R32 rows. */
  r32MatchHint?: string | null;
  cellByAddress: Map<string, WorkbookCell>;
  scoreDrafts: Record<string, string>;
  metaDrafts: Record<string, string>;
  onMetaDraft: (normalizedAddress: string, value: string) => void;
  onScoreDraft: (address: string, value: string) => void;
  standingTables: Map<string, StandingRow[] | null>;
  bestThirdPlaceRows: BestThirdPlaceRow[] | null;
};

function SheetRow({
  row,
  rowIndex0,
  groupCtx,
  rowVariant = 'default',
  r32MatchHint = null,
  cellByAddress,
  scoreDrafts,
  metaDrafts,
  onMetaDraft,
  onScoreDraft,
  standingTables,
  bestThirdPlaceRows,
}: SheetRowProps) {
  const excelRow = rowIndex0 + 1;
  const isMetaBlockRow = excelRow < 10;
  const firstVisibleColIndex = row.findIndex((c) => c.kind !== 'skip');
  const isR32HintRow =
    rowVariant === 'r32' &&
    Boolean(r32MatchHint) &&
    excelRow >= R32_FIRST_MATCH_ROW &&
    excelRow <= R32_LAST_MATCH_ROW;

  const trClass = [isMetaBlockRow ? 'tipset-row--meta' : '', isR32HintRow ? 'tipset-r32-match-row' : '']
    .filter(Boolean)
    .join(' ');

  return (
    <tr
      aria-rowindex={excelRow}
      className={trClass || undefined}
      aria-describedby={isR32HintRow ? `tipset-r32-hint-${excelRow}` : undefined}
    >
      {row.map((cell: GridCell, columnIndex) => {
        if (cell.kind === 'skip') return null;
        const entry = cell.entry;
        const gridCol = columnIndex + 1;
        const isFirstVisible = firstVisibleColIndex >= 0 && columnIndex === firstVisibleColIndex;
        const addrNorm = normalizeCellAddress(cell.address);
        const isMetaTextField = isMetaFreeTextCell(addrNorm);
        const hasListValidation = entry?.validation?.type === 'list';
        const isFormula = entry?.kind === 'formula';

        const groupMods =
          groupCtx == null
            ? []
            : classifyGroupGridCell(gridCol, excelRow, groupCtx.blockHeaderExcelRow, groupCtx.rowInBlock);

        const isRoundOf32TeamCol = Boolean(
          groupCtx == null &&
            excelRow >= R32_FIRST_MATCH_ROW &&
            excelRow <= R32_LAST_MATCH_ROW &&
            (gridCol === COL_B || gridCol === COL_F),
        );
        const isPredTeamCol = Boolean(
          groupCtx && groupCtx.rowInBlock >= 1 && (gridCol === COL_B || gridCol === COL_F),
        );
        const isStatNameCol = Boolean(groupCtx && gridCol === COL_K && groupCtx.rowInBlock >= 1);
        const awayScorePad = Boolean(groupCtx && gridCol === COL_E);
        const afterAwayScore = Boolean(groupCtx && gridCol === COL_F);

        const b3Idx = bestThirdPlaceDataRowIndex(excelRow);
        const b3Header = isBestThirdPlaceHeaderCell(excelRow, gridCol);
        const b3HasRow = Boolean(b3Idx !== null && bestThirdPlaceRows && bestThirdPlaceRows[b3Idx]);

        const is1x2Col = Boolean(groupCtx && gridCol >= COL_G && gridCol <= COL_I);
        const isGroupDashCol = Boolean(groupCtx && gridCol === COL_D);
        const isGroupStatHeaderNumCol = Boolean(
          groupCtx && groupCtx.rowInBlock === 0 && gridCol >= COL_L && gridCol <= COL_T,
        );
        const isGroupScoreCol = Boolean(
          groupCtx &&
            groupCtx.rowInBlock >= 1 &&
            (gridCol === COL_C || gridCol === COL_E) &&
            entry &&
            isScoreDigitListValidation(entry.validation),
        );
        const alignNum =
          entry?.kind === 'number' || (isFormula && typeof entry?.cached_value === 'number');
        const alignClass = is1x2Col
          ? 'tipset-align-center'
          : b3Header && gridCol === COL_J
            ? 'tipset-align-center'
            : b3Header && gridCol >= COL_L && gridCol <= COL_T
              ? 'tipset-align-right'
              : b3Header && gridCol === COL_K
              ? 'tipset-align-left'
              : b3HasRow && gridCol === COL_J
                ? 'tipset-align-right'
                : b3HasRow && gridCol >= COL_L && gridCol <= COL_T
                  ? 'tipset-align-right'
                  : b3HasRow && gridCol === COL_K
                    ? 'tipset-align-left'
                    : isGroupDashCol
                      ? 'tipset-align-center'
                      : isGroupStatHeaderNumCol
                        ? 'tipset-align-right'
                        : isGroupScoreCol
                          ? 'tipset-align-right'
                          : alignNum
                            ? 'tipset-align-right'
                            : 'tipset-align-left';

        const b3StatCol = Boolean(
          (b3Header && gridCol >= COL_L && gridCol <= COL_T) ||
            (b3Idx !== null && gridCol >= COL_L && gridCol <= COL_S),
        );
        const b3RankCol = Boolean(b3Idx !== null && gridCol === COL_J);

        let b3Tier = '';
        if (b3Idx !== null && bestThirdPlaceRows && gridCol >= COL_J && gridCol <= COL_T) {
          const br = bestThirdPlaceRows[b3Idx];
          if (br) {
            b3Tier = br.rank <= 8 ? 'tipset-b3__cell--top8' : 'tipset-b3__cell--rest';
          }
        }

        const r32Dash =
          rowVariant === 'r32' &&
          excelRow >= R32_FIRST_MATCH_ROW &&
          excelRow <= R32_LAST_MATCH_ROW &&
          gridCol === COL_C;

        const showR32HintBubble = isR32HintRow && isFirstVisible && r32MatchHint;

        const className = [
          'tipset-cell',
          ...groupMods,
          alignClass,
          hasListValidation || isMetaTextField ? 'tipset-cell--input' : '',
          isPredTeamCol || isRoundOf32TeamCol ? 'tipset-cell--pred-team' : '',
          isStatNameCol ? 'tipset-cell--stat-name' : '',
          awayScorePad ? 'tipset-cell--away-score-pad' : '',
          afterAwayScore ? 'tipset-cell--after-away-score' : '',
          b3Header ? 'tipset-b3__header' : '',
          b3StatCol ? 'tipset-b3__stat-col' : '',
          b3RankCol ? 'tipset-b3__rank-col' : '',
          b3Tier,
          r32Dash ? 'tipset-cell--r32-mid' : '',
          showR32HintBubble ? 'tipset-r32-match-row__anchor' : '',
        ]
          .filter(Boolean)
          .join(' ');

        const body = (
          <CellBody
            address={cell.address}
            entry={entry}
            excelRow={excelRow}
            gridCol={gridCol}
            groupCtx={groupCtx}
            cellByAddress={cellByAddress}
            scoreDrafts={scoreDrafts}
            metaDrafts={metaDrafts}
            onMetaDraft={onMetaDraft}
            onScoreDraft={onScoreDraft}
            standingTables={standingTables}
            bestThirdPlaceRows={bestThirdPlaceRows}
          />
        );

        return (
          <td
            key={cell.address}
            className={className}
            rowSpan={cell.rowSpan}
            colSpan={cell.colSpan}
            data-address={cell.address}
            aria-colindex={gridCol}
          >
            {r32Dash ? (
              <div className="tipset-r32-scoreband">
                <div className="tipset-r32-scoreband__pad" aria-hidden />
                <div className="tipset-r32-scoreband__dash">{body}</div>
                <div className="tipset-r32-scoreband__pad" aria-hidden />
              </div>
            ) : (
              body
            )}
            {showR32HintBubble ? (
              <span id={`tipset-r32-hint-${excelRow}`} className="tipset-r32-hint-bubble" role="tooltip">
                {r32MatchHint}
              </span>
            ) : null}
          </td>
        );
      })}
    </tr>
  );
}

export function SpreadsheetView({ title, grid, rowCount, colCount, cellByAddress }: SpreadsheetViewProps) {
  const { t, i18n } = useTranslation('app');
  const [scoreDrafts, setScoreDrafts] = useState<Record<string, string>>({});
  const [metaDrafts, setMetaDrafts] = useState<Record<string, string>>({});
  const [bracketDrafts, setBracketDrafts] = useState<Record<string, string>>({});

  const onScoreDraft = useCallback((address: string, value: string) => {
    setScoreDrafts((d) => ({ ...d, [address]: value }));
  }, []);

  const onMetaDraft = useCallback((normalizedAddress: string, value: string) => {
    setMetaDrafts((d) => ({ ...d, [normalizedAddress]: value }));
  }, []);

  const onBracketDraft = useCallback((normalizedAddress: string, value: string) => {
    const key = normalizeBracketAddress(normalizedAddress);
    setBracketDrafts((d) => ({ ...d, [key]: value }));
  }, []);

  const standingTables = useMemo(
    () => computeGroupStandings(cellByAddress, scoreDrafts),
    [cellByAddress, scoreDrafts],
  );

  const bestThirdPlaceRows = useMemo(
    () => computeBestThirdPlaceTable(standingTables, cellByAddress),
    [standingTables, cellByAddress],
  );

  const preambleRows = grid.slice(0, GROUP_STAGE_FIRST_HEADER_GRID_INDEX);

  const bestThirdGridRows = useMemo(() => {
    const from = BEST_THIRD_PLACE_FIRST_GRID_ROW_INDEX;
    const to = Math.min(BEST_THIRD_PLACE_GRID_SLICE_END_EXCLUSIVE, grid.length);
    if (from >= to) return [];
    return grid.slice(from, to);
  }, [grid]);

  const r32GridRows = useMemo(() => {
    const from = R32_FIRST_GRID_ROW_INDEX;
    const to = Math.min(R32_GRID_SLICE_END_EXCLUSIVE, grid.length);
    if (from >= to) return [];
    return grid.slice(from, to);
  }, [grid]);

  const r32MatchHintForExcelRow = useCallback(
    (excelRow: number): string | null => {
      if (excelRow < R32_FIRST_MATCH_ROW || excelRow > R32_LAST_MATCH_ROW) return null;
      const entry = cellByAddress.get(`P${excelRow}`);
      const raw = entry?.value != null && entry.value !== '' ? String(entry.value).trim() : '';
      if (!raw) return null;
      return translateWorkbookString(raw, i18n);
    },
    [cellByAddress, i18n],
  );

  const sharedRowProps = {
    cellByAddress,
    scoreDrafts,
    metaDrafts,
    onMetaDraft,
    onScoreDraft,
    standingTables,
    bestThirdPlaceRows,
  };

  return (
    <div className="tipset-shell">
      <header className="tipset-header">
        <h1 className="tipset-title">{title}</h1>
        <p className="tipset-sub">{t('spreadsheet.subtitle')}</p>
      </header>

      <div className="tipset-scroll">
        <div className="tipset-table-wrap">
          <div className="tipset-stack">
            <table
              className="tipset-table tipset-table--plain"
              role="grid"
              aria-label={title}
              aria-rowcount={rowCount}
              aria-colcount={colCount}
            >
              <tbody>
                {preambleRows.map((row, i) => (
                  <SheetRow
                    key={i}
                    row={row}
                    rowIndex0={i}
                    groupCtx={null}
                    {...sharedRowProps}
                  />
                ))}
              </tbody>
            </table>

            <div className="tipset-group-stage">
              <table
                className="tipset-table tipset-table--group-stage"
                role="grid"
                aria-colcount={colCount}
              >
                {GROUP_COLGROUP}
                {GROUP_STAGE_BLOCKS.map((block, gIdx) => {
                  const from = GROUP_STAGE_FIRST_HEADER_GRID_INDEX + gIdx * GROUP_STAGE_ROW_STRIDE;
                  const blockRows = grid.slice(from, from + GROUP_STAGE_BLOCK_DATA_ROWS);
                  const blockHeaderExcelRow = block.startRow - 1;
                  return (
                    <Fragment key={block.id}>
                      <tbody className="tipset-group-block" aria-label={block.id}>
                        {blockRows.map((row, i) => (
                          <SheetRow
                            key={from + i}
                            row={row}
                            rowIndex0={from + i}
                            groupCtx={{ blockHeaderExcelRow, rowInBlock: i }}
                            {...sharedRowProps}
                          />
                        ))}
                      </tbody>
                      {gIdx < GROUP_STAGE_BLOCKS.length - 1 ? (
                        <tbody className="tipset-group-spacer-tbody">
                          <tr className="tipset-group-spacer" aria-hidden>
                            <td colSpan={colCount} />
                          </tr>
                        </tbody>
                      ) : null}
                    </Fragment>
                  );
                })}
              </table>
            </div>

            {(bestThirdGridRows.length > 0 || r32GridRows.length > 0) && (
              <div className="tipset-post-unified">
                <div className="tipset-post-unified__header">
                  {translateWorkbookString('16 delsfinal (32 lag)', i18n)}
                </div>
                <div
                  className="tipset-post-unified__body"
                  style={{
                    gridTemplateColumns: `${GROUP_STAGE_GRID_MATCH_FR}fr ${GROUP_STAGE_GRID_X_FR}fr ${GROUP_STAGE_GRID_STANDINGS_FR}fr`,
                  }}
                >
                  {r32GridRows.length > 0 && (
                    <div className="tipset-post-group__r32">
                      <div
                        className="tipset-post-sheet-bridge tipset-post-sheet-bridge--r32"
                        style={{ width: `calc(100% * 100 / ${GROUP_STAGE_GRID_MATCH_FR})` }}
                      >
                        <table
                          className="tipset-table tipset-table--group-stage tipset-table--r32"
                          role="grid"
                          aria-colcount={colCount}
                        >
                          <colgroup>
                            {Array.from({ length: colCount }, (_, i) => {
                              const i0 = lettersToColIndex('G') - 1;
                              const o0 = lettersToColIndex('O') - 1;
                              const p0 = lettersToColIndex('P') - 1;
                              const t0 = lettersToColIndex('T') - 1;
                              const collapse = (i >= i0 && i <= o0) || (i >= p0 && i <= t0);
                              const cls = GROUP_TABLE_COL_CLASSES[i];
                              return (
                                <col
                                  key={i}
                                  className={[cls, collapse ? 'tipset-r32-col--collapse' : '']
                                    .filter(Boolean)
                                    .join(' ')}
                                  style={
                                    collapse || !cls ? undefined : { width: `${GROUP_COL_WIDTH_PCT[i]}%` }
                                  }
                                />
                              );
                            })}
                          </colgroup>
                          <tbody className="tipset-r32-block">
                            {r32GridRows.map((row, i) => {
                              const rowIndex0 = R32_FIRST_GRID_ROW_INDEX + i;
                              if (rowIndex0 + 1 === R32_HEADER_ROW) return null;
                              return (
                                <SheetRow
                                  key={rowIndex0}
                                  row={row}
                                  rowIndex0={rowIndex0}
                                  groupCtx={null}
                                  rowVariant="r32"
                                  r32MatchHint={r32MatchHintForExcelRow(rowIndex0 + 1)}
                                  {...sharedRowProps}
                                />
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}
                  {bestThirdGridRows.length > 0 && (
                    <div className="tipset-post-group__b3">
                      <div
                        className="tipset-post-sheet-bridge tipset-post-sheet-bridge--b3"
                        style={{
                          width: `calc(100% * 100 / ${GROUP_STAGE_GRID_STANDINGS_FR})`,
                          marginLeft: `calc(-100% * ${GROUP_STAGE_GRID_MATCH_FR + GROUP_STAGE_GRID_X_FR} / ${GROUP_STAGE_GRID_STANDINGS_FR})`,
                        }}
                      >
                        <table
                          className="tipset-table tipset-table--group-stage tipset-table--best-third"
                          role="grid"
                          aria-colcount={colCount}
                        >
                          {GROUP_COLGROUP}
                          <tbody className="tipset-best-third-block">
                            {bestThirdGridRows.map((row, i) => (
                              <SheetRow
                                key={BEST_THIRD_PLACE_FIRST_GRID_ROW_INDEX + i}
                                row={row}
                                rowIndex0={BEST_THIRD_PLACE_FIRST_GRID_ROW_INDEX + i}
                                groupCtx={null}
                                rowVariant="bestThird"
                                {...sharedRowProps}
                              />
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}
                </div>
              </div>
            )}

            <KnockoutBracketView
              cellByAddress={cellByAddress}
              scoreDrafts={scoreDrafts}
              standingTables={standingTables}
              bestThirdPlaceRows={bestThirdPlaceRows}
              bracketDrafts={bracketDrafts}
              onBracketDraft={onBracketDraft}
            />
          </div>
        </div>
      </div>
    </div>
  );
}
