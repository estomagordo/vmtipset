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
  groupStageFooterStartGridIndex,
} from '../config/groupStageBlocks';
import {
  getScoreSelectValue,
  isScoreDigitListValidation,
  SCORE_OPTION_VALUES,
} from '../lib/scoreCells';
import { translateWorkbookString } from '../i18n/workbookStrings';
import { isMetaFreeTextCell } from '../lib/metaTextFields';
import { classifyGroupGridCell, GROUP_COL_WIDTH_PCT, GROUP_TABLE_COL_CLASSES } from '../lib/groupGridStyle';
import { matchPickFromRow } from '../lib/match1x2';
import { lettersToColIndex } from '../lib/excelAddress';

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
}: CellBodyProps) {
  const { t, i18n } = useTranslation('app');
  const key = normalizeCellAddress(address);

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
const COL_L = lettersToColIndex('L');
const COL_T = lettersToColIndex('T');

type SheetRowProps = {
  row: GridCell[];
  rowIndex0: number;
  groupCtx: GroupRowCtx | null;
  cellByAddress: Map<string, WorkbookCell>;
  scoreDrafts: Record<string, string>;
  metaDrafts: Record<string, string>;
  onMetaDraft: (normalizedAddress: string, value: string) => void;
  onScoreDraft: (address: string, value: string) => void;
  standingTables: Map<string, StandingRow[] | null>;
};

function SheetRow({
  row,
  rowIndex0,
  groupCtx,
  cellByAddress,
  scoreDrafts,
  metaDrafts,
  onMetaDraft,
  onScoreDraft,
  standingTables,
}: SheetRowProps) {
  const excelRow = rowIndex0 + 1;
  const isMetaBlockRow = excelRow < 10;

  return (
    <tr aria-rowindex={excelRow} className={isMetaBlockRow ? 'tipset-row--meta' : undefined}>
      {row.map((cell: GridCell, columnIndex) => {
        if (cell.kind === 'skip') return null;
        const entry = cell.entry;
        const gridCol = columnIndex + 1;
        const addrNorm = normalizeCellAddress(cell.address);
        const isMetaTextField = isMetaFreeTextCell(addrNorm);
        const hasListValidation = entry?.validation?.type === 'list';
        const isFormula = entry?.kind === 'formula';

        const groupMods =
          groupCtx == null
            ? []
            : classifyGroupGridCell(gridCol, excelRow, groupCtx.blockHeaderExcelRow, groupCtx.rowInBlock);

        const isPredTeamCol = Boolean(
          groupCtx && groupCtx.rowInBlock >= 1 && (gridCol === COL_B || gridCol === COL_F),
        );
        const isStatNameCol = Boolean(groupCtx && gridCol === COL_K && groupCtx.rowInBlock >= 1);
        const awayScorePad = Boolean(groupCtx && gridCol === COL_E);
        const afterAwayScore = Boolean(groupCtx && gridCol === COL_F);

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
          : isGroupDashCol
            ? 'tipset-align-center'
            : isGroupStatHeaderNumCol
              ? 'tipset-align-right'
              : isGroupScoreCol
                ? 'tipset-align-right'
                : alignNum
                  ? 'tipset-align-right'
                  : 'tipset-align-left';

        const className = [
          'tipset-cell',
          ...groupMods,
          alignClass,
          hasListValidation || isMetaTextField ? 'tipset-cell--input' : '',
          isPredTeamCol ? 'tipset-cell--pred-team' : '',
          isStatNameCol ? 'tipset-cell--stat-name' : '',
          awayScorePad ? 'tipset-cell--away-score-pad' : '',
          afterAwayScore ? 'tipset-cell--after-away-score' : '',
        ]
          .filter(Boolean)
          .join(' ');

        return (
          <td
            key={cell.address}
            className={className}
            rowSpan={cell.rowSpan}
            colSpan={cell.colSpan}
            data-address={cell.address}
            aria-colindex={gridCol}
          >
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
            />
          </td>
        );
      })}
    </tr>
  );
}

export function SpreadsheetView({ title, grid, rowCount, colCount, cellByAddress }: SpreadsheetViewProps) {
  const { t } = useTranslation('app');
  const [scoreDrafts, setScoreDrafts] = useState<Record<string, string>>({});
  const [metaDrafts, setMetaDrafts] = useState<Record<string, string>>({});

  const onScoreDraft = useCallback((address: string, value: string) => {
    setScoreDrafts((d) => ({ ...d, [address]: value }));
  }, []);

  const onMetaDraft = useCallback((normalizedAddress: string, value: string) => {
    setMetaDrafts((d) => ({ ...d, [normalizedAddress]: value }));
  }, []);

  const standingTables = useMemo(
    () => computeGroupStandings(cellByAddress, scoreDrafts),
    [cellByAddress, scoreDrafts],
  );

  const preambleRows = grid.slice(0, GROUP_STAGE_FIRST_HEADER_GRID_INDEX);
  const footerStart = groupStageFooterStartGridIndex();
  const footerRows = footerStart < grid.length ? grid.slice(footerStart) : [];

  const sharedRowProps = {
    cellByAddress,
    scoreDrafts,
    metaDrafts,
    onMetaDraft,
    onScoreDraft,
    standingTables,
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

            {footerRows.length > 0 && (
              <table className="tipset-table tipset-table--plain tipset-footer-table" role="grid">
                <tbody>
                  {footerRows.map((row, i) => (
                    <SheetRow
                      key={footerStart + i}
                      row={row}
                      rowIndex0={footerStart + i}
                      groupCtx={null}
                      {...sharedRowProps}
                    />
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
