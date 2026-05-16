#!/usr/bin/env python3
"""
Extract structure from a .xlsx for rebuilding the UX in a web app.

By default:
- Only the mall sheet ("VM-tipset 2026 mall")
- Only worksheets that are visible (not hidden / veryHidden)
- Only cells in visible rows and columns (skip hidden row/col)
- Only the functional region for the web app: columns A–T, rows ≤169, plus B172:B174
  and C174:F174 (row 174 dropdown inputs)

Output: sheet meta, merged cells in region, cell values and formulas, Excel data validation
(list dropdowns and resolved option values, including sources outside the clipped columns).

For formula cells, a second pass with data_only=True attaches cached_value when the workbook
contains a last-calculated result (use --no-merge-cached-values to skip).

Empty cells that have validation (e.g. B174–F174) are included so inputs are not dropped.

Usage:
  python -m venv .venv && source .venv/bin/activate
  pip install -r requirements.txt
  python tools/parse_workbook.py "VM-tipset-2026-Mall version 1.3.xlsx"
  python tools/parse_workbook.py workbook.xlsx --json output.json
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

# Default sheet for VM-tipset mall templates (others are usually workbook-hidden).
DEFAULT_SHEET_NAME = "VM-tipset 2026 mall"

# Web-app functional region (columns A–T, rows 1–169) plus listed exceptions.
DEFAULT_REGION_MAX_COLUMN_LETTER = "T"
DEFAULT_REGION_MAX_ROW = 169
DEFAULT_REGION_EXTRA_RANGES: tuple[str, ...] = ("B172:B174", "C174:F174")


def _col_letters_to_index(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def _range_to_box(spec: str) -> tuple[int, int, int, int]:
    from openpyxl.utils import range_boundaries

    min_col, min_row, max_col, max_row = range_boundaries(spec)
    return (min_col, min_row, max_col, max_row)


def _boxes_intersect(a: tuple[int, int, int, int], b: tuple[int, int, int, int]) -> bool:
    ac0, ar0, ac1, ar1 = a
    bc0, br0, bc1, br1 = b
    return not (ac1 < bc0 or bc1 < ac0 or ar1 < br0 or br1 < ar0)


def _cell_in_region(
    row: int,
    col: int,
    *,
    main_box: tuple[int, int, int, int],
    extra_boxes: list[tuple[int, int, int, int]],
) -> bool:
    min_c, min_r, max_c, max_r = main_box
    if min_r <= row <= max_r and min_c <= col <= max_c:
        return True
    for box in extra_boxes:
        ec0, er0, ec1, er1 = box
        if er0 <= row <= er1 and ec0 <= col <= ec1:
            return True
    return False


def _merge_in_region(
    merge_spec: str,
    *,
    main_box: tuple[int, int, int, int],
    extra_boxes: list[tuple[int, int, int, int]],
) -> bool:
    box = _range_to_box(merge_spec)
    if _boxes_intersect(box, main_box):
        return True
    for ex in extra_boxes:
        if _boxes_intersect(box, ex):
            return True
    return False


def _row_hidden(ws, row: int) -> bool:
    dim = ws.row_dimensions.get(row)
    return bool(dim and dim.hidden)


def _col_hidden(ws, col: int) -> bool:
    from openpyxl.utils import get_column_letter

    letter = get_column_letter(col)
    dim = ws.column_dimensions.get(letter)
    return bool(dim and dim.hidden)


def _cell_visible(ws, cell) -> bool:
    return not _row_hidden(ws, cell.row) and not _col_hidden(ws, cell.column)


def _merged_range_visible(ws, coord: str) -> bool:
    """Keep merge if its top-left cell lies in a visible row and column."""
    from openpyxl.utils import range_boundaries

    min_col, min_row, _, _ = range_boundaries(coord)
    return not _row_hidden(ws, min_row) and not _col_hidden(ws, min_col)


def _a1(row: int, column: int) -> str:
    from openpyxl.utils import get_column_letter

    return f"{get_column_letter(column)}{row}"


def _iter_rc_in_sqref(sqref: str | None):
    if not sqref:
        return
    for part in str(sqref).split():
        part = part.strip()
        if not part:
            continue
        min_c, min_r, max_c, max_r = _range_to_box(part)
        for r in range(min_r, max_r + 1):
            for c in range(min_c, max_c + 1):
                yield r, c


def _visible_cell_coordinate(ws, row: int, col: int, *, include_hidden_rows_cols: bool) -> bool:
    if include_hidden_rows_cols:
        return True
    return not _row_hidden(ws, row) and not _col_hidden(ws, col)


def _cell_eligible_for_parse(
    row: int,
    col: int,
    *,
    region_clip: bool,
    main_box: tuple[int, int, int, int],
    extra_boxes: list[tuple[int, int, int, int]],
    ws,
    include_hidden_rows_cols: bool,
) -> bool:
    if not _visible_cell_coordinate(ws, row, col, include_hidden_rows_cols=include_hidden_rows_cols):
        return False
    if region_clip:
        return _cell_in_region(row, col, main_box=main_box, extra_boxes=extra_boxes)
    return True


def _resolve_list_formula(ws, formula1: str, *, cache: dict[str, list[Any]]) -> list[Any]:
    if formula1 in cache:
        return cache[formula1]
    raw = formula1.strip()
    out: list[Any] = []

    if raw.startswith('"') and raw.endswith('"') and len(raw) >= 2:
        inner = raw[1:-1].replace('""', '"')
        for piece in inner.split(","):
            p = piece.strip()
            if p != "":
                out.append(p)
        cache[formula1] = out
        return out

    try:
        min_c, min_r, max_c, max_r = _range_to_box(raw)
    except Exception:
        cache[formula1] = []
        return []

    for r in range(min_r, max_r + 1):
        for c in range(min_c, max_c + 1):
            v = ws.cell(row=r, column=c).value
            if v is not None and v != "":
                out.append(v)
    cache[formula1] = out
    return out


def _serialize_data_validation(ws, dv, *, list_cache: dict[str, list[Any]]) -> dict[str, Any]:
    out: dict[str, Any] = {
        "type": dv.type,
        "operator": dv.operator,
        "formula1": dv.formula1,
        "formula2": dv.formula2,
        "allow_blank": dv.allow_blank,
        "show_dropdown": getattr(dv, "showDropDown", None),
        "show_error_message": getattr(dv, "showErrorMessage", None),
        "error_title": getattr(dv, "errorTitle", None),
        "error": getattr(dv, "error", None),
        "prompt_title": getattr(dv, "promptTitle", None),
        "prompt": getattr(dv, "prompt", None),
        "sqref": str(dv.sqref) if dv.sqref else None,
    }
    if dv.type == "list" and dv.formula1:
        out["list_values"] = _resolve_list_formula(ws, dv.formula1, cache=list_cache)
    return out


def _cell_payload(cell) -> dict[str, Any]:
    """openpyxl cell -> serializable dict."""
    coord = cell.coordinate
    dt = cell.data_type
    number_format = cell.number_format or ""

    # Formula cells: data_type == 'f'. value is the formula string (e.g. "=A1+B1").
    if dt == "f":
        return {
            "address": coord,
            "kind": "formula",
            "formula": cell.value,
            "number_format": number_format,
        }

    v = cell.value
    if v is None:
        return {}

    kind = "empty"
    if dt == "n":
        kind = "number"
    elif dt == "s":
        kind = "string"
    elif dt == "b":
        kind = "bool"
    elif dt == "d":
        kind = "date"
    elif dt == "e":
        kind = "error"
    else:
        kind = dt or "other"

    out: dict[str, Any] = {
        "address": coord,
        "kind": kind,
        "value": v,
    }
    if number_format and number_format != "General":
        out["number_format"] = number_format
    return out


def parse_workbook(
    path: Path,
    *,
    data_only: bool,
    sheet_name: str | None,
    all_visible_sheets: bool,
    include_hidden_sheets: bool,
    include_hidden_rows_cols: bool,
    strict_sheet: bool,
    region_clip: bool,
    region_max_column_letter: str,
    region_max_row: int,
    region_extra_ranges: tuple[str, ...],
    merge_cached_values: bool,
) -> dict[str, Any]:
    from openpyxl import load_workbook

    wb_values = None
    if merge_cached_values and not data_only:
        wb_values = load_workbook(path, data_only=True, read_only=True)

    wb = load_workbook(path, data_only=data_only, read_only=False)
    try:
        all_names = list(wb.sheetnames)

        targets: list[str] = []
        for name in all_names:
            ws = wb[name]
            state = getattr(ws, "sheet_state", None) or "visible"

            if not include_hidden_sheets and state != "visible":
                continue

            if all_visible_sheets:
                targets.append(name)
            elif sheet_name is not None:
                if name == sheet_name:
                    targets.append(name)
            else:
                targets.append(name)

        if strict_sheet and sheet_name is not None and sheet_name not in all_names:
            known = ", ".join(repr(n) for n in all_names)
            raise ValueError(f"no sheet named {sheet_name!r}; available: {known}")

        if strict_sheet and sheet_name is not None and sheet_name not in targets:
            st_rec: list[str] = []
            for name in all_names:
                wsi = wb[name]
                st = getattr(wsi, "sheet_state", None) or "visible"
                st_rec.append(f"{name!r} ({st})")
            raise ValueError(
                f"sheet {sheet_name!r} is not among selected sheets "
                "(hidden worksheets are skipped by default; pass --include-hidden-sheets). "
                f"Sheets: {', '.join(st_rec)}"
            )

        max_col_idx = _col_letters_to_index(region_max_column_letter)
        main_box = (1, 1, max_col_idx, region_max_row)
        extra_boxes = [_range_to_box(s) for s in region_extra_ranges]

        sheets: list[dict[str, Any]] = []
        for name in targets:
            ws = wb[name]
            merged: list[str] = []
            for rng in ws.merged_cells.ranges:
                spec = str(rng)
                if not include_hidden_rows_cols and not _merged_range_visible(ws, spec):
                    continue
                if region_clip and not _merge_in_region(spec, main_box=main_box, extra_boxes=extra_boxes):
                    continue
                merged.append(spec)

            cells: list[dict[str, Any]] = []
            for row in ws.iter_rows():
                for cell in row:
                    if not include_hidden_rows_cols and not _cell_visible(ws, cell):
                        continue
                    if region_clip and not _cell_in_region(
                        cell.row, cell.column, main_box=main_box, extra_boxes=extra_boxes
                    ):
                        continue
                    payload = _cell_payload(cell)
                    if payload:
                        cells.append(payload)

            list_cache: dict[str, list[Any]] = {}
            by_addr: dict[str, dict[str, Any]] = {c["address"]: c for c in cells}

            from openpyxl.utils import coordinate_to_tuple

            data_validations: list[dict[str, Any]] = []
            seen_dv: set[tuple[Any, ...]] = set()
            for dv in ws.data_validations.dataValidation:
                touches = False
                for r, col in _iter_rc_in_sqref(dv.sqref):
                    if _cell_eligible_for_parse(
                        r,
                        col,
                        region_clip=region_clip,
                        main_box=main_box,
                        extra_boxes=extra_boxes,
                        ws=ws,
                        include_hidden_rows_cols=include_hidden_rows_cols,
                    ):
                        touches = True
                        break
                if not touches:
                    continue

                vdict = _serialize_data_validation(ws, dv, list_cache=list_cache)
                dedupe_key = (vdict.get("sqref"), vdict.get("type"), vdict.get("formula1"))
                if dedupe_key not in seen_dv:
                    seen_dv.add(dedupe_key)
                    data_validations.append(vdict)

                for r, col in _iter_rc_in_sqref(dv.sqref):
                    if not _cell_eligible_for_parse(
                        r,
                        col,
                        region_clip=region_clip,
                        main_box=main_box,
                        extra_boxes=extra_boxes,
                        ws=ws,
                        include_hidden_rows_cols=include_hidden_rows_cols,
                    ):
                        continue
                    addr = _a1(r, col)
                    if addr not in by_addr:
                        by_addr[addr] = {
                            "address": addr,
                            "kind": "empty",
                            "value": None,
                            "validation": vdict,
                        }
                    else:
                        by_addr[addr]["validation"] = vdict

            cells_out = sorted(by_addr.values(), key=lambda c: coordinate_to_tuple(c["address"]))

            if wb_values is not None:
                wsv = wb_values[name]
                for c in cells_out:
                    if c.get("kind") != "formula":
                        continue
                    addr = c.get("address")
                    if not addr:
                        continue
                    try:
                        cached = wsv[addr].value
                    except Exception:
                        continue
                    if cached is not None:
                        c["cached_value"] = cached

            sheets.append(
                {
                    "name": name,
                    "sheet_state": getattr(ws, "sheet_state", None) or "visible",
                    "max_row": ws.max_row,
                    "max_column": ws.max_column,
                    "merged_ranges": merged,
                    "data_validations": data_validations,
                    "cells": cells_out,
                }
            )

        return {
            "workbook": path.name,
            "data_only": data_only,
            "parse": {
                "sheet_name_filter": sheet_name,
                "all_visible_sheets": all_visible_sheets,
                "include_hidden_sheets": include_hidden_sheets,
                "include_hidden_rows_cols": include_hidden_rows_cols,
                "region_clip": region_clip,
                "region_max_column_letter": region_max_column_letter,
                "region_max_row": region_max_row,
                "region_extra_ranges": list(region_extra_ranges),
                "merge_cached_values": merge_cached_values and not data_only,
            },
            "sheet_count": len(sheets),
            "sheets": sheets,
        }
    finally:
        wb.close()
        if wb_values is not None:
            wb_values.close()


def main() -> int:
    parser = argparse.ArgumentParser(description="Parse .xlsx to JSON (values + formulas).")
    parser.add_argument(
        "workbook",
        type=Path,
        nargs="?",
        default=Path("VM-tipset-2026-Mall version 1.3.xlsx"),
        help="Path to .xlsx (default: VM-tipset-2026-Mall version 1.3.xlsx)",
    )
    parser.add_argument(
        "--sheet",
        metavar="NAME",
        default=DEFAULT_SHEET_NAME,
        help=(
            f"Only parse this worksheet (default: {DEFAULT_SHEET_NAME!r}). "
            "Ignored when --all-visible-sheets is set."
        ),
    )
    parser.add_argument(
        "--all-visible-sheets",
        action="store_true",
        help="Parse every visible sheet (ignore --sheet).",
    )
    parser.add_argument(
        "--include-hidden-sheets",
        action="store_true",
        help="Allow hidden / veryHidden worksheets in addition to visible ones.",
    )
    parser.add_argument(
        "--include-hidden-rows-cols",
        action="store_true",
        help="Include cells in hidden rows or columns (and merged ranges whose top-left is hidden).",
    )
    parser.add_argument(
        "--no-region-clip",
        action="store_false",
        dest="region_clip",
        help=(
            "Keep the full parsed worksheet (within visibility/hidden-row settings). "
            "Default: clip to columns A–T and rows ≤169, plus B172:B174."
        ),
    )
    parser.add_argument(
        "--region-max-column",
        metavar="LETTER",
        default=DEFAULT_REGION_MAX_COLUMN_LETTER,
        help=f"Last column letter included when region clipping is on (default: {DEFAULT_REGION_MAX_COLUMN_LETTER!r}).",
    )
    parser.add_argument(
        "--region-max-row",
        type=int,
        metavar="N",
        default=DEFAULT_REGION_MAX_ROW,
        help=f"Highest row included in the main block when clipping (default: {DEFAULT_REGION_MAX_ROW}).",
    )
    parser.add_argument(
        "--region-extra-range",
        action="append",
        metavar="RANGE",
        dest="region_extra_ranges",
        help=(
            "Additional A1-style range to include when clipping (repeatable). "
            f"Default when omitting: {', '.join(DEFAULT_REGION_EXTRA_RANGES)}"
        ),
    )
    parser.set_defaults(region_clip=True, merge_cached_values=True)
    parser.add_argument(
        "--json",
        type=Path,
        metavar="FILE",
        help="Write full dump to FILE (UTF-8)",
    )
    parser.add_argument(
        "--data-only",
        action="store_true",
        help="Load cached values only (formulas become computed values or None if no cache)",
    )
    parser.add_argument(
        "--no-merge-cached-values",
        action="store_false",
        dest="merge_cached_values",
        help="Do not add cached_value for formula cells (second data_only read). Default: merge.",
    )
    parser.add_argument(
        "--summary",
        action="store_true",
        help="Print per-sheet counts only (no cell listing on stdout)",
    )
    args = parser.parse_args()

    wb_path: Path = args.workbook
    if not wb_path.is_file():
        print(f"error: not a file: {wb_path}", file=sys.stderr)
        return 1

    sheet_name: str | None
    if args.all_visible_sheets:
        sheet_name = None
    else:
        sheet_name = args.sheet

    extra = (
        tuple(args.region_extra_ranges)
        if args.region_extra_ranges
        else DEFAULT_REGION_EXTRA_RANGES
    )

    try:
        data = parse_workbook(
            wb_path,
            data_only=args.data_only,
            sheet_name=sheet_name,
            all_visible_sheets=bool(args.all_visible_sheets),
            include_hidden_sheets=bool(args.include_hidden_sheets),
            include_hidden_rows_cols=bool(args.include_hidden_rows_cols),
            strict_sheet=sheet_name is not None,
            region_clip=bool(args.region_clip),
            region_max_column_letter=str(args.region_max_column).upper(),
            region_max_row=int(args.region_max_row),
            region_extra_ranges=extra,
            merge_cached_values=bool(args.merge_cached_values),
        )
    except ValueError as e:
        print(f"error: {e}", file=sys.stderr)
        return 1

    if args.json:
        args.json.write_text(
            json.dumps(data, ensure_ascii=False, indent=2, default=str) + "\n",
            encoding="utf-8",
        )
        print(f"wrote {args.json}")

    if args.summary:
        for sh in data["sheets"]:
            formulas = sum(1 for c in sh["cells"] if c.get("kind") == "formula")
            values = len(sh["cells"]) - formulas
            clip = data.get("parse", {})
            clip_note = ""
            if clip.get("region_clip"):
                clip_note = (
                    f" | clip A:{clip.get('region_max_column_letter')} "
                    f"row≤{clip.get('region_max_row')} + {clip.get('region_extra_ranges')}"
                )
            print(
                f'{sh["name"]}: rows≤{sh["max_row"]} cols≤{sh["max_column"]} '
                f"cells={len(sh['cells'])} (values≈{values}, formulas≈{formulas}) "
                f"merged={len(sh['merged_ranges'])} dv={len(sh.get('data_validations', []))}"
                f"{clip_note}"
            )
    elif not args.json:
        # Default: pretty JSON to stdout (can be large).
        json.dump(data, sys.stdout, ensure_ascii=False, indent=2, default=str)
        sys.stdout.write("\n")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
