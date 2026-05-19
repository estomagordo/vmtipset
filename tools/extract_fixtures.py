#!/usr/bin/env python3
"""
Build fixtures.json — tournament structure and team names separated from workbook_dump.json.

Can run standalone on an existing dump:
  python tools/extract_fixtures.py tools/workbook_dump.json -o frontend/public/fixtures.json

Or import extract_fixtures_from_sheet() from parse_workbook after parsing.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

# Keep in sync with frontend/src/config/groupStageBlocks.ts
GROUP_STAGE_FIRST_ROW = 14
GROUP_STAGE_ROW_STRIDE = 8
GROUP_STAGE_FIXTURE_ROWS = 6
GROUP_STAGE_LETTERS = tuple("ABCDEFGHIJKL")

R32_HEADER_ROW = 109
R32_FIRST_MATCH_ROW = 110
R32_LAST_MATCH_ROW = 125

# R32 slot resolution (2026 FIFA); was hardcoded in frontend/src/lib/roundOf32.ts
R32_SLOTS: dict[str, tuple[dict[str, Any], dict[str, Any]]] = {
    "M73": (
        {"type": "group_rank", "group": "A", "rank": 2},
        {"type": "group_rank", "group": "B", "rank": 2},
    ),
    "M74": (
        {"type": "group_rank", "group": "E", "rank": 1},
        {"type": "best_third", "annex_index": 3},
    ),
    "M75": (
        {"type": "group_rank", "group": "F", "rank": 1},
        {"type": "group_rank", "group": "C", "rank": 2},
    ),
    "M76": (
        {"type": "group_rank", "group": "C", "rank": 1},
        {"type": "group_rank", "group": "F", "rank": 2},
    ),
    "M77": (
        {"type": "group_rank", "group": "I", "rank": 1},
        {"type": "best_third", "annex_index": 5},
    ),
    "M78": (
        {"type": "group_rank", "group": "E", "rank": 2},
        {"type": "group_rank", "group": "I", "rank": 2},
    ),
    "M79": (
        {"type": "group_rank", "group": "A", "rank": 1},
        {"type": "best_third", "annex_index": 0},
    ),
    "M80": (
        {"type": "group_rank", "group": "L", "rank": 1},
        {"type": "best_third", "annex_index": 7},
    ),
    "M81": (
        {"type": "group_rank", "group": "D", "rank": 1},
        {"type": "best_third", "annex_index": 2},
    ),
    "M82": (
        {"type": "group_rank", "group": "G", "rank": 1},
        {"type": "best_third", "annex_index": 4},
    ),
    "M83": (
        {"type": "group_rank", "group": "K", "rank": 2},
        {"type": "group_rank", "group": "L", "rank": 2},
    ),
    "M84": (
        {"type": "group_rank", "group": "H", "rank": 1},
        {"type": "group_rank", "group": "J", "rank": 2},
    ),
    "M85": (
        {"type": "group_rank", "group": "B", "rank": 1},
        {"type": "best_third", "annex_index": 1},
    ),
    "M86": (
        {"type": "group_rank", "group": "J", "rank": 1},
        {"type": "group_rank", "group": "H", "rank": 2},
    ),
    "M87": (
        {"type": "group_rank", "group": "K", "rank": 1},
        {"type": "best_third", "annex_index": 6},
    ),
    "M88": (
        {"type": "group_rank", "group": "D", "rank": 2},
        {"type": "group_rank", "group": "G", "rank": 2},
    ),
}

# frontend/src/lib/knockoutBracket.ts
KNOCKOUT_BRACKET: dict[str, Any] = {
    "r16": {
        "B145": ["M73", "M75"],
        "B147": ["M74", "M77"],
        "B149": ["M83", "M84"],
        "B151": ["M81", "M82"],
        "B153": ["M76", "M78"],
        "B155": ["M79", "M80"],
        "B157": ["M86", "M88"],
        "B159": ["M85", "M87"],
        "R145": ["M76", "M78"],
        "R147": ["M79", "M80"],
        "R149": ["M86", "M88"],
        "R151": ["M85", "M87"],
        "R153": ["M73", "M75"],
        "R155": ["M74", "M77"],
        "R157": ["M83", "M84"],
        "R159": ["M81", "M82"],
    },
    "qf": {
        "C146": ["B145", "B147"],
        "C150": ["B149", "B151"],
        "C154": ["B153", "B155"],
        "C158": ["B157", "B159"],
        "P146": ["R145", "R147"],
        "P150": ["R149", "R151"],
        "P154": ["R153", "R155"],
        "P158": ["R157", "R159"],
    },
    "sf": {
        "F148": ["C146", "C150"],
        "F156": ["C154", "C158"],
        "N148": ["P146", "P150"],
        "N156": ["P154", "P158"],
    },
    "finalist_left_semis": ["F148", "F156"],
    "finalist_right_semis": ["N148", "N156"],
    "finalist_picks": {"left": "G152", "right": "L152"},
    "champion": "J152",
}


def _cell_display(cells: dict[str, dict[str, Any]], address: str) -> str | None:
    c = cells.get(address.upper())
    if not c:
        return None
    raw = c.get("cached_value")
    if raw is None:
        raw = c.get("value")
    if raw is None or raw == "":
        return None
    return str(raw).strip()


def _extract_groups(cells: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    groups: list[dict[str, Any]] = []
    for i, letter in enumerate(GROUP_STAGE_LETTERS):
        start_row = GROUP_STAGE_FIRST_ROW + i * GROUP_STAGE_ROW_STRIDE
        header_row = start_row - 1
        group_id = f"Grupp {letter}"
        teams_ordered: list[str] = []
        seen: set[str] = set()
        fixtures: list[dict[str, Any]] = []

        for j in range(GROUP_STAGE_FIXTURE_ROWS):
            row = start_row + j
            home = _cell_display(cells, f"B{row}")
            away = _cell_display(cells, f"F{row}")
            match_id = _cell_display(cells, f"A{row}") or f"M{j + 1 + i * GROUP_STAGE_FIXTURE_ROWS}"
            if home and home not in seen:
                seen.add(home)
                teams_ordered.append(home)
            if away and away not in seen:
                seen.add(away)
                teams_ordered.append(away)
            fixtures.append(
                {
                    "match_id": match_id,
                    "excel_row": row,
                    "home_team": home,
                    "away_team": away,
                    "score_cells": {"home": f"C{row}", "away": f"E{row}"},
                }
            )

        groups.append(
            {
                "id": group_id,
                "letter": letter,
                "header_row": header_row,
                "start_row": start_row,
                "teams": teams_ordered,
                "fixtures": fixtures,
            }
        )
    return groups


def _extract_round_of_32(cells: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    matches: list[dict[str, Any]] = []
    for row in range(R32_FIRST_MATCH_ROW, R32_LAST_MATCH_ROW + 1):
        match_id = _cell_display(cells, f"A{row}") or f"M{73 + row - R32_FIRST_MATCH_ROW}"
        slot_hint = _cell_display(cells, f"P{row}")
        home_slot, away_slot = R32_SLOTS.get(match_id, ({}, {}))
        matches.append(
            {
                "match_id": match_id,
                "excel_row": row,
                "slot_hint": slot_hint,
                "home": home_slot,
                "away": away_slot,
            }
        )
    return matches


def extract_fixtures_from_dump(dump: dict[str, Any]) -> dict[str, Any]:
    if not dump.get("sheets"):
        raise ValueError("dump has no sheets")
    sheet = dump["sheets"][0]
    cells_list = sheet.get("cells") or []
    cells = {c["address"].upper(): c for c in cells_list}

    return {
        "version": 1,
        "source_workbook": dump.get("workbook"),
        "layout": {
            "group_stage": {
                "first_fixture_row": GROUP_STAGE_FIRST_ROW,
                "row_stride": GROUP_STAGE_ROW_STRIDE,
                "fixture_rows_per_group": GROUP_STAGE_FIXTURE_ROWS,
                "standing_rows_per_group": 4,
                "letters": list(GROUP_STAGE_LETTERS),
            },
            "round_of_32": {
                "header_row": R32_HEADER_ROW,
                "first_match_row": R32_FIRST_MATCH_ROW,
                "last_match_row": R32_LAST_MATCH_ROW,
            },
        },
        "groups": _extract_groups(cells),
        "round_of_32": _extract_round_of_32(cells),
        "knockout_bracket": KNOCKOUT_BRACKET,
    }


def extract_fixtures_from_sheet(sheet: dict[str, Any], *, workbook_name: str | None = None) -> dict[str, Any]:
    dump = {"workbook": workbook_name, "sheets": [sheet]}
    return extract_fixtures_from_dump(dump)


def main() -> int:
    parser = argparse.ArgumentParser(description="Extract fixtures.json from workbook_dump.json")
    parser.add_argument(
        "dump",
        type=Path,
        nargs="?",
        default=Path("frontend/public/workbook_dump.json"),
        help="Path to workbook_dump.json",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("frontend/public/fixtures.json"),
        help="Output fixtures.json path",
    )
    args = parser.parse_args()
    if not args.dump.is_file():
        print(f"error: not a file: {args.dump}", file=sys.stderr)
        return 1
    dump = json.loads(args.dump.read_text(encoding="utf-8"))
    fixtures = extract_fixtures_from_dump(dump)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(fixtures, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"wrote {args.output} ({len(fixtures['groups'])} groups, {len(fixtures['round_of_32'])} R32 matches)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
