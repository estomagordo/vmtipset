#!/usr/bin/env python3
"""
Summarize tools/workbook_dump.json (from parse_workbook.py).

Usage:
  python tools/parse_workbook.py ... --json tools/workbook_dump.json
  python tools/analyze_workbook_dump.py
  python tools/analyze_workbook_dump.py --json other.json
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter
from pathlib import Path
from typing import Any


def _col_letter(n: int) -> str:
    """1 -> A, 27 -> AA (same idea as Excel columns)."""
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _addr_to_rc(addr: str) -> tuple[int, int] | None:
    m = re.match(r"^([A-Z]+)(\d+)$", addr)
    if not m:
        return None
    letters, row_s = m.group(1), m.group(2)
    col = 0
    for ch in letters:
        col = col * 26 + (ord(ch) - ord("A") + 1)
    return int(row_s), col


def _report(data: dict[str, Any], *, max_rows: int) -> None:
    if data.get("sheet_count", 0) != 1:
        print("note: expected one sheet in dump for focused report; found:", data.get("sheet_count"))

    sheet = data["sheets"][0]
    cells: list[dict[str, Any]] = sheet["cells"]

    print("workbook:", data["workbook"])
    print("parse flags:", data.get("parse"))
    print("sheet:", sheet["name"], "| worksheet visibility:", sheet.get("sheet_state"))
    print("dims: max_row ≤", sheet["max_row"], "max_col ≤", sheet["max_column"])
    print("merged_ranges:", len(sheet["merged_ranges"]))
    print()

    by_kind = Counter(c["kind"] for c in cells)
    print("cells by kind:", dict(by_kind))

    formulas = [c for c in cells if c.get("kind") == "formula"]
    values = [c for c in cells if c.get("kind") != "formula"]

    func_counter: Counter[str] = Counter()
    for c in formulas:
        f = (c.get("formula") or "").lstrip("=")
        m = re.match(r"([A-Za-z_]+)\s*\(", f)
        if m:
            func_counter[m.group(1).upper()] += 1
        else:
            func_counter["(ref/arith/other)"] += 1

    print("\nformula shape — first token looks like:")
    for name, n in func_counter.most_common(12):
        print(f"  {name}: {n}")

    rc = [x for x in (_addr_to_rc(c["address"]) for c in cells) if x]
    rs = [r for r, _ in rc]
    cs = [c for _, c in rc]
    print("\noccupied bounds: rows", min(rs), "–", max(rs), "cols", _col_letter(min(cs)), "–", _col_letter(max(cs)))

    strings = [c for c in values if c.get("kind") == "string"]
    numbers = [c for c in values if c.get("kind") == "number"]
    print("\nvalue cells: strings", len(strings), "| numbers", len(numbers))

    col_val: Counter[int] = Counter()
    col_any: Counter[int] = Counter()
    for c in cells:
        pt = _addr_to_rc(c["address"])
        if not pt:
            continue
        _, col = pt
        col_any[col] += 1
        if c.get("kind") != "formula":
            col_val[col] += 1

    print("\nmost non-formula cells by column (labels / inputs / blanks to fill):")
    for col, n in col_val.most_common(12):
        print(f"  {_col_letter(col)}: {n}")

    print("\nmost cells overall (formulas + values):")
    for col, n in col_any.most_common(12):
        print(f"  {_col_letter(col)}: {n}")

    row_formula: Counter[int] = Counter()
    row_value: Counter[int] = Counter()
    for c in cells:
        pt = _addr_to_rc(c["address"])
        if not pt:
            continue
        r, _ = pt
        if c.get("kind") == "formula":
            row_formula[r] += 1
        else:
            row_value[r] += 1

    value_only_rows = [
        (r, row_value[r])
        for r in range(min(rs), max(rs) + 1)
        if row_value[r] >= 3 and row_formula[r] == 0
    ]
    print(f"\nrows with only value cells (≥3 strings/numbers, no formulas on that row): {len(value_only_rows)}")
    print("(sample)", value_only_rows[:15])

    print(f"\nheader-like strings (first {max_rows} rows):")
    for c in sorted(strings, key=lambda x: _addr_to_rc(x["address"]) or (10**9, 10**9)):
        pt = _addr_to_rc(c["address"])
        if not pt or pt[0] > max_rows:
            continue
        v = c.get("value")
        if isinstance(v, str) and v.strip():
            tail = v.replace("\n", " ")
            if len(tail) > 72:
                tail = tail[:69] + "…"
            print(f"  {c['address']}: {tail!r}")


def main() -> int:
    p = argparse.ArgumentParser(description="Summarize workbook_dump.json from parse_workbook.py")
    p.add_argument("--json", type=Path, default=Path("tools/workbook_dump.json"))
    p.add_argument("--max-header-rows", type=int, default=25)
    args = p.parse_args()

    if not args.json.is_file():
        print(f"error: missing {args.json} — run parse_workbook.py --json first", file=sys.stderr)
        return 1

    data = json.loads(args.json.read_text(encoding="utf-8"))
    _report(data, max_rows=args.max_header_rows)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
