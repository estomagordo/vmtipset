#!/usr/bin/env python3
"""
Extract user-visible strings from workbook_dump.json and emit i18n JSON + key map.

Run from repo root after updating frontend/public/workbook_dump.json:
  python tools/generate_workbook_labels.py

Outputs:
  frontend/src/i18n/locales/sv/workbookLabels.json
  frontend/src/i18n/locales/en/workbookLabels.json
  frontend/src/i18n/generated/workbookStringToKey.json
"""

from __future__ import annotations

import hashlib
import json
import re
import unicodedata
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DUMP = ROOT / "frontend/public/workbook_dump.json"
OUT_SV = ROOT / "frontend/src/i18n/locales/sv/workbookLabels.json"
OUT_EN = ROOT / "frontend/src/i18n/locales/en/workbookLabels.json"
OUT_MAP = ROOT / "frontend/src/i18n/generated/workbookStringToKey.json"


def is_formula_artifact(s: str) -> bool:
    return s.strip().startswith("=")


def fold_ascii(s: str) -> str:
    n = unicodedata.normalize("NFKD", s)
    return "".join(c for c in n if not unicodedata.combining(c))


def slug_base(s: str) -> str:
    x = fold_ascii(s.strip().lower())
    x = re.sub(r"[^a-z0-9]+", "_", x)
    x = re.sub(r"_+", "_", x).strip("_")
    return x


def collect_unique_strings(data: dict) -> list[str]:
    unique: set[str] = set()
    cells = data["sheets"][0]["cells"]
    for c in cells:
        k = c.get("kind")
        if k == "string" and c.get("value") is not None:
            v = c["value"]
            if isinstance(v, str) and v.strip():
                unique.add(v)
        if k == "formula" and "cached_value" in c:
            cv = c["cached_value"]
            if isinstance(cv, str) and cv.strip():
                unique.add(cv)
        if k == "empty" and c.get("validation"):
            for lv in c["validation"].get("list_values") or []:
                if isinstance(lv, str) and lv.strip():
                    unique.add(lv)
    for dv in data["sheets"][0].get("data_validations") or []:
        for lv in dv.get("list_values") or []:
            if isinstance(lv, str) and lv.strip():
                unique.add(lv)
    return sorted(unique)


# Swedish (source) -> English. Bracket / match codes that are language-neutral map to themselves.
EN: dict[str, str] = {
    " - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ": (
        " - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
    ),
    "10 poäng": "10 points",
    "15 p": "15 pts",
    "16 delsfinal (32 lag)": "Round of 16 (32 teams)",
    "1A vs 3C/3E/3F/3H/3I": "1A vs 3C/3E/3F/3H/3I",
    "1B vs 3E/3F/3G/3I/3J": "1B vs 3E/3F/3G/3I/3J",
    "1C vs 2F": "1C vs 2F",
    "1D vs 3B/3E/3F/3I/3J": "1D vs 3B/3E/3F/3I/3J",
    "1E vs 3A/3B/3C/3D/3F": "1E vs 3A/3B/3C/3D/3F",
    "1F vs 2C": "1F vs 2C",
    "1G vs 3A/3E/3H/3I/3J": "1G vs 3A/3E/3H/3I/3J",
    "1H vs 2J": "1H vs 2J",
    "1I vs 3C/3D/3F/3G/3H": "1I vs 3C/3D/3F/3G/3H",
    "1J vs 2H": "1J vs 2H",
    "1K vs 3D/3E/3I/3J/3L": "1K vs 3D/3E/3I/3J/3L",
    "1L vs 3E/3H/3I/3J/3K": "1L vs 3E/3H/3I/3J/3K",
    "1X2 = 1p": "1X2 = 1 pt",
    "2A vs 2B": "2A vs 2B",
    "2D vs 2G": "2D vs 2G",
    "2E vs 2I": "2E vs 2I",
    "2K vs 2L": "2K vs 2L",
    "4 p/lag": "4 pts/team",
    "5 poäng": "5 points",
    "6 p/lag": "6 pts/team",
    "8 bästa 3:or": "8 best third-placed",
    "8 p/lag": "8 pts/team",
    "<< Instruktioner >>": "<< Instructions >>",
    "Algeriet": "Algeria",
    "Argentina": "Argentina",
    "Australien": "Australia",
    "Belgien": "Belgium",
    "Bosnien": "Bosnia and Herzegovina",
    "Brasilien": "Brazil",
    "Bronsfinal": "Bronze medal match",
    "Colombia": "Colombia",
    "Curaçao": "Curaçao",
    "DR Kongo": "DR Congo",
    "Din email:": "Your email:",
    "Din poäng": "Your score",
    "Ditt namn:": "Your name:",
    "Ditt spel": "Your picks",
    "Ecuador": "Ecuador",
    "Egypten": "Egypt",
    "Elfenbenskusten": "Ivory Coast",
    "England": "England",
    "F": "L",
    "Final": "Final",
    "Frankrike": "France",
    "GM": "GF",
    "Ghana": "Ghana",
    "Grupp A": "Group A",
    "Grupp B": "Group B",
    "Grupp C": "Group C",
    "Grupp D": "Group D",
    "Grupp E": "Group E",
    "Grupp F": "Group F",
    "Grupp G": "Group G",
    "Grupp H": "Group H",
    "Grupp I": "Group I",
    "Grupp J": "Group J",
    "Grupp K": "Group K",
    "Grupp L": "Group L",
    "Haiti": "Haiti",
    "IM": "GA",
    "Irak": "Iraq",
    "Iran": "Iran",
    "Ja": "Yes",
    "Japan": "Japan",
    "Jordanien": "Jordan",
    "Kanada": "Canada",
    "Kap Verde": "Cape Verde",
    "Kroatien": "Croatia",
    "Kvartsfinal": "Quarter-final",
    "Lag som gör flest mål": "Team with the most goals",
    "M+/-": "GD",
    "Marocko": "Morocco",
    "Match nr": "Match no.",
    "Mexiko": "Mexico",
    "Nederländerna": "Netherlands",
    "Nej": "No",
    "Norge": "Norway",
    "Nya Zeeland": "New Zealand",
    "O": "D",
    "Panama": "Panama",
    "Paraguay": "Paraguay",
    "Portugal": "Portugal",
    "Poäng": "Pts",
    "Qatar": "Qatar",
    "Resultat = 2p": "Result = 2 pts",
    "S": "P",
    "Saudiarabien": "Saudi Arabia",
    "Schweiz": "Switzerland",
    "Semifinal": "Semi-final",
    "Senegal": "Senegal",
    "Skottland": "Scotland",
    "Skyttekung": "Top scorer",
    "Slutspel tillsammans med de 8 bästa grupp 3:orna": (
        "Knockout stage with the 8 best third-placed teams"
    ),
    "Sluttabell = 2p/lag": "Final table = 2 pts/team",
    "Spanien": "Spain",
    "Straffar i finalen": "Penalties in the final",
    "Sverige": "Sweden",
    "Sydafrika": "South Africa",
    "Sydkorea": "South Korea",
    "Tjeckien": "Czechia",
    "Tunisien": "Tunisia",
    "Turkiet": "Turkey",
    "Tyskland": "Germany",
    "USA": "USA",
    "Uruguay": "Uruguay",
    "Uzbekistan": "Uzbekistan",
    "V": "W",
    "Vinnare": "Winner",
    "X": "X",
    "Åttondelsfinal": "Round of 16",
    "Österrike": "Austria",
    "―": "—",
    "VM-tipset 2026 mall": "World Cup 2026 pool (template)",
}


def stable_key_body(s: str, used_slugs: dict[str, int]) -> str:
    base = slug_base(s)
    if not base:
        base = f"u_{hashlib.sha256(s.encode('utf-8')).hexdigest()[:14]}"
    n = used_slugs.get(base, 0)
    used_slugs[base] = n + 1
    if n == 0:
        return base
    return f"{base}_{n}"


def main() -> int:
    data = json.loads(DUMP.read_text(encoding="utf-8"))
    all_strings = collect_unique_strings(data)
    to_label = {s for s in all_strings if not is_formula_artifact(s)}
    sheet_title = data["sheets"][0].get("name")
    if isinstance(sheet_title, str) and sheet_title.strip():
        to_label.add(sheet_title.strip())
    to_label = sorted(to_label)

    for s in to_label:
        if s not in EN:
            if re.fullmatch(r"M\d+", s):
                EN[s] = s
            else:
                EN[s] = s

    used_slugs: dict[str, int] = {}
    string_to_key: dict[str, str] = {}
    labels_sv: dict[str, str] = {}
    labels_en: dict[str, str] = {}

    for s in to_label:
        key_body = stable_key_body(s, used_slugs)
        key = f"wb.{key_body}"
        string_to_key[s] = key
        labels_sv[key] = s
        labels_en[key] = EN[s]

    OUT_SV.parent.mkdir(parents=True, exist_ok=True)
    OUT_EN.parent.mkdir(parents=True, exist_ok=True)
    OUT_MAP.parent.mkdir(parents=True, exist_ok=True)

    OUT_SV.write_text(json.dumps(labels_sv, ensure_ascii=False, sort_keys=True, indent=2) + "\n", encoding="utf-8")
    OUT_EN.write_text(json.dumps(labels_en, ensure_ascii=False, sort_keys=True, indent=2) + "\n", encoding="utf-8")
    OUT_MAP.write_text(json.dumps(string_to_key, ensure_ascii=False, sort_keys=True, indent=2) + "\n", encoding="utf-8")
    print(f"wrote {len(to_label)} labels")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
