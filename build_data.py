from __future__ import annotations

import json
from datetime import date, datetime, time, timedelta
from pathlib import Path

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parents[1]
BOOK = ROOT / "outputs" / "portalwars2-log-extract" / "Splitgate_match_stats_and_medals.xlsx"
OUT = Path(__file__).resolve().parent / "data.js"


def encode(value):
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    if isinstance(value, timedelta):
        return value.total_seconds()
    return value


def rows(sheet, start_row):
    headers = [cell.value for cell in sheet[start_row]]
    output = []
    for values in sheet.iter_rows(min_row=start_row + 1, values_only=True):
        if not any(value is not None for value in values):
            continue
        output.append({header: encode(value) for header, value in zip(headers, values)})
    return output


book = load_workbook(BOOK, data_only=True, read_only=True)
matches = rows(book["Match Summary"], 8)[:6]
named_winner_overrides = {"1": "Delta", "3": "Echo", "4": "Alpha", "5": "Alpha"}
for match in matches:
    override = named_winner_overrides.get(str(match.get("Match")))
    if override:
        match["Winner"] = override
data = {
    "matches": matches,
    "teams": rows(book["Teams"], 5),
    "players": rows(book["Player Medals"], 5),
    "events": rows(book["Medal Events"], 5),
    "generatedAt": datetime.now().astimezone().isoformat(),
    "source": "PortalWars2.log",
}

OUT.write_text("window.SPLITSTATS_DATA = " + json.dumps(data, ensure_ascii=False, separators=(",", ":")) + ";\n", encoding="utf-8")
print(f"Wrote {OUT} ({len(data['matches'])} matches, {len(data['teams'])} team rows, {len(data['players'])} player rows, {len(data['events'])} events)")
