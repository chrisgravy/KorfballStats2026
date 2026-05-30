import json
from openpyxl import load_workbook
from datetime import datetime
import re

wb = load_workbook("2026-KSA-State-League-Timetable.xlsx", data_only=True)
ws = wb["2026 timetable"]

rows = list(ws.iter_rows(values_only=True))

TEAM_MAP = {
    "Boomers": "Adelaide Boomers",

    "Scaldis": "Scaldis Tigers",

    "Scaldis Black": "Scaldis Tigers - Black",

    "North": "North Adelaide Roosters",

    "Arista": "Arista Marion",

    "Flinders": "Flinders Sharks",

    "Flinders Blue": "Flinders Sharks - Blue"
}

matches = []

for i, row in enumerate(rows):

    # -------------------------
    # Round Row
    # -------------------------
    if row[1] and isinstance(row[1], str) and row[1].startswith("Round"):

        current_round = row[1]

        venue_row = rows[i + 1]
        current_venue = venue_row[1]

        match = re.search(
            r"Round\s+(\d+)\s+-\s+(\d+/\d+/\d+)",
            row[1]
        )

        if match:
            round_number = int(match.group(1))

            current_date = datetime.strptime(
                match.group(2),
                "%d/%m/%y"
            ).date()

    # LEFT COURT
    if row[4] in [1, 2]:

        matches.append({
            "round": current_round,
            "round_number": round_number,
            "division": row[4],
            "date": current_date.isoformat(),
            "datetime": datetime.combine(current_date, row[3]).isoformat(),
            "venue": current_venue,

            "home_team": row[5],
            "away_team": row[6],

            "home_club": TEAM_MAP.get(str(row[5]).strip(), row[5]),
            "away_club": TEAM_MAP.get(str(row[6]).strip(), row[6])
        })


    # RIGHT COURT
    if row[9] in [1, 2]:

        matches.append({
            "round": current_round,
            "round_number": round_number,
            "division": row[9],
            "date": current_date.isoformat(),
            "datetime": datetime.combine(current_date, row[3]).isoformat(),
            "venue": current_venue,

            "home_team": row[10],
            "away_team": row[11],

            "home_club": TEAM_MAP.get(str(row[10]).strip(), row[10]),
            "away_club": TEAM_MAP.get(str(row[11]).strip(), row[11])
        })

print(f"Found {len(matches)} matches")


with open("matches.json", "w") as f:
    json.dump(matches, f, indent=4)
