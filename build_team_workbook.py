"""
build_team_workbook.py

Purpose
-------
Turn team_match_deltas.csv (one row per team per match, from build_deltas.py)
and the raw results file (E1/E2/E3.csv, from download_results.py) into one
Excel workbook. League-agnostic — works for Championship, League One, or
League Two, whichever CSVs you point it at (no --league flag needed here):

  - "League Table"     - standard standings (P/W/D/L/GF/GA/GD/Pts) built
                          from every played match in E1.csv, via formulas
                          against a backing "Results Data" sheet.
  - "Overview"          - every team's xG form (season average + rolling),
                          via formulas against each team's own sheet.
  - "Supremacy Delta"   - grid: every team down the left, one column per
                          match ("MW1", "MW2", ...), showing how much their
                          actual xG supremacy (actual for - actual against)
                          beat or missed their pre-match expected supremacy
                          (pre-match for - pre-match against). Rows sorted
                          ascending by season average.
  - "Total xG Delta"    - same grid layout, for how much higher/lower
                          scoring (by combined xG) each match was than the
                          pre-match market expected, regardless of who
                          scored it. Also sorted ascending by average.
  - one sheet per team  - full match log: goals, pre-match market xG,
                          post-match actual xG, deltas, rolling form,
                          Supremacy Delta, and Match Total xG Delta.

"MW1/MW2/..." are each team's own 1st, 2nd, 3rd... match of the season in
date order - not the officially numbered EFL fixture-list gameweek, since
that's not in the source data and postponed/rearranged games can knock a
team's Nth match out of sync with the calendar week. For a normal run of
fixtures they line up; see the cell comment on each grid's "Team" header.

Usage
-----
    python build_team_workbook.py \
        --input data/team_match_deltas_league-one.csv \
        --results data/E2.csv \
        --output data/league-one_xg_report.xlsx \
        --rolling-window 6

The --rolling-window value only needs to match whatever you passed to
build_deltas.py if you want the "current form" labels to show the right
number - it doesn't recompute anything.
"""

import argparse
import re
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# Reuse the same team-name normalisation as build_deltas.py, so the League
# Table (built from E1.csv's own naming, e.g. "QPR") uses the same names as
# the team tabs (built from team_match_deltas.csv, already normalised).
sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_deltas import TEAM_NAME_MAP, normalise_name

FONT_NAME = "Arial"
HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF")
BODY_FONT = Font(name=FONT_NAME)
BOLD_FONT = Font(name=FONT_NAME, bold=True)
SUMMARY_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

GREEN = "63BE7B"
RED = "F8696B"
WHITE = "FFFFFF"

# Column headers for each team's match-log table, in order, mapped to the
# team_match_deltas.csv column they come from. "supremacy_delta" and
# "total_xg_delta" are computed in Excel from the other columns on the same
# row, not read from the CSV.
MATCH_COLUMNS = [
    ("Date", "date"),
    ("Opponent", "opponent"),
    ("Venue", "venue"),
    ("Goals For", "goals_for"),
    ("Goals Against", "goals_against"),
    ("xG Pre-Match (For)", "xg_pre_market_for"),
    ("xG Pre-Match (Against)", "xg_pre_market_against"),
    ("xG Actual (For)", "xg_actual_for"),
    ("xG Actual (Against)", "xg_actual_against"),
    ("xG Delta (For)", "xg_delta_for"),
    ("xG Delta (Against)", "xg_delta_against"),
    ("Rolling Delta (For)", "xg_delta_for_rolling"),
    ("Rolling Delta (Against)", "xg_delta_against_rolling"),
    ("Supremacy Delta", "supremacy_delta"),
    ("Match Total xG Delta", "total_xg_delta"),
]

MW_NOTE = ('"MW" here means this team\'s own 1st, 2nd, 3rd... match of the '
           "season in date order - not the official EFL gameweek number. "
           "They usually line up, but postponed/rearranged fixtures can "
           "knock a team's Nth match out of sync with the calendar week.")


def safe_sheet_name(name: str, used: set) -> str:
    """Excel sheet names: <=31 chars, no \\ / ? * [ ] :, and unique."""
    cleaned = re.sub(r"[\\/?*\[\]:]", "", name)[:31]
    candidate = cleaned
    n = 2
    while candidate in used:
        suffix = f" ({n})"
        candidate = cleaned[: 31 - len(suffix)] + suffix
        n += 1
    used.add(candidate)
    return candidate


def style_header_row(ws: Worksheet, row: int, n_cols: int):
    for col in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
    ws.row_dimensions[row].height = 30


def color_scale(ws: Worksheet, rng: str, invert: bool):
    """invert=False: low=red, high=green (use where higher is better).
    invert=True: low=green, high=red (use where a lower/more negative
    number is the good outcome, e.g. conceding less than expected)."""
    low, high = (GREEN, RED) if invert else (RED, GREEN)
    ws.conditional_formatting.add(
        rng,
        ColorScaleRule(
            start_type="min", start_color=low,
            mid_type="percentile", mid_value=50, mid_color=WHITE,
            end_type="max", end_color=high,
        ),
    )


def autosize_columns(ws: Worksheet, n_cols: int, header_row: int = 1, min_width=4, max_width=24,
                      max_row=None, width_hints: dict | None = None):
    """Width fits the DISPLAYED data in each column, not the header and not
    a formula's literal text (e.g. '=(H2+I2)-(F2+G2)' is 17 characters but
    might display '-1.53' - 5 characters). For columns of literal values
    this is scanned straight from the cells; for formula columns, pass
    width_hints = {1-indexed column: [list of the numbers/strings that
    formula will evaluate to]} computed from the same Python data the
    formula reads, since we can't run Excel here to find out what it
    displays."""
    max_row = max_row if max_row is not None else ws.max_row
    width_hints = width_hints or {}
    for col in range(1, n_cols + 1):
        letter = get_column_letter(col)
        if col in width_hints and width_hints[col]:
            longest = max(len(str(v)) for v in width_hints[col])
        else:
            longest = max(
                (len(str(ws.cell(row=r, column=col).value))
                 for r in range(header_row + 1, max_row + 1)
                 if ws.cell(row=r, column=col).value is not None),
                default=min_width,
            )
        ws.column_dimensions[letter].width = max(min_width, min(max_width, longest + 1))


def write_team_sheet(wb: Workbook, sheet_name: str, team_df: pd.DataFrame, rolling_window: int):
    """Returns a dict of everything the Overview sheet and the two grid
    sheets need to reference this team's data by formula:
        sheet_name, n_matches, col_letter (header -> column letter),
        summary_rows (label -> row number in the season-summary block)."""
    ws = wb.create_sheet(sheet_name)
    team_df = team_df.sort_values("date")

    headers = [h for h, _ in MATCH_COLUMNS]
    n_cols = len(headers)
    ws.append(headers)
    style_header_row(ws, 1, n_cols)

    col_letter = {h: get_column_letter(i + 1) for i, (h, _) in enumerate(MATCH_COLUMNS)}
    supremacy_col = [h for h, _ in MATCH_COLUMNS].index("Supremacy Delta") + 1
    total_delta_col = [h for h, _ in MATCH_COLUMNS].index("Match Total xG Delta") + 1
    supremacy_hints, total_delta_hints = [], []

    pf, pa = col_letter["xG Pre-Match (For)"], col_letter["xG Pre-Match (Against)"]
    af, aa = col_letter["xG Actual (For)"], col_letter["xG Actual (Against)"]

    for row_i, (_, r) in enumerate(team_df.iterrows(), start=2):
        row_values = []
        for header, col in MATCH_COLUMNS:
            if col == "supremacy_delta":
                # (actual for - actual against) - (pre-match for - pre-match
                # against): how much the team's actual xG superiority over
                # their opponent beat (positive) or missed (negative) what
                # the market expected that superiority to be.
                row_values.append(f"=({af}{row_i}-{aa}{row_i})-({pf}{row_i}-{pa}{row_i})")
                supremacy_hints.append(round(
                    (r["xg_actual_for"] - r["xg_actual_against"])
                    - (r["xg_pre_market_for"] - r["xg_pre_market_against"]), 2))
                continue
            if col == "total_xg_delta":
                # (actual for + actual against) - (pre-match for + pre-match
                # against): how much higher/lower-scoring (by xG) this match
                # was than the market expected, regardless of which side.
                row_values.append(f"=({af}{row_i}+{aa}{row_i})-({pf}{row_i}+{pa}{row_i})")
                total_delta_hints.append(round(
                    (r["xg_actual_for"] + r["xg_actual_against"])
                    - (r["xg_pre_market_for"] + r["xg_pre_market_against"]), 2))
                continue
            val = r[col]
            if col == "date":
                val = pd.to_datetime(val).strftime("%Y-%m-%d")
            elif col == "venue":
                val = str(val).title()
            elif isinstance(val, float):
                val = round(val, 2)
            row_values.append(val)
        ws.append(row_values)

    last_row = ws.max_row
    n_matches = last_row - 1
    decimal_cols = {i + 1 for i, (_, col) in enumerate(MATCH_COLUMNS)
                     if col not in ("date", "opponent", "venue", "goals_for", "goals_against")}
    for row in ws.iter_rows(min_row=2, max_row=last_row, max_col=n_cols):
        for cell in row:
            cell.font = BODY_FONT
            if cell.column in decimal_cols:
                cell.number_format = "0.00"

    # Colour-scale every delta column. "(For)" and "Supremacy Delta": green
    # = better than expected (higher). "(Against)" columns: green = better
    # than expected (LOWER / more negative - conceded less than expected).
    # "Match Total xG Delta" is left uncoloured - it's a fact about the
    # match, not a team performance stat with a "good" direction.
    for header in ("xG Delta (For)", "Rolling Delta (For)", "Supremacy Delta"):
        letter = col_letter[header]
        color_scale(ws, f"{letter}2:{letter}{last_row}", invert=False)
    for header in ("xG Delta (Against)", "Rolling Delta (Against)"):
        letter = col_letter[header]
        color_scale(ws, f"{letter}2:{letter}{last_row}", invert=True)

    # Summary block: real formulas over the table above, not re-typed
    # numbers, so it stays correct if you ever edit a row by hand. Labels
    # are merged across A:D and the value sits in E, so long label text
    # doesn't force the narrow Date/Opponent/Venue/Goals columns wide.
    summary_row = last_row + 2
    ws.cell(row=summary_row, column=1, value="Season summary").font = BOLD_FONT

    def col_range(header_name):
        letter = col_letter[header_name]
        return f"{letter}2:{letter}{last_row}"

    labels_formulas = [
        ("Matches played", f"=COUNTA({col_range('Date')})"),
        ("Total goals for", f"=SUM({col_range('Goals For')})"),
        ("Total goals against", f"=SUM({col_range('Goals Against')})"),
        ("Avg xG Delta (For)", f"=AVERAGE({col_range('xG Delta (For)')})"),
        ("Avg xG Delta (Against)", f"=AVERAGE({col_range('xG Delta (Against)')})"),
        ("Avg Supremacy Delta", f"=AVERAGE({col_range('Supremacy Delta')})"),
        ("Avg Total xG Delta", f"=AVERAGE({col_range('Match Total xG Delta')})"),
        (f"Current form (last {rolling_window}, For)", f"={col_letter['Rolling Delta (For)']}{last_row}"),
        (f"Current form (last {rolling_window}, Against)", f"={col_letter['Rolling Delta (Against)']}{last_row}"),
    ]
    integer_labels = {"Matches played", "Total goals for", "Total goals against"}
    summary_rows = {}
    for i, (label, formula) in enumerate(labels_formulas):
        r = summary_row + 1 + i
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        label_cell = ws.cell(row=r, column=1, value=label)
        label_cell.font = BODY_FONT
        value_cell = ws.cell(row=r, column=5, value=formula)
        value_cell.font = BODY_FONT
        value_cell.fill = SUMMARY_FILL
        value_cell.number_format = "0" if label in integer_labels else "0.00"
        summary_rows[label] = r

    ws.freeze_panes = "A2"
    # Width is based only on the match-table rows (not the summary block
    # below, whose long labels live in a merged A:D cell) so the data
    # columns stay narrow - just wide enough for the numbers. Formula
    # columns get their width from the actual numbers they'll evaluate to
    # (computed here in Python), not the formula text itself.
    autosize_columns(ws, n_cols, header_row=1, max_row=last_row, width_hints={
        supremacy_col: supremacy_hints,
        total_delta_col: total_delta_hints,
    })

    return {
        "sheet_name": sheet_name,
        "n_matches": n_matches,
        "col_letter": col_letter,
        "summary_rows": summary_rows,
    }


def write_overview_sheet(wb: Workbook, team_info: dict, team_stats: dict, rolling_window: int):
    ws = wb.create_sheet("Overview")
    headers = [
        "Team", "Matches Played", "Total Goals For", "Total Goals Against",
        "Avg xG Delta (For)", "Avg xG Delta (Against)",
        f"Current Form ({rolling_window}, For)", f"Current Form ({rolling_window}, Against)",
    ]
    ws.append(headers)
    style_header_row(ws, 1, len(headers))

    stat_keys = ["matches", "gf", "ga", "avg_for", "avg_against", "form_for", "form_against"]
    label_keys = ["Matches played", "Total goals for", "Total goals against",
                  "Avg xG Delta (For)", "Avg xG Delta (Against)"]
    width_hints = {i + 2: [] for i in range(len(stat_keys))}
    for team in sorted(team_info):
        info = team_info[team]
        sheet_name, sr = info["sheet_name"], info["summary_rows"]
        rolling_for_label = f"Current form (last {rolling_window}, For)"
        rolling_against_label = f"Current form (last {rolling_window}, Against)"
        row_labels = label_keys + [rolling_for_label, rolling_against_label]
        ws.append([team] + [f"='{sheet_name}'!$E${sr[lbl]}" for lbl in row_labels])
        for i, key in enumerate(stat_keys):
            width_hints[i + 2].append(team_stats[team][key])

    last_row = ws.max_row
    decimal_overview_cols = {5, 6, 7, 8}
    for row in ws.iter_rows(min_row=2, max_row=last_row, max_col=len(headers)):
        for cell in row:
            cell.font = BODY_FONT
            if cell.column in decimal_overview_cols:
                cell.number_format = "0.00"

    for col_idx, invert in ((5, False), (6, True), (7, False), (8, True)):
        col_letter = get_column_letter(col_idx)
        rng = f"{col_letter}2:{col_letter}{last_row}"
        color_scale(ws, rng, invert=invert)

    ws.freeze_panes = "A2"
    autosize_columns(ws, len(headers), width_hints=width_hints)
    ws.column_dimensions["A"].width = max(ws.column_dimensions["A"].width, 12)  # Team


def write_results_data_sheet(wb: Workbook, results_df: pd.DataFrame):
    """Raw backing data for the League Table - one row per played match.
    Kept as plain data (it's a straight import of source results), with
    the League Table's P/W/D/L/GF/GA/Pts all computed from it via formulas."""
    ws = wb.create_sheet("Results Data")
    ws.sheet_state = "hidden"
    headers = ["Date", "HomeTeam", "AwayTeam", "FTHG", "FTAG"]
    ws.append(headers)
    style_header_row(ws, 1, len(headers))
    for _, r in results_df.iterrows():
        ws.append([
            pd.to_datetime(r["Date"], dayfirst=True).strftime("%Y-%m-%d"),
            normalise_name(r["HomeTeam"], TEAM_NAME_MAP),
            normalise_name(r["AwayTeam"], TEAM_NAME_MAP),
            int(r["FTHG"]), int(r["FTAG"]),
        ])
    autosize_columns(ws, len(headers))
    return ws.max_row  # last data row


def write_league_table_sheet(wb: Workbook, results_df: pd.DataFrame, last_results_row: int):
    """Standard standings (P/W/D/L/GF/GA/GD/Pts), row order decided in
    Python (points, then goal difference, then goals for - the standard
    tie-break) but every number a formula against the 'Results Data' sheet,
    so it stays correct if that data is edited."""
    home = results_df.rename(columns={"HomeTeam": "team", "AwayTeam": "opponent",
                                       "FTHG": "gf", "FTAG": "ga"})[["team", "opponent", "gf", "ga"]]
    away = results_df.rename(columns={"AwayTeam": "team", "HomeTeam": "opponent",
                                       "FTAG": "gf", "FTHG": "ga"})[["team", "opponent", "gf", "ga"]]
    long_df = pd.concat([home, away], ignore_index=True)
    long_df["team"] = long_df["team"].apply(lambda n: normalise_name(n, TEAM_NAME_MAP))
    long_df["points"] = long_df.apply(lambda r: 3 if r.gf > r.ga else (1 if r.gf == r.ga else 0), axis=1)
    long_df["result"] = long_df.apply(lambda r: "W" if r.gf > r.ga else ("D" if r.gf == r.ga else "L"), axis=1)
    standings = (
        long_df.groupby("team")
        .agg(played=("gf", "count"), gf=("gf", "sum"), ga=("ga", "sum"), points=("points", "sum"),
             won=("result", lambda s: (s == "W").sum()),
             drawn=("result", lambda s: (s == "D").sum()),
             lost=("result", lambda s: (s == "L").sum()))
        .reset_index()
    )
    standings["gd"] = standings["gf"] - standings["ga"]
    standings = standings.sort_values(["points", "gd", "gf"], ascending=False).reset_index(drop=True)

    ws = wb.create_sheet("League Table")
    headers = ["Pos", "Team", "P", "W", "D", "L", "GF", "GA", "GD", "Pts"]
    ws.append(headers)
    style_header_row(ws, 1, len(headers))

    def rd_range(letter):
        return f"'Results Data'!${letter}$2:${letter}${last_results_row}"

    home_team_rng = rd_range("B")
    away_team_rng = rd_range("C")
    home_gf_rng = rd_range("D")
    away_gf_rng = rd_range("E")

    for i, row in enumerate(standings.itertuples(), start=2):
        team = row.team
        ws.cell(row=i, column=1, value=f"=ROW()-1")                       # Pos
        ws.cell(row=i, column=2, value=team)                              # Team
        ws.cell(row=i, column=3,                                          # P
                value=f"=COUNTIF({home_team_rng},$B{i})+COUNTIF({away_team_rng},$B{i})")
        ws.cell(row=i, column=4,                                          # W
                value=(f"=SUMPRODUCT(({home_team_rng}=$B{i})*({home_gf_rng}>{away_gf_rng}))"
                       f"+SUMPRODUCT(({away_team_rng}=$B{i})*({away_gf_rng}>{home_gf_rng}))"))
        ws.cell(row=i, column=5,                                          # D
                value=(f"=SUMPRODUCT(({home_team_rng}=$B{i})*({home_gf_rng}={away_gf_rng}))"
                       f"+SUMPRODUCT(({away_team_rng}=$B{i})*({away_gf_rng}={home_gf_rng}))"))
        ws.cell(row=i, column=6, value=f"=C{i}-D{i}-E{i}")                # L
        ws.cell(row=i, column=7,                                          # GF
                value=f"=SUMIF({home_team_rng},$B{i},{home_gf_rng})+SUMIF({away_team_rng},$B{i},{away_gf_rng})")
        ws.cell(row=i, column=8,                                          # GA
                value=f"=SUMIF({home_team_rng},$B{i},{away_gf_rng})+SUMIF({away_team_rng},$B{i},{home_gf_rng})")
        ws.cell(row=i, column=9, value=f"=G{i}-H{i}")                     # GD
        ws.cell(row=i, column=10, value=f"=D{i}*3+E{i}")                  # Pts

    last_row = ws.max_row
    for row in ws.iter_rows(min_row=2, max_row=last_row, max_col=len(headers)):
        for cell in row:
            cell.font = BODY_FONT
            if cell.column == 1:
                cell.alignment = Alignment(horizontal="center")

    ws.freeze_panes = "A2"
    width_hints = {
        1: [len(standings)],                                     # Pos
        3: standings["played"].tolist(),                         # P
        4: standings["won"].tolist(),                             # W
        5: standings["drawn"].tolist(),                           # D
        6: standings["lost"].tolist(),                            # L
        7: standings["gf"].tolist(),                              # GF
        8: standings["ga"].tolist(),                              # GA
        9: standings["gd"].tolist(),                              # GD
        10: standings["points"].tolist(),                         # Pts
    }
    autosize_columns(ws, len(headers), width_hints=width_hints)
    ws.column_dimensions["B"].width = max(ws.column_dimensions["B"].width, 14)  # Team
    return ws


def write_matrix_sheet(wb: Workbook, sheet_name: str, team_info: dict,
                        values_by_team: dict, avg_by_team: dict,
                        source_col_header: str, avg_summary_label: str,
                        colored: bool, invert: bool):
    """Teams down the left, one column per match ("MW1", "MW2", ...), plus
    an Average column. Row order: ascending by season average. Every match
    cell is a formula referencing that team's own sheet (source_col_header
    column, that row), so it updates if the underlying data changes;
    values_by_team/avg_by_team (plain Python numbers) are only used to
    decide row order and column widths."""
    max_matches = max((len(v) for v in values_by_team.values()), default=0)
    sorted_teams = sorted(values_by_team.keys(), key=lambda t: avg_by_team[t])

    ws = wb.create_sheet(sheet_name)
    headers = ["Team", "Average"] + [f"MW{i}" for i in range(1, max_matches + 1)]
    ws.append(headers)
    style_header_row(ws, 1, len(headers))
    ws.cell(row=1, column=1).comment = Comment(MW_NOTE, "build_team_workbook.py")

    width_hints = {2: [avg_by_team[t] for t in sorted_teams]}
    for mw in range(1, max_matches + 1):
        width_hints[mw + 2] = [v[mw - 1] for v in values_by_team.values() if len(v) >= mw]

    for row_i, team in enumerate(sorted_teams, start=2):
        info = team_info[team]
        sheet_ref = info["sheet_name"]
        avg_row = info["summary_rows"][avg_summary_label]
        metric_letter = info["col_letter"][source_col_header]
        n = info["n_matches"]

        ws.cell(row=row_i, column=1, value=team).font = BODY_FONT
        avg_cell = ws.cell(row=row_i, column=2, value=f"='{sheet_ref}'!$E${avg_row}")
        avg_cell.font = BOLD_FONT
        avg_cell.number_format = "0.00"
        avg_cell.fill = SUMMARY_FILL

        for mw in range(1, max_matches + 1):
            col = mw + 2
            cell = ws.cell(row=row_i, column=col)
            if mw <= n:
                data_row = mw + 1  # row 2 on the team sheet = MW1
                cell.value = f"='{sheet_ref}'!${metric_letter}${data_row}"
                cell.number_format = "0.00"
            cell.font = BODY_FONT

    last_row = ws.max_row
    last_col_letter = get_column_letter(len(headers))
    if colored and last_row >= 2:
        color_scale(ws, f"B2:{last_col_letter}{last_row}", invert=invert)

    ws.freeze_panes = "C2"  # keep Team + Average visible while scrolling across matchweeks
    autosize_columns(ws, len(headers), width_hints=width_hints)
    ws.column_dimensions["A"].width = max(ws.column_dimensions["A"].width, 20)  # Team
    return ws


def main(input_path, results_path, output_path, rolling_window):
    df = pd.read_csv(input_path, parse_dates=["date"])
    teams = sorted(df["team"].unique())

    wb = Workbook()
    wb.remove(wb.active)  # drop the default blank sheet

    if results_path:
        results_df = pd.read_csv(results_path)
        results_df = results_df.dropna(subset=["FTHG", "FTAG"])  # only played matches
        last_results_row = write_results_data_sheet(wb, results_df)
        write_league_table_sheet(wb, results_df, last_results_row)

    used_names = set()
    team_info = {}
    team_stats = {}
    supremacy_by_team, total_delta_by_team = {}, {}
    for team in teams:
        sheet_name = safe_sheet_name(team, used_names)
        team_df = df[df["team"] == team].sort_values("date")
        team_info[team] = write_team_sheet(wb, sheet_name, team_df, rolling_window)

        supremacy_series = ((team_df["xg_actual_for"] - team_df["xg_actual_against"])
                             - (team_df["xg_pre_market_for"] - team_df["xg_pre_market_against"])).round(2)
        total_series = ((team_df["xg_actual_for"] + team_df["xg_actual_against"])
                         - (team_df["xg_pre_market_for"] + team_df["xg_pre_market_against"])).round(2)
        supremacy_by_team[team] = supremacy_series.tolist()
        total_delta_by_team[team] = total_series.tolist()

        team_stats[team] = {
            "matches": len(team_df),
            "gf": int(team_df["goals_for"].sum()),
            "ga": int(team_df["goals_against"].sum()),
            "avg_for": round(team_df["xg_delta_for"].mean(), 2),
            "avg_against": round(team_df["xg_delta_against"].mean(), 2),
            "form_for": round(team_df["xg_delta_for_rolling"].iloc[-1], 2),
            "form_against": round(team_df["xg_delta_against_rolling"].iloc[-1], 2),
            "avg_supremacy": round(supremacy_series.mean(), 2),
            "avg_total_delta": round(total_series.mean(), 2),
        }

    write_overview_sheet(wb, team_info, team_stats, rolling_window)

    avg_supremacy_by_team = {t: team_stats[t]["avg_supremacy"] for t in teams}
    avg_total_by_team = {t: team_stats[t]["avg_total_delta"] for t in teams}
    write_matrix_sheet(wb, "Supremacy Delta", team_info, supremacy_by_team, avg_supremacy_by_team,
                        source_col_header="Supremacy Delta", avg_summary_label="Avg Supremacy Delta",
                        colored=True, invert=False)
    write_matrix_sheet(wb, "Total xG Delta", team_info, total_delta_by_team, avg_total_by_team,
                        source_col_header="Match Total xG Delta", avg_summary_label="Avg Total xG Delta",
                        colored=False, invert=False)

    # Sheet order: League Table, Overview, the two grids, then teams A-Z
    # (Results Data stays hidden but present).
    front = [n for n in ("League Table", "Overview", "Supremacy Delta", "Total xG Delta") if n in wb.sheetnames]
    order = front + [n for n in wb.sheetnames if n not in front and n != "Results Data"]
    if "Results Data" in wb.sheetnames:
        order.append("Results Data")
    wb._sheets = [wb[n] for n in order]

    wb.save(output_path)
    print(f"Wrote workbook with {len(teams)} team sheets + League Table + Overview + "
          f"Supremacy Delta + Total xG Delta to {output_path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", default="data/team_match_deltas.csv")
    parser.add_argument("--results", default="data/E1.csv",
                         help="Raw results file (for the League Table). Pass '' to skip the League Table.")
    parser.add_argument("--output", default="data/championship_xg_report.xlsx")
    parser.add_argument("--rolling-window", type=int, default=6,
                         help="Only used to label the 'current form' column - must match what you passed to build_deltas.py")
    args = parser.parse_args()

    main(args.input, args.results or None, args.output, args.rolling_window)
