"""
build_team_workbook.py

Purpose
-------
Turn team_match_deltas.csv (one row per team per match, from build_deltas.py),
the raw results file (E1/E2/E3/... .csv, from download_results.py), and
optionally an upcoming-fixtures pre-match xG file (from download_fixtures.py
+ derive_prematch_xg.py) into one Excel workbook. League-agnostic — works for
any league whose CSVs you point it at (no --league flag needed here):

  - "About This Report" - a short front page explaining every sheet and the
                          handful of judgement calls made where the brief
                          for this report was ambiguous (sort order, what
                          "supremacy" means, etc.) - read this first.
  - "League Table"       - standard standings (P/W/D/L/GF/GA/GD/Pts) built
                          from every played match in the results file, via
                          formulas against a backing "Results Data" sheet.
  - "Overview"            - every team's xG form (season average + median +
                          rolling), via formulas against each team's own
                          sheet. Sorted by Avg xG Delta (For), best first.
  - "Home & Away Splits"  - league-wide Total/Home/Away breakdown of average
                          + median supremacy, and Total/Home/Away goals for
                          and against, one row per team.
  - "Supremacy Delta"     - grid: every team down the left, one column per
                          match, showing how much their actual xG supremacy
                          (actual for - actual against) beat or missed their
                          pre-match expected supremacy (pre-match for -
                          pre-match against). Sorted by Avg xG Delta (For),
                          best first (see the About sheet for why).
  - "Supremacy Delta (Home)" / "Supremacy Delta (Away)" - the same grid,
                          restricted to each team's home-only / away-only
                          matches.
  - "Total xG Delta"      - same grid layout, for how much higher/lower
                          scoring (by combined xG) each match was than the
                          pre-match market expected, regardless of who
                          scored it. Colour-scaled (green = higher-scoring
                          than expected, red = lower) and sorted by average,
                          highest first.
  - "Upcoming Fixtures"   - only written if a fixtures pre-match xG file is
                          passed with --fixtures. Every fixture the market
                          has odds for right now, ranked by an "expected
                          supremacy" that blends the market's own pre-match
                          view of that fixture with how much each side has
                          been over/under-performing their own expectation
                          all season (see the About sheet for the exact
                          formula).
  - one sheet per team    - full match log: goals, pre-match market xG,
                          post-match actual xG, deltas, rolling form,
                          Supremacy Delta, Match Total xG Delta; a season
                          summary block (average AND median of every delta
                          stat); and a Home/Away split mini-table.

"MW1/MW2/..." (and "H1/H2.../A1/A2...") are each team's own 1st, 2nd, 3rd...
match of the season (or of that venue) in date order - not the officially
numbered fixture-list gameweek, since that's not in the source data and
postponed/rearranged games can knock a team's Nth match out of sync with the
calendar week. For a normal run of fixtures they line up; see the cell
comment on each grid's "Team" header.

Usage
-----
    python build_team_workbook.py \
        --input data/team_match_deltas_league-one.csv \
        --results data/E2.csv \
        --fixtures data/prematch_xg_fixtures_league-one.csv \
        --output data/league-one_xg_report.xlsx \
        --rolling-window 6

--fixtures is optional - omit it (or pass '') and the workbook is built
without an "Upcoming Fixtures" sheet, same as before this feature existed.
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
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# Reuse the same team-name normalisation as build_deltas.py, so every sheet
# (built from E1.csv-style naming, e.g. "QPR") uses the same names as the
# team tabs (built from team_match_deltas.csv, already normalised).
sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_deltas import TEAM_NAME_MAP, normalise_name

FONT_NAME = "Arial"
HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF")
BODY_FONT = Font(name=FONT_NAME)
BOLD_FONT = Font(name=FONT_NAME, bold=True)
TITLE_FONT = Font(name=FONT_NAME, bold=True, size=16, color="1F4E78")
SUBTITLE_FONT = Font(name=FONT_NAME, italic=True, size=10, color="595959")
SUMMARY_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
BAND_FILL = PatternFill(start_color="F2F6FA", end_color="F2F6FA", fill_type="solid")
SECTION_FILL = PatternFill(start_color="8EA9C1", end_color="8EA9C1", fill_type="solid")
THIN = Side(style="thin", color="BFBFBF")
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

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
           "season in date order - not the official league gameweek number. "
           "They usually line up, but postponed/rearranged fixtures can "
           "knock a team's Nth match out of sync with the calendar week.")

ABOUT_NOTE = (
    "Where the brief for this report left a choice open, this workbook "
    "makes one consistently rather than guessing differently sheet to "
    "sheet - see the About This Report sheet for exactly which calls "
    "were made (sort orders, what \"expected supremacy\" means for "
    "upcoming fixtures, etc.)."
)


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
        cell.border = THIN_BORDER
    ws.row_dimensions[row].height = 30


def color_scale(ws: Worksheet, rng: str, invert: bool):
    """Percentile-based scale: invert=False: low=red, high=green (use where
    higher is better). invert=True: low=green, high=red (use where a
    lower/more negative number is the good outcome, e.g. conceding less
    than expected)."""
    low, high = (GREEN, RED) if invert else (RED, GREEN)
    ws.conditional_formatting.add(
        rng,
        ColorScaleRule(
            start_type="min", start_color=low,
            mid_type="percentile", mid_value=50, mid_color=WHITE,
            end_type="max", end_color=high,
        ),
    )


def color_scale_zero(ws: Worksheet, rng: str, invert: bool = False):
    """Zero-anchored scale for delta-style stats where 0 = "exactly as
    expected" is meaningful on its own, not just relative to the rest of
    the column: positive -> shades of green, negative -> shades of red,
    with white pinned at zero itself (not the column's median/percentile,
    which is what plain color_scale() above anchors to). invert=True flips
    it for stats where negative is the good outcome (e.g. "Against" columns
    - conceding less than expected)."""
    low, high = (GREEN, RED) if invert else (RED, GREEN)
    ws.conditional_formatting.add(
        rng,
        ColorScaleRule(
            start_type="min", start_color=low,
            mid_type="num", mid_value=0, mid_color=WHITE,
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


def polish_sheet(ws: Worksheet, header_row: int, first_data_row: int, last_data_row: int, n_cols: int,
                  band: bool = True):
    """Shared professional-look pass, applied after a sheet's data is
    written: no gridlines (the borders below give structure instead), a
    thin border around every header/data cell, and (optionally) very light
    banding on alternate data rows so wide tables stay readable. Doesn't
    touch fonts/number formats - those are set per-sheet since they vary."""
    ws.sheet_view.showGridLines = False
    for col in range(1, n_cols + 1):
        ws.cell(row=header_row, column=col).border = THIN_BORDER
    for row in range(first_data_row, last_data_row + 1):
        for col in range(1, n_cols + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = THIN_BORDER
            if band and cell.fill.start_color.rgb in (None, "00000000") and (row - first_data_row) % 2 == 1:
                cell.fill = BAND_FILL


def stat_block(series: pd.Series) -> dict:
    """Average + median for a pandas Series, rounded, with a safe fallback
    (0.0) for an empty series (e.g. a team with no away matches yet)."""
    if len(series) == 0:
        return {"avg": 0.0, "median": 0.0}
    return {"avg": round(float(series.mean()), 2), "median": round(float(series.median()), 2)}


def write_team_sheet(wb: Workbook, sheet_name: str, team_df: pd.DataFrame, rolling_window: int):
    """Returns a dict of everything the Overview/Home & Away Splits/grid
    sheets need to reference this team's data by formula or by value:
        sheet_name, n_matches, col_letter (header -> column letter),
        summary_rows (label -> row number in the season-summary block),
        stats (plain-Python season/home/away summary numbers)."""
    ws = wb.create_sheet(sheet_name)
    team_df = team_df.sort_values("date").reset_index(drop=True)

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

    # Computed once here (not read back from the CSV) so the season-summary
    # block, the Home/Away split table, and this team's contribution to the
    # league-wide sheets all agree with what's actually on the rows above.
    supremacy_full = ((team_df["xg_actual_for"] - team_df["xg_actual_against"])
                       - (team_df["xg_pre_market_for"] - team_df["xg_pre_market_against"])).round(2)
    total_full = ((team_df["xg_actual_for"] + team_df["xg_actual_against"])
                  - (team_df["xg_pre_market_for"] + team_df["xg_pre_market_against"])).round(2)

    for row_i, (_, r) in enumerate(team_df.iterrows(), start=2):
        row_values = []
        for header, col in MATCH_COLUMNS:
            if col == "supremacy_delta":
                # (actual for - actual against) - (pre-match for - pre-match
                # against): how much the team's actual xG superiority over
                # their opponent beat (positive) or missed (negative) what
                # the market expected that superiority to be.
                row_values.append(f"=({af}{row_i}-{aa}{row_i})-({pf}{row_i}-{pa}{row_i})")
                supremacy_hints.append(supremacy_full.iloc[row_i - 2])
                continue
            if col == "total_xg_delta":
                # (actual for + actual against) - (pre-match for + pre-match
                # against): how much higher/lower-scoring (by xG) this match
                # was than the market expected, regardless of which side.
                row_values.append(f"=({af}{row_i}+{aa}{row_i})-({pf}{row_i}+{pa}{row_i})")
                total_delta_hints.append(total_full.iloc[row_i - 2])
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

    # Colour-scale every delta column, zero-anchored (0 = exactly as
    # expected) rather than relative to the column's own spread. "(For)"
    # and "Supremacy Delta": green = better than expected (higher).
    # "(Against)" columns: green = better than expected (LOWER / more
    # negative - conceded less than expected). "Match Total xG Delta" is
    # coloured too now (green = higher-scoring match than the market
    # expected, red = lower) - it's a fact about the match rather than a
    # team performance stat, but still meaningfully signed around zero.
    for header in ("xG Delta (For)", "Rolling Delta (For)", "Supremacy Delta"):
        letter = col_letter[header]
        color_scale_zero(ws, f"{letter}2:{letter}{last_row}", invert=False)
    for header in ("xG Delta (Against)", "Rolling Delta (Against)"):
        letter = col_letter[header]
        color_scale_zero(ws, f"{letter}2:{letter}{last_row}", invert=True)
    letter = col_letter["Match Total xG Delta"]
    color_scale_zero(ws, f"{letter}2:{letter}{last_row}", invert=False)

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
        ("Matches played", f"=COUNTA({col_range('Date')})", True),
        ("Total goals for", f"=SUM({col_range('Goals For')})", True),
        ("Total goals against", f"=SUM({col_range('Goals Against')})", True),
        ("Avg xG Delta (For)", f"=AVERAGE({col_range('xG Delta (For)')})", False),
        ("Median xG Delta (For)", f"=MEDIAN({col_range('xG Delta (For)')})", False),
        ("Avg xG Delta (Against)", f"=AVERAGE({col_range('xG Delta (Against)')})", False),
        ("Median xG Delta (Against)", f"=MEDIAN({col_range('xG Delta (Against)')})", False),
        ("Avg Supremacy Delta", f"=AVERAGE({col_range('Supremacy Delta')})", False),
        ("Median Supremacy Delta", f"=MEDIAN({col_range('Supremacy Delta')})", False),
        ("Avg Total xG Delta", f"=AVERAGE({col_range('Match Total xG Delta')})", False),
        ("Median Total xG Delta", f"=MEDIAN({col_range('Match Total xG Delta')})", False),
        (f"Current form (last {rolling_window}, For)", f"={col_letter['Rolling Delta (For)']}{last_row}", False),
        (f"Current form (last {rolling_window}, Against)", f"={col_letter['Rolling Delta (Against)']}{last_row}", False),
    ]
    summary_rows = {}
    for i, (label, formula, is_integer) in enumerate(labels_formulas):
        r = summary_row + 1 + i
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        label_cell = ws.cell(row=r, column=1, value=label)
        label_cell.font = BODY_FONT
        value_cell = ws.cell(row=r, column=5, value=formula)
        value_cell.font = BODY_FONT
        value_cell.fill = SUMMARY_FILL
        value_cell.number_format = "0" if is_integer else "0.00"
        summary_rows[label] = r

    # Home / Away split mini-table: literal values (computed in Python from
    # the rows above), not formulas - a MEDIAN-of-a-condition needs an
    # array formula that's fragile across Excel versions, and these three
    # rows are a read-only snapshot rather than something you'd hand-edit.
    home_mask = team_df["venue"] == "home"
    away_mask = team_df["venue"] == "away"
    split_row = summary_rows[labels_formulas[-1][0]] + 2
    ws.cell(row=split_row, column=1, value="Home / Away split").font = BOLD_FONT
    split_headers = ["Venue", "Matches", "GF", "GA", "Avg Supremacy", "Median Supremacy",
                      "Avg Total xG Delta", "Median Total xG Delta"]
    for c, h in enumerate(split_headers, start=1):
        cell = ws.cell(row=split_row + 1, column=c, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    home_split = stat_block(supremacy_full[home_mask])
    away_split = stat_block(supremacy_full[away_mask])
    home_total_split = stat_block(total_full[home_mask])
    away_total_split = stat_block(total_full[away_mask])
    venue_rows = [
        ("Home", int(home_mask.sum()), int(team_df.loc[home_mask, "goals_for"].sum()),
         int(team_df.loc[home_mask, "goals_against"].sum()),
         home_split["avg"], home_split["median"], home_total_split["avg"], home_total_split["median"]),
        ("Away", int(away_mask.sum()), int(team_df.loc[away_mask, "goals_for"].sum()),
         int(team_df.loc[away_mask, "goals_against"].sum()),
         away_split["avg"], away_split["median"], away_total_split["avg"], away_total_split["median"]),
        ("Total", n_matches, int(team_df["goals_for"].sum()), int(team_df["goals_against"].sum()),
         stat_block(supremacy_full)["avg"], stat_block(supremacy_full)["median"],
         stat_block(total_full)["avg"], stat_block(total_full)["median"]),
    ]
    for r_off, row_vals in enumerate(venue_rows, start=1):
        r = split_row + 1 + r_off
        for c, v in enumerate(row_vals, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = BOLD_FONT if row_vals[0] == "Total" else BODY_FONT
            if c >= 5:
                cell.number_format = "0.00"
            if row_vals[0] == "Total":
                cell.fill = SUMMARY_FILL

    ws.freeze_panes = "A2"
    # Width is based only on the match-table rows (not the summary/split
    # blocks below, whose labels live in their own header row) so the data
    # columns stay narrow - just wide enough for the numbers. Formula
    # columns get their width from the actual numbers they'll evaluate to
    # (computed here in Python), not the formula text itself.
    autosize_columns(ws, n_cols, header_row=1, max_row=last_row, width_hints={
        supremacy_col: supremacy_hints,
        total_delta_col: total_delta_hints,
    })
    polish_sheet(ws, header_row=1, first_data_row=2, last_data_row=last_row, n_cols=n_cols)

    stats = {
        "matches": n_matches,
        "gf": int(team_df["goals_for"].sum()),
        "ga": int(team_df["goals_against"].sum()),
        "avg_for": round(team_df["xg_delta_for"].mean(), 2),
        "median_for": round(team_df["xg_delta_for"].median(), 2),
        "avg_against": round(team_df["xg_delta_against"].mean(), 2),
        "median_against": round(team_df["xg_delta_against"].median(), 2),
        "form_for": round(team_df["xg_delta_for_rolling"].iloc[-1], 2),
        "form_against": round(team_df["xg_delta_against_rolling"].iloc[-1], 2),
        "avg_supremacy": stat_block(supremacy_full)["avg"],
        "median_supremacy": stat_block(supremacy_full)["median"],
        "avg_total_delta": stat_block(total_full)["avg"],
        "median_total_delta": stat_block(total_full)["median"],
        "home_matches": int(home_mask.sum()),
        "home_gf": int(team_df.loc[home_mask, "goals_for"].sum()),
        "home_ga": int(team_df.loc[home_mask, "goals_against"].sum()),
        "home_avg_supremacy": home_split["avg"],
        "home_median_supremacy": home_split["median"],
        "away_matches": int(away_mask.sum()),
        "away_gf": int(team_df.loc[away_mask, "goals_for"].sum()),
        "away_ga": int(team_df.loc[away_mask, "goals_against"].sum()),
        "away_avg_supremacy": away_split["avg"],
        "away_median_supremacy": away_split["median"],
    }

    return {
        "sheet_name": sheet_name,
        "n_matches": n_matches,
        "col_letter": col_letter,
        "summary_rows": summary_rows,
        "stats": stats,
        # Full-season / home-only / away-only supremacy & total-delta
        # series in date order, for the league-wide grid sheets.
        "supremacy_series": supremacy_full.tolist(),
        "total_series": total_full.tolist(),
        "home_supremacy_series": supremacy_full[home_mask].tolist(),
        "away_supremacy_series": supremacy_full[away_mask].tolist(),
    }


def write_overview_sheet(wb: Workbook, team_info: dict, rolling_window: int):
    """Ranks every team by Avg xG Delta (For), best first - see the About
    This Report sheet for why "in order of xG for" was read that way."""
    ws = wb.create_sheet("Overview")
    headers = [
        "Team", "Matches Played", "Total Goals For", "Total Goals Against",
        "Avg xG Delta (For)", "Median xG Delta (For)",
        "Avg xG Delta (Against)", "Median xG Delta (Against)",
        "Avg Supremacy Delta", "Median Supremacy Delta",
        "Avg Total xG Delta", "Median Total xG Delta",
        f"Current Form ({rolling_window}, For)", f"Current Form ({rolling_window}, Against)",
    ]
    ws.append(headers)
    style_header_row(ws, 1, len(headers))

    row_order = sorted(team_info.keys(), key=lambda t: team_info[t]["stats"]["avg_for"], reverse=True)

    stat_keys = ["matches", "gf", "ga", "avg_for", "median_for", "avg_against", "median_against",
                 "avg_supremacy", "median_supremacy", "avg_total_delta", "median_total_delta",
                 "form_for", "form_against"]
    width_hints = {i + 2: [] for i in range(len(stat_keys))}
    for team in row_order:
        info = team_info[team]
        sheet_name, sr, s = info["sheet_name"], info["summary_rows"], info["stats"]
        row_labels = [
            "Matches played", "Total goals for", "Total goals against",
            "Avg xG Delta (For)", "Median xG Delta (For)",
            "Avg xG Delta (Against)", "Median xG Delta (Against)",
            "Avg Supremacy Delta", "Median Supremacy Delta",
            "Avg Total xG Delta", "Median Total xG Delta",
            f"Current form (last {rolling_window}, For)", f"Current form (last {rolling_window}, Against)",
        ]
        ws.append([team] + [f"='{sheet_name}'!$E${sr[lbl]}" for lbl in row_labels])
        for i, key in enumerate(stat_keys):
            width_hints[i + 2].append(s[key])

    last_row = ws.max_row
    decimal_overview_cols = set(range(5, len(headers) + 1))
    for row in ws.iter_rows(min_row=2, max_row=last_row, max_col=len(headers)):
        for cell in row:
            cell.font = BODY_FONT
            if cell.column in decimal_overview_cols:
                cell.number_format = "0.00"

    zero_centered_cols = {
        5: False, 6: False,    # xG Delta (For) avg/median: higher = better
        7: True, 8: True,      # xG Delta (Against) avg/median: lower = better
        9: False, 10: False,   # Supremacy Delta avg/median: higher = better
        11: False, 12: False,  # Total xG Delta avg/median: no "good" direction, just signed
        13: False, 14: True,   # Current form For/Against
    }
    for col_idx, invert in zero_centered_cols.items():
        col_letter = get_column_letter(col_idx)
        color_scale_zero(ws, f"{col_letter}2:{col_letter}{last_row}", invert=invert)

    ws.freeze_panes = "A2"
    autosize_columns(ws, len(headers), width_hints=width_hints)
    ws.column_dimensions["A"].width = max(ws.column_dimensions["A"].width, 14)  # Team
    polish_sheet(ws, header_row=1, first_data_row=2, last_data_row=last_row, n_cols=len(headers))


def write_home_away_split_sheet(wb: Workbook, team_info: dict):
    """League-wide Total/Home/Away breakdown - supremacy first, goals
    second - one row per team, both tables in the same Total-avg-supremacy
    order so the two are easy to cross-reference."""
    ws = wb.create_sheet("Home & Away Splits")
    row_order = sorted(team_info.keys(),
                        key=lambda t: team_info[t]["stats"]["avg_supremacy"], reverse=True)

    ws.cell(row=1, column=1, value="Supremacy Delta - Total / Home / Away").font = BOLD_FONT
    ws.cell(row=1, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=7)
    sup_headers = ["Team", "Total Avg", "Total Median", "Home Avg", "Home Median", "Away Avg", "Away Median"]
    for c, h in enumerate(sup_headers, start=1):
        ws.cell(row=2, column=c, value=h)
    style_header_row(ws, 2, len(sup_headers))

    sup_width_hints = {i + 2: [] for i in range(6)}
    r = 2
    for team in row_order:
        s = team_info[team]["stats"]
        r += 1
        vals = [team, s["avg_supremacy"], s["median_supremacy"],
                s["home_avg_supremacy"], s["home_median_supremacy"],
                s["away_avg_supremacy"], s["away_median_supremacy"]]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = BODY_FONT
            if c > 1:
                cell.number_format = "0.00"
        for i in range(6):
            sup_width_hints[i + 2].append(vals[i + 1])
    sup_last_row = r
    for col_idx in (2, 3, 4, 5, 6, 7):
        letter = get_column_letter(col_idx)
        color_scale_zero(ws, f"{letter}3:{letter}{sup_last_row}", invert=False)
    autosize_columns(ws, len(sup_headers), header_row=2, max_row=sup_last_row, width_hints=sup_width_hints)
    polish_sheet(ws, header_row=2, first_data_row=3, last_data_row=sup_last_row, n_cols=len(sup_headers))

    goals_header_row = sup_last_row + 2
    ws.cell(row=goals_header_row, column=1, value="Total Goals - Total / Home / Away").font = BOLD_FONT
    ws.cell(row=goals_header_row, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=goals_header_row, start_column=1, end_row=goals_header_row, end_column=7)
    goals_headers = ["Team", "Total GF", "Total GA", "Home GF", "Home GA", "Away GF", "Away GA"]
    for c, h in enumerate(goals_headers, start=1):
        ws.cell(row=goals_header_row + 1, column=c, value=h)
    style_header_row(ws, goals_header_row + 1, len(goals_headers))

    goals_width_hints = {i + 2: [] for i in range(6)}
    r = goals_header_row + 1
    for team in row_order:
        s = team_info[team]["stats"]
        r += 1
        vals = [team, s["gf"], s["ga"], s["home_gf"], s["home_ga"], s["away_gf"], s["away_ga"]]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = BODY_FONT
        for i in range(6):
            goals_width_hints[i + 2].append(vals[i + 1])
    goals_last_row = r
    autosize_columns(ws, len(goals_headers), header_row=goals_header_row + 1, max_row=goals_last_row,
                      width_hints=goals_width_hints)
    polish_sheet(ws, header_row=goals_header_row + 1, first_data_row=goals_header_row + 2,
                 last_data_row=goals_last_row, n_cols=len(goals_headers))
    ws.column_dimensions["A"].width = max(ws.column_dimensions["A"].width, 20)
    ws.freeze_panes = "A3"


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
    polish_sheet(ws, header_row=1, first_data_row=2, last_data_row=last_row, n_cols=len(headers))
    return ws


def write_matrix_sheet(wb: Workbook, sheet_name: str, team_info: dict,
                        values_by_team: dict, sort_key_by_team: dict,
                        source_col_header: str, avg_summary_label: str,
                        colored: bool, invert: bool, col_prefix: str = "MW",
                        formula_row_lookup: bool = True):
    """Teams down the left, one column per match (col_prefix + "1", "2",
    ...), plus an Average column. Row order: DESCENDING by
    sort_key_by_team[team] (not necessarily the same stat as
    values_by_team - e.g. the Supremacy Delta grid is sorted by Avg xG
    Delta (For), see the About This Report sheet for why).

    When formula_row_lookup is True (the full-season grids), every match
    cell is a formula referencing that team's own sheet at a fixed row
    (source_col_header column, MW-th data row) - correct because the
    full-season grid's column N always means "this team's Nth match
    overall", matching the team sheet's own row order. The Home-only /
    Away-only grids can't use that trick (a team's home matches are
    scattered among its sheet's rows, not contiguous), so
    formula_row_lookup=False writes the values_by_team numbers directly
    instead - a Python-computed snapshot, not a live formula, same
    trade-off as the Home/Away split table on each team sheet."""
    max_matches = max((len(v) for v in values_by_team.values()), default=0)
    sorted_teams = sorted(values_by_team.keys(), key=lambda t: sort_key_by_team[t], reverse=True)

    ws = wb.create_sheet(sheet_name)
    headers = ["Team", "Average"] + [f"{col_prefix}{i}" for i in range(1, max_matches + 1)]
    ws.append(headers)
    style_header_row(ws, 1, len(headers))
    ws.cell(row=1, column=1).comment = Comment(MW_NOTE, "build_team_workbook.py")

    avg_by_team = {t: round(sum(v) / len(v), 2) if v else 0.0 for t, v in values_by_team.items()}
    width_hints = {2: [avg_by_team[t] for t in sorted_teams]}
    for mw in range(1, max_matches + 1):
        width_hints[mw + 2] = [v[mw - 1] for v in values_by_team.values() if len(v) >= mw]

    for row_i, team in enumerate(sorted_teams, start=2):
        info = team_info[team]
        n = len(values_by_team[team])

        ws.cell(row=row_i, column=1, value=team).font = BODY_FONT
        if formula_row_lookup:
            sheet_ref = info["sheet_name"]
            avg_row = info["summary_rows"][avg_summary_label]
            avg_cell = ws.cell(row=row_i, column=2, value=f"='{sheet_ref}'!$E${avg_row}")
        else:
            avg_cell = ws.cell(row=row_i, column=2, value=avg_by_team[team])
        avg_cell.font = BOLD_FONT
        avg_cell.number_format = "0.00"
        avg_cell.fill = SUMMARY_FILL

        for mw in range(1, max_matches + 1):
            col = mw + 2
            cell = ws.cell(row=row_i, column=col)
            if mw <= n:
                if formula_row_lookup:
                    sheet_ref = info["sheet_name"]
                    metric_letter = info["col_letter"][source_col_header]
                    data_row = mw + 1  # row 2 on the team sheet = MW1
                    cell.value = f"='{sheet_ref}'!${metric_letter}${data_row}"
                else:
                    cell.value = values_by_team[team][mw - 1]
                cell.number_format = "0.00"
            cell.font = BODY_FONT

    last_row = ws.max_row
    last_col_letter = get_column_letter(len(headers))
    if colored and last_row >= 2:
        color_scale_zero(ws, f"B2:{last_col_letter}{last_row}", invert=invert)

    ws.freeze_panes = "C2"  # keep Team + Average visible while scrolling across matchweeks
    autosize_columns(ws, len(headers), width_hints=width_hints)
    ws.column_dimensions["A"].width = max(ws.column_dimensions["A"].width, 20)  # Team
    polish_sheet(ws, header_row=1, first_data_row=2, last_data_row=last_row, n_cols=len(headers), band=False)
    return ws


def write_fixtures_sheet(wb: Workbook, fixtures_df: pd.DataFrame, team_info: dict):
    """Upcoming fixtures ranked by an "expected supremacy" for the home
    side that blends the market's own pre-match view of THIS fixture with
    how much each team has been beating/missing its own expectation all
    season:

        Combined Expected Supremacy (Home)
          = Market Pre-Match Supremacy (Home)
            + Home Team's Season Avg Supremacy Delta
            - Away Team's Season Avg Supremacy Delta

    where Market Pre-Match Supremacy (Home) = xg_pre_market_home -
    xg_pre_market_away, i.e. what the odds alone say about this specific
    match. A team with no season data yet (e.g. just promoted, or not yet
    covered by team_match_deltas) contributes 0 to its side of the
    adjustment and is flagged in the Notes column, rather than the whole
    row being dropped."""
    ws = wb.create_sheet("Upcoming Fixtures")
    headers = ["Date", "Home Team", "Away Team", "Market Pre-Match Supremacy (Home)",
               "Home Team Season Avg Supremacy", "Away Team Season Avg Supremacy",
               "Combined Expected Supremacy (Home)", "Notes"]
    ws.append(headers)
    style_header_row(ws, 1, len(headers))

    rows = []
    for _, r in fixtures_df.iterrows():
        home_norm = normalise_name(r["home_team"], TEAM_NAME_MAP)
        away_norm = normalise_name(r["away_team"], TEAM_NAME_MAP)
        market_supremacy = round(r["xg_pre_market_home"] - r["xg_pre_market_away"], 2)

        notes = []
        home_stats = team_info.get(home_norm, {}).get("stats")
        away_stats = team_info.get(away_norm, {}).get("stats")
        home_avg = home_stats["avg_supremacy"] if home_stats else 0.0
        away_avg = away_stats["avg_supremacy"] if away_stats else 0.0
        if not home_stats:
            notes.append(f"No season data yet for {home_norm}")
        if not away_stats:
            notes.append(f"No season data yet for {away_norm}")
        combined = round(market_supremacy + home_avg - away_avg, 2)

        date_val = r.get("date")
        try:
            date_str = pd.to_datetime(date_val).strftime("%Y-%m-%d")
        except (ValueError, TypeError):
            date_str = str(date_val)

        rows.append((date_str, home_norm, away_norm, market_supremacy, home_avg, away_avg,
                     combined, "; ".join(notes)))

    rows.sort(key=lambda row: row[6], reverse=True)

    width_hints = {i + 1: [] for i in range(len(headers))}
    for row_vals in rows:
        ws.append(list(row_vals))
        for i, v in enumerate(row_vals, start=1):
            width_hints[i].append(v)

    last_row = ws.max_row
    decimal_cols = {4, 5, 6, 7}
    for row in ws.iter_rows(min_row=2, max_row=last_row, max_col=len(headers)):
        for cell in row:
            cell.font = BODY_FONT
            if cell.column in decimal_cols:
                cell.number_format = "0.00"

    if last_row >= 2:
        color_scale_zero(ws, f"G2:G{last_row}", invert=False)

    ws.freeze_panes = "A2"
    autosize_columns(ws, len(headers), width_hints=width_hints, max_width=40)
    polish_sheet(ws, header_row=1, first_data_row=2, last_data_row=last_row, n_cols=len(headers))
    return ws


def write_about_sheet(wb: Workbook, league_note: str | None = None):
    ws = wb.create_sheet("About This Report")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 100

    ws.cell(row=1, column=1, value="About This Report").font = TITLE_FONT
    r = 3
    paragraphs = [
        ("Sheets in this workbook", BOLD_FONT),
        ("League Table - standard standings from every played match.", BODY_FONT),
        ("Overview - every team's season xG form, sorted by Avg xG Delta (For), best first.", BODY_FONT),
        ("Home & Away Splits - Total/Home/Away average and median supremacy, and Total/Home/Away goals, for every team.", BODY_FONT),
        ("Supremacy Delta / (Home) / (Away) - match-by-match grids of actual vs. expected xG supremacy.", BODY_FONT),
        ("Total xG Delta - match-by-match grid of how much higher/lower-scoring each match was than the market expected.", BODY_FONT),
        ("Upcoming Fixtures - only present if fixture data was supplied; see below for how it's ranked.", BODY_FONT),
        ("One tab per team - full match log, season summary (average AND median of every delta stat), and a Home/Away split table.", BODY_FONT),
        ("", BODY_FONT),
        ("Judgement calls made where the brief was ambiguous", BOLD_FONT),
        ("\u2022 \"In order of xG for\": Overview and the Supremacy Delta grid are both sorted by Avg xG Delta (For), "
         "highest (most attacking overperformance) first - not by average Supremacy Delta, which was the grid's "
         "original sort key.", BODY_FONT),
        ("\u2022 \"Total xG data sheet ... order by average\": Total xG Delta is sorted by its own average, "
         "highest (most above the market's pre-match total-goals expectation) first.", BODY_FONT),
        ("\u2022 Conditional formatting on every delta-style column is now zero-anchored (green above 0, red below 0, "
         "white at exactly 0) rather than relative to that column's own min/max/median - 0 has a real meaning here "
         "(\"exactly as the market expected\"), so it's a fixed reference point rather than a moving one.", BODY_FONT),
        ("\u2022 Upcoming Fixtures' \"Combined Expected Supremacy (Home)\" = the market's own pre-match supremacy for "
         "that specific fixture, adjusted by how much each side has been over/under-performing its own expectation "
         "all season (home team's season average added, away team's season average subtracted). This is a simple, "
         "transparent blend - not a fitted model - so treat it as a way to surface interesting fixtures, not a "
         "prediction.", BODY_FONT),
        ("\u2022 The Supremacy Delta (Home) / (Away) grids are literal snapshots (computed once, at report-generation "
         "time) rather than live formulas like the full-season grid, since a team's home-only matches aren't a "
         "contiguous block of rows on its own sheet. Re-run the pipeline to refresh them.", BODY_FONT),
    ]
    for text, font in paragraphs:
        cell = ws.cell(row=r, column=1, value=text)
        cell.font = font
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 30 if text else 8
        r += 1

    if league_note:
        r += 1
        cell = ws.cell(row=r, column=1, value=league_note)
        cell.font = SUBTITLE_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 45


def main(input_path, results_path, output_path, rolling_window, fixtures_path=None, league_note=None):
    df = pd.read_csv(input_path, parse_dates=["date"])
    teams = sorted(df["team"].unique())

    wb = Workbook()
    wb.remove(wb.active)  # drop the default blank sheet

    write_about_sheet(wb, league_note=league_note)

    if results_path:
        results_df = pd.read_csv(results_path)
        results_df = results_df.dropna(subset=["FTHG", "FTAG"])  # only played matches
        last_results_row = write_results_data_sheet(wb, results_df)
        write_league_table_sheet(wb, results_df, last_results_row)

    used_names = set()
    team_info = {}
    supremacy_by_team, total_delta_by_team = {}, {}
    home_supremacy_by_team, away_supremacy_by_team = {}, {}
    for team in teams:
        sheet_name = safe_sheet_name(team, used_names)
        team_df = df[df["team"] == team].sort_values("date")
        info = write_team_sheet(wb, sheet_name, team_df, rolling_window)
        team_info[team] = info

        supremacy_by_team[team] = info["supremacy_series"]
        total_delta_by_team[team] = info["total_series"]
        home_supremacy_by_team[team] = info["home_supremacy_series"]
        away_supremacy_by_team[team] = info["away_supremacy_series"]

    write_overview_sheet(wb, team_info, rolling_window)
    write_home_away_split_sheet(wb, team_info)

    avg_for_by_team = {t: team_info[t]["stats"]["avg_for"] for t in teams}
    avg_total_by_team = {t: team_info[t]["stats"]["avg_total_delta"] for t in teams}
    write_matrix_sheet(wb, "Supremacy Delta", team_info, supremacy_by_team, sort_key_by_team=avg_for_by_team,
                        source_col_header="Supremacy Delta", avg_summary_label="Avg Supremacy Delta",
                        colored=True, invert=False)
    write_matrix_sheet(wb, "Supremacy Delta (Home)", team_info, home_supremacy_by_team,
                        sort_key_by_team=avg_for_by_team, source_col_header="Supremacy Delta",
                        avg_summary_label="Avg Supremacy Delta", colored=True, invert=False,
                        col_prefix="H", formula_row_lookup=False)
    write_matrix_sheet(wb, "Supremacy Delta (Away)", team_info, away_supremacy_by_team,
                        sort_key_by_team=avg_for_by_team, source_col_header="Supremacy Delta",
                        avg_summary_label="Avg Supremacy Delta", colored=True, invert=False,
                        col_prefix="A", formula_row_lookup=False)
    write_matrix_sheet(wb, "Total xG Delta", team_info, total_delta_by_team, sort_key_by_team=avg_total_by_team,
                        source_col_header="Match Total xG Delta", avg_summary_label="Avg Total xG Delta",
                        colored=True, invert=False)

    fixtures_sheet_written = False
    if fixtures_path:
        fixtures_df = pd.read_csv(fixtures_path)
        if not fixtures_df.empty:
            write_fixtures_sheet(wb, fixtures_df, team_info)
            fixtures_sheet_written = True
        else:
            print(f"'{fixtures_path}' has no rows - skipping the Upcoming Fixtures sheet.")

    # Sheet order: About, League Table, Overview, Home & Away Splits, the
    # grids, Upcoming Fixtures, then teams A-Z (Results Data stays hidden
    # but present).
    front = [n for n in ("About This Report", "League Table", "Overview", "Home & Away Splits",
                          "Supremacy Delta", "Supremacy Delta (Home)", "Supremacy Delta (Away)",
                          "Total xG Delta", "Upcoming Fixtures") if n in wb.sheetnames]
    order = front + [n for n in wb.sheetnames if n not in front and n != "Results Data"]
    if "Results Data" in wb.sheetnames:
        order.append("Results Data")
    wb._sheets = [wb[n] for n in order]

    wb.save(output_path)
    extra = " + Upcoming Fixtures" if fixtures_sheet_written else ""
    print(f"Wrote workbook with {len(teams)} team sheets + League Table + Overview + "
          f"Home & Away Splits + Supremacy Delta (+ Home/Away) + Total xG Delta{extra} to {output_path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", default="data/team_match_deltas.csv")
    parser.add_argument("--results", default="data/E1.csv",
                         help="Raw results file (for the League Table). Pass '' to skip the League Table.")
    parser.add_argument("--fixtures", default=None,
                         help="Pre-match xG file for upcoming fixtures (from download_fixtures.py + "
                              "derive_prematch_xg.py). Omit or pass '' to skip the Upcoming Fixtures sheet.")
    parser.add_argument("--output", default="data/championship_xg_report.xlsx")
    parser.add_argument("--rolling-window", type=int, default=6,
                         help="Only used to label the 'current form' column - must match what you passed to build_deltas.py")
    args = parser.parse_args()

    main(args.input, args.results or None, args.output, args.rolling_window, fixtures_path=args.fixtures or None)
