"""
run_weekly_update.py

Purpose
-------
Run the whole weekly refresh, for one EFL division, in one command:

    1. download_results.py   -> data/<fd_code>.csv              (odds + results)
    2. derive_prematch_xg.py -> data/prematch_xg_<league>.csv    (market-implied xG)
    3. scrape_fotmob.py      -> data/fotmob_<league>_xg.csv      (shot-based xG, incremental)
    4. build_deltas.py       -> data/team_match_deltas_<league>.csv
    5. build_team_workbook.py -> data/<league>_xg_report.xlsx    (one tab per team)

Pick the division with --league (championship / league-one / league-two,
see leagues.py) — everything downstream (filenames, football-data.co.uk
code, FotMob league id) follows from that one flag:

    python run_weekly_update.py --league league-one
    python run_weekly_update.py --league league-two
    python run_weekly_update.py                        # defaults to championship

This is what the scheduled GitHub Actions workflows in .github/workflows/
run (one workflow file per league, each calling this with its own
--league) — running it by hand is exactly the same command.

Each step is a subprocess call to the existing script, so this file
doesn't duplicate any pipeline logic — it just sequences it and stops
with a clear error if a step fails, rather than silently producing a
stale/partial report.
"""

import argparse
import subprocess
import sys
from pathlib import Path

from leagues import get_league

ROLLING_WINDOW = 6


def build_steps(league_key: str, league: dict) -> list:
    fotmob_xg_path = f"data/fotmob_{league_key}_xg.csv"
    results_path = f"data/{league['fd_code']}.csv"
    prematch_path = f"data/prematch_xg_{league_key}.csv"
    deltas_path = f"data/team_match_deltas_{league_key}.csv"
    report_path = f"data/{league_key}_xg_report.xlsx"

    return [
        ("Downloading latest results/odds",
         [sys.executable, "download_results.py",
          "--league", league_key, "--output", results_path]),
        ("Deriving pre-match market-implied xG",
         [sys.executable, "derive_prematch_xg.py",
          "--league", league_key,
          "--input", results_path, "--output", prematch_path]),
        ("Scraping post-match shot-based xG from FotMob (incremental)",
         [sys.executable, "scrape_fotmob.py",
          "--league", league_key,
          "--output", fotmob_xg_path, "--since", "AUTO"]),
        ("Building per-team-per-match form deltas",
         [sys.executable, "build_deltas.py",
          "--prematch", prematch_path,
          "--postmatch", fotmob_xg_path,
          "--output", deltas_path,
          "--rolling-window", str(ROLLING_WINDOW)]),
        ("Building the Excel workbook (one tab per team, league table, overview)",
         [sys.executable, "build_team_workbook.py",
          "--input", deltas_path,
          "--results", results_path,
          "--output", report_path,
          "--rolling-window", str(ROLLING_WINDOW)]),
    ], fotmob_xg_path, report_path


def resolve_since_date(fotmob_xg_path: str) -> str:
    """Pick a --since date for the FotMob scrape: a week before the latest
    date already in the league's fotmob xG file, so a rerun re-checks the
    most recent gameweek (in case xG wasn't posted yet last time) without
    re-scraping the whole season. Falls back to fetching everything on the
    very first run."""
    import pandas as pd

    path = Path(fotmob_xg_path)
    if not path.exists():
        return None
    df = pd.read_csv(path)
    if df.empty or "date" not in df.columns:
        return None
    latest = pd.to_datetime(df["date"]).max()
    since = (latest - pd.Timedelta(days=7)).date().isoformat()
    return since


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--league", default="championship",
                         choices=["championship", "league-one", "league-two"])
    args = parser.parse_args()

    league = get_league(args.league)
    steps, fotmob_xg_path, report_path = build_steps(args.league, league)
    since = resolve_since_date(fotmob_xg_path)

    print(f"=== League: {league['label']} ===")

    for label, cmd in steps:
        if "AUTO" in cmd:
            idx = cmd.index("AUTO")
            if since:
                cmd[idx] = since
            else:
                # No prior data: drop --since entirely and fetch the full season.
                del cmd[idx - 1:idx + 1]

        print(f"\n=== {label} ===")
        print(" ".join(cmd))
        result = subprocess.run(cmd)
        if result.returncode != 0:
            print(f"\nStep failed: {label} (exit code {result.returncode}). Stopping.")
            sys.exit(result.returncode)

    print(f"\nWeekly update complete. Report: {report_path}")


if __name__ == "__main__":
    main()
