"""
download_fixtures.py

Purpose
-------
Download the *upcoming* fixtures (not yet played) for a given league,
odds and all, so the pipeline can produce a pre-match expected-goals
number for matches that haven't happened yet - which is what the
"Upcoming Fixtures" sheet in the Excel report needs.

football-data.co.uk's per-season files (E1.csv, SC0.csv, etc. - what
download_results.py fetches) only ever contain matches that have
already been played. Upcoming fixtures with their opening odds live in
a separate, single, ALL-LEAGUES-COMBINED file:

    https://www.football-data.co.uk/fixtures.csv

This script downloads that one file and filters it down to just the
rows for the league you asked for (matching on the "Div" column against
that league's fd_code from leagues.py), so the rest of the pipeline
never has to know the combined file exists.

Usage
-----
    python download_fixtures.py --league championship
    python download_fixtures.py --league premier-league --output data/fixtures_premier-league.csv

The output has the same column shape as the historical results files
(Date, HomeTeam, AwayTeam, B365H/D/A, Over/Under 2.5 odds where
available, ...) minus FTHG/FTAG/FTR, since those matches haven't been
played yet. That means derive_prematch_xg.py works on it completely
unchanged - it never reads FTHG/FTAG anyway.

Notes
-----
- This file only covers the NEXT round or two of fixtures with odds
  already posted (typically odds go up a few days before kickoff, per
  football-data.co.uk's own notes - "collected Friday afternoons ...
  and Tuesday afternoons for midweek games"), so don't expect a full
  season of future fixtures here - just what's coming up soon. Re-run
  this weekly (run_weekly_update.py does) to keep the Upcoming Fixtures
  sheet current.
- If a league currently has no upcoming fixtures with posted odds (e.g.
  between rounds, or the file hasn't been refreshed yet), this writes
  an empty (header-only) output file rather than failing - the
  Upcoming Fixtures sheet then just shows nothing for that league
  rather than erroring.
"""

import argparse
from io import StringIO
from pathlib import Path

import pandas as pd
import requests

from leagues import get_league, league_choices

FIXTURES_URL = "https://www.football-data.co.uk/fixtures.csv"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )
}


def download_and_filter(fd_code: str) -> pd.DataFrame:
    print(f"Downloading {FIXTURES_URL} ...")
    resp = requests.get(FIXTURES_URL, headers=HEADERS, timeout=30)
    resp.raise_for_status()

    # This combined file covers every league football-data.co.uk tracks,
    # so a byte-count sanity check (like download_results.py's) isn't
    # useful here - instead just check it actually parses and has a Div
    # column before trusting it.
    df = pd.read_csv(StringIO(resp.content.decode("latin-1")))
    if "Div" not in df.columns:
        raise ValueError(
            f"Downloaded fixtures.csv has no 'Div' column - football-data.co.uk "
            f"may have changed this file's format. Columns found: {list(df.columns)}"
        )

    filtered = df[df["Div"] == fd_code].copy()
    return filtered


def main(league_key: str, output_path: str):
    league = get_league(league_key)
    filtered = download_and_filter(league["fd_code"])

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    filtered.to_csv(out, index=False)

    if filtered.empty:
        print(f"No upcoming {league['label']} fixtures with posted odds right now "
              f"- wrote an empty file to {out}. This is normal between rounds; "
              f"re-run closer to the next fixtures.")
    else:
        print(f"Wrote {len(filtered)} upcoming {league['label']} fixtures to {out}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--league", default="championship", choices=league_choices())
    parser.add_argument("--output", default=None,
                         help="Defaults to data/fixtures_<league>.csv")
    args = parser.parse_args()

    output = args.output or f"data/fixtures_{args.league}.csv"
    main(args.league, output)
