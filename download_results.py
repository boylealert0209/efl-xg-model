"""
download_results.py

Purpose
-------
Download the current-season results/odds file for a given EFL division
from football-data.co.uk (this is the "E1"/"E2"/"E3.csv" the rest of the
pipeline runs on) without you having to visit the site and save it by
hand each week.

football-data.co.uk publishes one cumulative CSV per league per season at
a predictable URL:

    https://www.football-data.co.uk/mmz4281/<season>/<code>.csv

where <season> is like "2627" for the 2026-27 season and <code> depends
on the division: E1 = Championship, E2 = League One, E3 = League Two
(see leagues.py). The file for the current season is updated in place as
new results come in — every fixture played so far, not just the latest
gameweek — so re-downloading it each week and overwriting the same file
is enough; there is no need to fetch old seasons separately.

Usage
-----
    python download_results.py --league league-one
    python download_results.py --league league-two --season 2627   # override auto-detected season
    python download_results.py --league championship --output data/E1.csv  # default shown

This is meant to be step 1 of the weekly refresh — see run_weekly_update.py
to run the whole pipeline (this download, pre-match xG, FotMob scrape,
deltas, and the Excel workbook) in one go.
"""

import argparse
from datetime import date, datetime
from pathlib import Path

import requests

from leagues import get_league

BASE_URL = "https://www.football-data.co.uk/mmz4281"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )
}


def current_season_code(today: date | None = None) -> str:
    """Work out the football-data.co.uk season code (e.g. '2627' for
    2026-27) from today's date.

    The Championship season runs August through May, so:
      - Jul-Dec  -> season starts this calendar year
      - Jan-Jun  -> season started last calendar year
    """
    today = today or date.today()
    start_year = today.year if today.month >= 7 else today.year - 1
    end_year = start_year + 1
    return f"{str(start_year)[2:]}{str(end_year)[2:]}"


def download(season: str, league_code: str, output_path: str) -> None:
    url = f"{BASE_URL}/{season}/{league_code}.csv"
    print(f"Downloading {url} ...")
    resp = requests.get(url, headers=HEADERS, timeout=30)
    resp.raise_for_status()

    if len(resp.content) < 1000:
        # A too-small response is almost always a redirect/error page, not
        # a real season file (e.g. season hasn't started, or the code
        # changed) — fail loudly instead of silently overwriting good data
        # with junk.
        raise ValueError(
            f"Downloaded file from {url} looks too small ({len(resp.content)} "
            f"bytes) to be a real results file. Check the season code and "
            f"that the page still exists in a browser before retrying."
        )

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)

    # Keep a dated backup of whatever was there before, so a bad download
    # never destroys the last known-good file without a way back.
    if out.exists():
        backup = out.with_name(f"{out.stem}_backup_{datetime.now():%Y%m%d_%H%M%S}{out.suffix}")
        out.replace(backup)
        print(f"Backed up previous file to {backup}")

    out.write_bytes(resp.content)
    print(f"Wrote {len(resp.content):,} bytes to {out}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--league", default="championship",
                         choices=["championship", "league-one", "league-two"])
    parser.add_argument("--season", default=None,
                         help="Season code, e.g. 2627. Defaults to auto-detecting from today's date.")
    parser.add_argument("--output", default=None,
                         help="Defaults to data/<football-data.co.uk code>.csv for the chosen league.")
    args = parser.parse_args()

    league = get_league(args.league)
    season = args.season or current_season_code()
    output = args.output or f"data/{league['fd_code']}.csv"
    print(f"League: {league['label']} ({league['fd_code']})")
    print(f"Using season code: {season}")
    download(season, league["fd_code"], output)
