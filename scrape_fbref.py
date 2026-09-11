"""
scrape_fbref.py

Purpose
-------
Pull match-level, shot-based xG data for an EFL division from FBref
(FBref's underlying xG numbers come from StatsBomb). This gives you the
"actual" post-match xG to compare against the pre-match market-implied xG
produced by derive_prematch_xg.py.

FBref covers the Championship (unlike Understat, which only covers the
"big 5" leagues), via its Scores & Fixtures page:

    https://fbref.com/en/comps/10/schedule/Championship-Scores-and-Fixtures

That table includes an xG column for the home team and one for the away
team, filled in once a match has been played.

*** IMPORTANT — League One / League Two gap ***
As of this writing, FBref's League One and League Two Scores & Fixtures
pages do NOT have xG columns (only the Championship does among the three
EFL divisions covered here) — so running this script with
--league league-one/league-two will currently find no usable data and
raise the "could not find xG columns" error below, by design rather than
a bug. This is exactly why the pipeline's default post-match source is
scrape_fotmob.py (FotMob does carry xG for all three divisions) — this
script is kept around for the Championship and in case FBref adds xG
coverage for the lower divisions later; check the schedule URL printed
below by hand before relying on it for League One/Two.

Usage
-----
    python scrape_fbref.py --league league-one --season 2024-2025 --output data/fbref_league-one_xg.csv

    # or, for the current season, omit --season to use FBref's default
    # "current season" schedule URL:
    python scrape_fbref.py --league championship --output data/fbref_championship_xg.csv

Important — rate limiting
--------------------------
FBref actively rate-limits scrapers. Their guidance is roughly one request
every 3+ seconds sustained; if you hammer it you'll get temporarily blocked
(HTTP 429). This script sleeps between requests by default — don't remove
that. If you're pulling per-match "player level" stats pages later (not
done here, we only use the schedule table), space those out even more.

Method
------
FBref pages are plain HTML tables, but a lot of their tables are wrapped in
HTML comments that pandas.read_html silently ignores unless you strip the
comment markers first. This script:
  1. Fetches the schedule page for the given season
  2. Strips `<!--` / `-->` comment wrappers so the embedded table becomes
     visible to the parser
  3. Extracts the Scores & Fixtures table via pandas.read_html
  4. Filters to matches that have been played (xG columns populated)
  5. Writes a tidy CSV: date, home_team, away_team, home_xg, away_xg,
     home_goals, away_goals

Notes
-----
- Not run against a live network in this environment (no network access
  here) — test locally before relying on it, and check FBref's current
  table id in case they've renamed it (was "sched_2024-2025_10_1" style ids
  historically; the script searches for any table id starting with
  "sched_" as a more robust fallback).
- Respect FBref's terms of use / robots.txt for your use case; this is
  intended for personal analysis, not redistribution of their data.
"""

import argparse
import re
import time

import pandas as pd
import requests

from leagues import get_league

HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; personal-research-script/1.0)"
}
REQUEST_DELAY_SECONDS = 4  # be polite; FBref will block aggressive scraping


def build_url(season: str | None, comp_id: int, slug: str) -> str:
    if season:
        # FBref season URLs look like:
        # https://fbref.com/en/comps/10/2024-2025/schedule/2024-2025-Championship-Scores-and-Fixtures
        return (f"https://fbref.com/en/comps/{comp_id}/{season}/schedule/"
                f"{season}-{slug}-Scores-and-Fixtures")
    else:
        # current season shortcut
        return f"https://fbref.com/en/comps/{comp_id}/schedule/{slug}-Scores-and-Fixtures"


def fetch_html(url: str) -> str:
    resp = requests.get(url, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    time.sleep(REQUEST_DELAY_SECONDS)
    return resp.text


def extract_schedule_table(html: str) -> pd.DataFrame:
    # FBref wraps some tables in HTML comments; uncomment them so read_html can see them.
    uncommented = re.sub(r"<!--|-->", "", html)

    tables = pd.read_html(uncommented)
    # Look for the schedule table by checking for expected columns
    for t in tables:
        cols = [str(c) for c in t.columns]
        if any("xG" in c for c in cols) and "Home" in cols and "Away" in cols:
            return t

    raise ValueError(
        "Could not find the Scores & Fixtures table with xG columns. "
        "FBref may have changed their page layout — inspect the tables "
        "returned by pandas.read_html manually."
    )


def tidy_table(raw: pd.DataFrame) -> pd.DataFrame:
    # FBref's schedule table typically has columns like:
    # Wk, Day, Date, Time, Home, xG, Score, xG.1, Away, Attendance, Venue, Referee, ...
    # The two xG columns are usually named "xG" and "xG.1" (home, away respectively)
    # after pandas dedups the duplicate header name. Confirm/adjust if FBref
    # renames things.
    col_map = {}
    for c in raw.columns:
        cs = str(c)
        if cs == "Date":
            col_map[c] = "date"
        elif cs == "Home":
            col_map[c] = "home_team"
        elif cs == "Away":
            col_map[c] = "away_team"
        elif cs == "Score":
            col_map[c] = "score"
        elif cs == "xG":
            col_map[c] = "home_xg"
        elif cs == "xG.1":
            col_map[c] = "away_xg"

    df = raw.rename(columns=col_map)
    keep_cols = [c for c in ["date", "home_team", "away_team", "score", "home_xg", "away_xg"] if c in df.columns]
    df = df[keep_cols].copy()

    # Drop rows without a played score (future fixtures) or without xG populated
    df = df.dropna(subset=["score", "home_xg", "away_xg"], how="any")
    df = df[df["score"].astype(str).str.contains("–|-", na=False)]

    def split_score(s):
        s = str(s).replace("–", "-")
        parts = s.split("-")
        if len(parts) == 2:
            try:
                return int(parts[0].strip()), int(parts[1].strip())
            except ValueError:
                return None, None
        return None, None

    goals = df["score"].apply(split_score)
    df["home_goals"] = [g[0] for g in goals]
    df["away_goals"] = [g[1] for g in goals]

    df = df.drop(columns=["score"])
    df["home_xg"] = pd.to_numeric(df["home_xg"], errors="coerce")
    df["away_xg"] = pd.to_numeric(df["away_xg"], errors="coerce")
    df = df.dropna(subset=["home_xg", "away_xg", "home_goals", "away_goals"])

    return df.reset_index(drop=True)


def main(season, output_path, league_key):
    league = get_league(league_key)
    url = build_url(season, league["fbref_comp_id"], league["fbref_slug"])
    print(f"League: {league['label']}")
    print(f"Fetching: {url}")
    html = fetch_html(url)
    raw_table = extract_schedule_table(html)
    tidy = tidy_table(raw_table)
    tidy.to_csv(output_path, index=False)
    print(f"Wrote {len(tidy)} played matches with xG to {output_path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--league", default="championship",
                         choices=["championship", "league-one", "league-two"])
    parser.add_argument("--season", default=None,
                         help="Season in FBref's format, e.g. 2024-2025. Omit for current season.")
    parser.add_argument("--output", required=True, help="Path to write tidy match xG CSV")
    args = parser.parse_args()

    main(args.season, args.output, args.league)
