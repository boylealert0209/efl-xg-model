"""
build_deltas.py

Purpose
-------
Merge the pre-match market-implied xG (from derive_prematch_xg.py) with the
post-match shot-based xG (from scrape_fbref_championship.py) and compute the
per-team, per-match deltas that drive the "form vs expectation" tracking
described in the brief:

    xg_delta_for      = actual match xG (shot-based) - pre-match expected (market)
    xg_delta_against  = opponent's actual match xG    - opponent's pre-match expected

A positive xg_delta_for means the team created more/better chances than the
market expected of them. A negative xg_delta_against means the team
conceded fewer/worse chances than the market expected them to concede
(i.e. defended better than expected).

Output is one row PER TEAM PER MATCH (not per fixture), which makes rolling
form averages trivial with a groupby + rolling window.

Usage
-----
    python build_deltas.py \
        --prematch data/prematch_xg.csv \
        --postmatch data/fotmob_league-one_xg.csv \
        --output data/team_match_deltas_league-one.csv \
        --rolling-window 6

(League-agnostic — works the same for whichever league's prematch/
postmatch CSVs you point it at; nothing here needs a --league flag.)

Matching caveat
----------------
Team name spellings often differ between football-data.co.uk and FBref
(e.g. "Nott'm Forest" vs "Nottingham Forest", "QPR" vs "Queens Park
Rangers"). This script does a straightforward date + name match and will
report any fixtures it couldn't pair up. Build out the TEAM_NAME_MAP dict
below as you hit mismatches — this is normal and expected on first run.
"""

import argparse
import pandas as pd

# Extend this as you discover mismatches between the two sources.
# Maps a variant name (however it appears in either file) -> the name to
# standardise on. Add one line per mismatch you spot.
#
# The League One/Two entries below are a starting point based on how
# football-data.co.uk's short club names typically differ from FotMob's
# fuller ones (e.g. "Doncaster" vs "Doncaster Rovers") — not a verified
# list for the current season's actual divisions, since club names carry
# over across seasons/divisions but which clubs are IN League One vs
# League Two changes every year with promotion/relegation. Run
# build_deltas.py once against real data and extend this from whatever it
# reports as unmatched; that's normal and expected on first run for any
# league, per the note above.
TEAM_NAME_MAP = {
"Bradford": "Bradford City", 
"Peterboro": "Peterborough United", 
 "Crawley": "Crawley Town", 
"York City": "York", 
"Bristol Rovers": 
"Bristol Rvs", 
"Newport": "Newport County",
"Sheff Wed": "Sheffield Wednesday", 
"Peterborough": "Peterborough United",
"Oxford": "Oxford United", 
"Oxford Utd": "Oxford United",
"Sheffield Weds": "Sheffield Wednesday", 
"Wigan": "Wigan Athletic", "Burton": 
"Burton Albion", "Leicester": 
"Leicester City", 
"Wimbledon": "AFC Wimbledon",
    "Nott'm Forest": "Nottingham Forest",
    "QPR": "Queens Park Rangers",
    "Sheff Utd": "Sheffield United",
    "West Brom": "West Bromwich Albion",
    "Preston": "Preston North End",
    # Common League One / League Two short-name -> full-name mismatches
    "Doncaster": "Doncaster Rovers",
    "Bolton": "Bolton Wanderers",
    "Wycombe": "Wycombe Wanderers",
    "Huddersfield": "Huddersfield Town",
    "Stockport": "Stockport County",
    "Peterborough": "Peterborough United",
    "Rotherham": "Rotherham United",
    "Northampton": "Northampton Town",
    "Crewe": "Crewe Alexandra",
    "Grimsby": "Grimsby Town",
    "Salford": "Salford City",
    "Shrewsbury": "Shrewsbury Town",
    "Colchester": "Colchester United",
    "Tranmere": "Tranmere Rovers",
    "Oldham": "Oldham Athletic",
    "Cheltenham": "Cheltenham Town",
    "Accrington": "Accrington Stanley",
    "MK Dons": "Milton Keynes Dons",
    "Fleetwood": "Fleetwood Town",
    "Swindon": "Swindon Town",
    # Premier League short-name -> full-name starter entries (same caveat
    # as above: not verified against the current season's actual PL
    # membership, just common football-data.co.uk vs FotMob spellings).
    "Man United": "Manchester United",
    "Man Utd": "Manchester United",
    "Man City": "Manchester City",
    "Newcastle": "Newcastle United",
    "Tottenham": "Tottenham Hotspur",
    "Wolves": "Wolverhampton Wanderers",
    "West Ham": "West Ham United",
    "Brighton": "Brighton & Hove Albion",
    "Leeds": "Leeds United",
    "Nott'm Forest": "Nottingham Forest",
    # Scottish Premiership starter entries.
    "Hearts": "Heart of Midlothian",
    "Dundee Utd": "Dundee United",
    "St Johnstone": "St Johnstone",
    # National League starter entries - the division has ~24 clubs whose
    # names change every season with promotion/relegation from a much
    # larger non-league pyramid, so this is deliberately just a few of
    # the more common short/long-name mismatches seen historically, not
    # a verified list for the current season - extend from whatever
    # build_deltas.py reports as unmatched, same as every other league.
    "FC Halifax": "FC Halifax Town",
    "Dag and Red": "Dagenham & Redbridge",
    "Forest Green": "Forest Green Rovers",
    "Boreham Wood": "Boreham Wood",
    "Ebbsfleet": "Ebbsfleet United",
    "Solihull": "Solihull Moors",
    "Woking": "Woking",
    "Yeovil": "Yeovil Town",
    "Altrincham": "Altrincham",
    "Aldershot": "Aldershot Town",
}


def normalise_name(name, mapping):
    return mapping.get(name, name)


def load_and_merge(prematch_path, postmatch_path):
    # Both inputs use unambiguous ISO (YYYY-MM-DD[THH:MM:SS]) dates by the
    # time they reach this script (derive_prematch_xg.py normalises the
    # odds file's dates; FotMob's are already ISO) so no dayfirst guessing
    # is needed here.
    pre = pd.read_csv(prematch_path, parse_dates=["date"])
    post = pd.read_csv(postmatch_path, parse_dates=["date"])

    # Normalise both date columns to plain (timezone-naive) calendar dates.
    # The odds file has plain dates; FotMob's dates carry UTC timezone info
    # (e.g. "...T14:00:00.000Z"), which pandas won't merge against a
    # timezone-naive column without this step first. tz_convert/tz_localize
    # only apply to tz-aware columns, so check before calling either.
    for df in (pre, post):
        if isinstance(df["date"].dtype, pd.DatetimeTZDtype):
            df["date"] = df["date"].dt.tz_localize(None)
        df["date"] = df["date"].dt.normalize()

    pre["home_team_norm"] = pre["home_team"].apply(lambda n: normalise_name(n, TEAM_NAME_MAP))
    pre["away_team_norm"] = pre["away_team"].apply(lambda n: normalise_name(n, TEAM_NAME_MAP))
    post["home_team_norm"] = post["home_team"].apply(lambda n: normalise_name(n, TEAM_NAME_MAP))
    post["away_team_norm"] = post["away_team"].apply(lambda n: normalise_name(n, TEAM_NAME_MAP))

    merged = pd.merge(
        pre, post,
        left_on=["date", "home_team_norm", "away_team_norm"],
        right_on=["date", "home_team_norm", "away_team_norm"],
        how="inner",
        suffixes=("_pre", "_post"),
    )

    unmatched_pre = pre.shape[0] - merged.shape[0]
    if unmatched_pre > 0:
        print(f"Warning: {unmatched_pre} pre-match rows did not find a same-date "
              f"team-name match in the post-match data. Check TEAM_NAME_MAP, "
              f"or date misalignment (e.g. postponed/rearranged fixtures).")
        # Show exactly which rows failed to match, so the mismatch is
        # visible instead of guessed at.
        matched_keys = set(
            zip(merged["date"], merged["home_team_norm"], merged["away_team_norm"])
        )
        pre_keys = set(
            zip(pre["date"], pre["home_team_norm"], pre["away_team_norm"])
        )
        missing = pre_keys - matched_keys
        print("Unmatched pre-match fixtures (date, home, away):")
        for date, home, away in sorted(missing, key=lambda x: str(x[0])):
            print(f"  {date.date()}  {home}  vs  {away}")
        print("Compare each of these against data/fotmob_championship_xg.csv "
              "for the same date to see the exact spelling FotMob used.")

    return merged


def to_long_form(merged, rolling_window):
    """Reshape from one-row-per-fixture to one-row-per-team-per-match, and
    compute the xg_delta_for / xg_delta_against columns plus rolling
    averages."""

    home_rows = pd.DataFrame({
        "date": merged["date"],
        "team": merged["home_team_norm"],
        "opponent": merged["away_team_norm"],
        "venue": "home",
        "xg_pre_market_for": merged["xg_pre_market_home"],
        "xg_pre_market_against": merged["xg_pre_market_away"],
        "xg_actual_for": merged["home_xg"],
        "xg_actual_against": merged["away_xg"],
        "goals_for": merged["home_goals"],
        "goals_against": merged["away_goals"],
    })

    away_rows = pd.DataFrame({
        "date": merged["date"],
        "team": merged["away_team_norm"],
        "opponent": merged["home_team_norm"],
        "venue": "away",
        "xg_pre_market_for": merged["xg_pre_market_away"],
        "xg_pre_market_against": merged["xg_pre_market_home"],
        "xg_actual_for": merged["away_xg"],
        "xg_actual_against": merged["home_xg"],
        "goals_for": merged["away_goals"],
        "goals_against": merged["home_goals"],
    })

    long_df = pd.concat([home_rows, away_rows], ignore_index=True)
    long_df["xg_delta_for"] = long_df["xg_actual_for"] - long_df["xg_pre_market_for"]
    long_df["xg_delta_against"] = long_df["xg_actual_against"] - long_df["xg_pre_market_against"]

    long_df = long_df.sort_values(["team", "date"]).reset_index(drop=True)

    long_df["xg_delta_for_rolling"] = (
        long_df.groupby("team")["xg_delta_for"]
        .transform(lambda s: s.rolling(rolling_window, min_periods=1).mean())
    )
    long_df["xg_delta_against_rolling"] = (
        long_df.groupby("team")["xg_delta_against"]
        .transform(lambda s: s.rolling(rolling_window, min_periods=1).mean())
    )

    return long_df


def main(prematch_path, postmatch_path, output_path, rolling_window):
    merged = load_and_merge(prematch_path, postmatch_path)
    long_df = to_long_form(merged, rolling_window)
    long_df.to_csv(output_path, index=False)
    print(f"Wrote {len(long_df)} team-match rows to {output_path}")
    print(f"Columns: {list(long_df.columns)}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--prematch", required=True)
    parser.add_argument("--postmatch", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--rolling-window", type=int, default=6,
                         help="Number of matches for the rolling form average")
    args = parser.parse_args()

    main(args.prematch, args.postmatch, args.output, args.rolling_window)
