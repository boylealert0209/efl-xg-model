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
import math
import numpy as np
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
    "Nottm Forest": "Nottingham Forest",  # FotMob's spelling (no apostrophe) - "Nott'm Forest" above is football-data.co.uk's
    # Scottish Premiership starter entries.
    "Hearts": "Heart of Midlothian",
    "Dundee Utd": "Dundee United",
    "St. Mirren": "St Mirren",
    "St. Johnstone": "St Johnstone",
    "Dundee FC": "Dundee",  # FotMob disambiguates "Dundee" from "Dundee United" this way
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
    # Both inputs use unambiguous ISO (YYYY-MM-DD[THH:MM:SS[.fff]]) dates by
    # the time they reach this script (derive_prematch_xg.py normalises the
    # odds file's dates; FotMob's are already ISO) - but FotMob mixes
    # "...T14:00:00Z" (finished matches) with "...T14:00:00.000Z"
    # (matches that just kicked off or are still live), sometimes within
    # the SAME column on the SAME run. Reading with parse_dates=[...] lets
    # pandas auto-detect a single format from the first few rows, and if a
    # later row doesn't match that format it silently leaves the WHOLE
    # column as plain strings instead of raising - which then breaks
    # every .dt access below with an unhelpful AttributeError far from the
    # real cause. format="mixed" parses each value independently instead
    # of assuming one shared format, so a mix of with/without milliseconds
    # (or with/without a timezone at all) in the same column is fine.
    pre = pd.read_csv(prematch_path)
    post = pd.read_csv(postmatch_path)
    for df in (pre, post):
        df["date"] = pd.to_datetime(df["date"], format="mixed", utc=True)

    # Normalise both date columns to plain (timezone-naive) calendar dates.
    # The odds file has plain dates; FotMob's dates carry UTC timezone info
    # (e.g. "...T14:00:00.000Z"). Both are tz-aware now (utc=True above
    # attaches UTC even to naive inputs), so tz_localize(None) always
    # applies - no need to check first.
    for df in (pre, post):
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


def add_opponent_adjustment(long_df, rolling_window):
    """Add opponent-strength-adjusted deltas.

    Raw xg_delta_for/against tell you how a team did vs. the market's
    expectation for THAT match, but say nothing about whether the opponent
    was strong or weak. A team can rack up a big positive xg_delta_for
    just by playing a side that leaks chances all season.

    Adjustment: subtract the opponent's own SEASON-TO-DATE tendency on the
    matching side, using leave-one-out averages so this exact match isn't
    used to estimate the opponent's own strength (avoids circularity):

        xg_delta_for_adj      = xg_delta_for      - opponent's average
                                 xg_delta_against, excluding this match
        xg_delta_against_adj  = xg_delta_against  - opponent's average
                                 xg_delta_for, excluding this match

    Why the criss-cross (delta_for adjusted by the opponent's delta
    AGAINST, not their delta_for)? Because what matters for judging a
    team's attacking output is how well the opponent generally defends,
    not how well the opponent generally attacks. If the opponent usually
    concedes less than expected (a strong defence, so a negative average
    delta_against), then a team's positive delta_for against them is
    extra impressive — subtracting a negative number increases the
    adjusted score. Same logic in reverse for defence.

    Because long_df has one row per team per match (see to_long_form), the
    mirror row for the same fixture always exists with team/opponent
    swapped, and by construction:

        this_row.xg_delta_against == opponent's own xg_delta_for for
                                      this exact match
        this_row.xg_delta_for     == opponent's own xg_delta_against for
                                      this exact match

    so the opponent's own contribution to their season total can be
    subtracted directly from this row without a self-join.

    Requires at least 2 matches played by the opponent; rows where the
    opponent has only 1 match get NaN (nothing to leave out and still
    have a season-to-date average).
    """
    team_stats = long_df.groupby("team").agg(
        total_delta_for=("xg_delta_for", "sum"),
        total_delta_against=("xg_delta_against", "sum"),
        n_matches=("xg_delta_for", "size"),
    )

    opp_total_for = long_df["opponent"].map(team_stats["total_delta_for"])
    opp_total_against = long_df["opponent"].map(team_stats["total_delta_against"])
    opp_n = long_df["opponent"].map(team_stats["n_matches"])

    denom = (opp_n - 1).where(opp_n > 1)  # NaN when opponent has only 1 match

    # Opponent's leave-one-out average delta_against, excluding THIS match
    # (this row's own xg_delta_for is the opponent's delta_against for
    # this exact fixture, per the docstring above).
    opp_avg_delta_against_excl = (opp_total_against - long_df["xg_delta_for"]) / denom
    # Opponent's leave-one-out average delta_for, excluding THIS match.
    opp_avg_delta_for_excl = (opp_total_for - long_df["xg_delta_against"]) / denom

    long_df["opponent_avg_xg_delta_for"] = opp_avg_delta_for_excl
    long_df["opponent_avg_xg_delta_against"] = opp_avg_delta_against_excl
    long_df["xg_delta_for_adj"] = long_df["xg_delta_for"] - opp_avg_delta_against_excl
    long_df["xg_delta_against_adj"] = long_df["xg_delta_against"] - opp_avg_delta_for_excl

    long_df = long_df.sort_values(["team", "date"]).reset_index(drop=True)
    long_df["xg_delta_for_adj_rolling"] = (
        long_df.groupby("team")["xg_delta_for_adj"]
        .transform(lambda s: s.rolling(rolling_window, min_periods=1).mean())
    )
    long_df["xg_delta_against_adj_rolling"] = (
        long_df.groupby("team")["xg_delta_against_adj"]
        .transform(lambda s: s.rolling(rolling_window, min_periods=1).mean())
    )

    return long_df


def _poisson_match_xpts(lambda_for, lambda_against, max_goals=10):
    """Vectorised expected points from a pair of xG values, via an
    independent-Poisson scoreline model (the standard, simple way to turn
    two xG numbers into win/draw/loss probabilities - not fitted to this
    league's actual scoring distribution, just the textbook approach).

    lambda_for/lambda_against: 1-D numpy arrays of xG (Poisson rate) for
    each row. Returns (xpts, p_win, p_draw) arrays, where
    xpts = 3*p_win + 1*p_draw.

    Method: build each side's goals-scored distribution P(0..max_goals)
    from the Poisson pmf, take the outer product per row to get the joint
    scoreline distribution, then sum the strictly-lower-triangular part
    (home goals > away goals) for P(win) and the diagonal for P(draw).
    max_goals=10 leaves <0.001 probability mass uncovered even for an
    unusually high xG match, so it doesn't meaningfully bias the result.
    """
    k = np.arange(0, max_goals + 1)
    factorial_k = np.array([math.factorial(int(i)) for i in k], dtype=float)

    def goal_dist(lam):
        lam = np.asarray(lam, dtype=float)[:, None]  # (N, 1)
        return np.exp(-lam) * (lam ** k[None, :]) / factorial_k[None, :]  # (N, K)

    dist_for = goal_dist(lambda_for)       # (N, K)
    dist_against = goal_dist(lambda_against)  # (N, K)

    # Joint scoreline probability per row: joint[n, i, j] = P(for=i, against=j)
    joint = dist_for[:, :, None] * dist_against[:, None, :]  # (N, K, K)

    win_mask = np.tril(np.ones((len(k), len(k))), k=-1)   # for-goals > against-goals
    draw_mask = np.eye(len(k))

    p_win = (joint * win_mask[None, :, :]).sum(axis=(1, 2))
    p_draw = (joint * draw_mask[None, :, :]).sum(axis=(1, 2))
    xpts = 3 * p_win + p_draw
    return xpts, p_win, p_draw


def add_expected_points(long_df):
    """Add xPts (expected points) columns, per the brief's item #5:
    "this gives a truer league table" than actual results, since it's
    based on the quality of chances created/conceded rather than the
    bounce of the ball on the day.

    Two versions, both per team-match row:
      xpts_actual  - deserved points from this match's ACTUAL (shot-based,
                      post-match) xG - "how many points did the way this
                      team actually played merit".
      xpts_market  - deserved points from the PRE-MATCH MARKET xG - "how
                      many points did the market/bookmakers expect this
                      team to earn going in".
      xpts_delta   = xpts_actual - xpts_market (over/under-performance vs.
                      the market, in points terms rather than xG terms -
                      complements xg_delta_for/against).

    Also a running cumulative total of each per team (season-to-date, in
    date order) - the "truer league table" itself: sort teams by
    xpts_actual_cumulative's final value.
    """
    xpts_actual, _, _ = _poisson_match_xpts(
        long_df["xg_actual_for"].to_numpy(), long_df["xg_actual_against"].to_numpy()
    )
    xpts_market, _, _ = _poisson_match_xpts(
        long_df["xg_pre_market_for"].to_numpy(), long_df["xg_pre_market_against"].to_numpy()
    )
    long_df["xpts_actual"] = xpts_actual.round(3)
    long_df["xpts_market"] = xpts_market.round(3)
    long_df["xpts_delta"] = (long_df["xpts_actual"] - long_df["xpts_market"]).round(3)

    long_df = long_df.sort_values(["team", "date"]).reset_index(drop=True)
    long_df["xpts_actual_cumulative"] = (
        long_df.groupby("team")["xpts_actual"].cumsum().round(3)
    )
    long_df["xpts_market_cumulative"] = (
        long_df.groupby("team")["xpts_market"].cumsum().round(3)
    )

    return long_df


def main(prematch_path, postmatch_path, output_path, rolling_window):
    merged = load_and_merge(prematch_path, postmatch_path)
    long_df = to_long_form(merged, rolling_window)
    long_df = add_opponent_adjustment(long_df, rolling_window)
    long_df = add_expected_points(long_df)
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
