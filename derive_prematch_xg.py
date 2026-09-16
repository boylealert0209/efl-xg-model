"""
derive_prematch_xg.py

Purpose
-------
Take pre-match betting odds (1X2, and optionally an Over/Under 2.5 total-goals
line) and back out a pair of pre-match "expected goals" numbers for the home
and away team — i.e. the two Poisson goal-expectancies (lambda_home,
lambda_away) that, when run through a Dixon-Coles adjusted bivariate Poisson
scoreline matrix, reproduce the de-vigged match-odds probabilities.

This is NOT shot-based xG. It's a market-implied expected goals split. Label
it that way in your schema (e.g. `xg_pre_market_home`, not `xg_home`) so it
never gets confused with the FBref/StatsBomb shot-based xG you'll pull after
the match.

Input CSV format
-----------------
Designed to work directly with football-data.co.uk historical odds files
(e.g. the Championship "E1.csv" files). Expected columns (only these are
used, others are ignored):

    Date, HomeTeam, AwayTeam, B365H, B365D, B365A

Optional, if present, will be used to pin down the split more precisely
(otherwise a fallback assumption is used — see note below):

    BbAv>2.5, BbAv<2.5   (or "Avg>2.5"/"Avg<2.5" depending on file vintage)

Usage
-----
    python derive_prematch_xg.py --input data/E1_2024_25.csv --output data/prematch_xg.csv

Method
------
1. Convert 1X2 odds to implied probabilities and de-vig (remove the
   overround) using simple proportional (multiplicative) normalisation.
   This is a reasonable default; Shin's method is noted as an extension.

2. If an Over/Under 2.5 line is available, de-vig it the same way to get
   P(Over 2.5) and P(Under 2.5). This pins down the *total* goal expectancy
   independently of the home/away split, which makes the two-parameter
   solve well-determined.

3. Solve numerically for (lambda_home, lambda_away) such that a Dixon-Coles
   adjusted bivariate Poisson scoreline grid reproduces:
       - P(home win), P(draw), P(away win)  [from step 1]
       - P(total goals > 2.5)               [from step 2, if available]

   If no O/U line is available, we fall back to a single free parameter:
   assume a fixed total-goals expectancy prior (league-average, configurable)
   and solve only for the home/away *split* that matches the 1X2 odds. This
   is materially less accurate than using an O/U line — get one if you can.

Notes
-----
- Dixon-Coles rho (low-score correlation adjustment) is fixed at -0.1 by
  default; this is a commonly used approximation. You can refit it properly
  later once you have a season of results, by maximising likelihood against
  actual scorelines.
- This script has NOT been run against live data in this environment
  (no network access here) — test it locally against a real
  football-data.co.uk file before trusting the output.
"""

import argparse
import numpy as np
import pandas as pd
from scipy.optimize import minimize
from scipy.stats import poisson

from leagues import get_league, league_choices

RHO_DEFAULT = -0.1
MAX_GOALS = 10  # scoreline grid truncation


def devig_proportional(odds_list):
    """Convert a list of decimal odds to de-vigged (fair) probabilities
    using simple proportional normalisation."""
    implied = [1.0 / o for o in odds_list]
    overround = sum(implied)
    return [p / overround for p in implied]


def dixon_coles_adjustment(x, y, lam, mu, rho):
    """Low-score correlation adjustment factor tau(x,y)."""
    if x == 0 and y == 0:
        return 1 - lam * mu * rho
    elif x == 0 and y == 1:
        return 1 + lam * rho
    elif x == 1 and y == 0:
        return 1 + mu * rho
    elif x == 1 and y == 1:
        return 1 - rho
    else:
        return 1.0


def score_matrix(lam_home, lam_away, rho=RHO_DEFAULT, max_goals=MAX_GOALS):
    """Build the Dixon-Coles adjusted scoreline probability matrix."""
    home_probs = poisson.pmf(np.arange(max_goals + 1), lam_home)
    away_probs = poisson.pmf(np.arange(max_goals + 1), lam_away)
    matrix = np.outer(home_probs, away_probs)

    for x in range(2):
        for y in range(2):
            matrix[x, y] *= dixon_coles_adjustment(x, y, lam_home, lam_away, rho)

    matrix /= matrix.sum()  # renormalise after adjustment
    return matrix


def probs_from_matrix(matrix):
    """Extract P(home win), P(draw), P(away win), P(over 2.5) from grid."""
    n = matrix.shape[0]
    home_win = sum(matrix[i, j] for i in range(n) for j in range(n) if i > j)
    draw = sum(matrix[i, j] for i in range(n) for j in range(n) if i == j)
    away_win = sum(matrix[i, j] for i in range(n) for j in range(n) if i < j)
    over_2_5 = sum(matrix[i, j] for i in range(n) for j in range(n) if i + j > 2.5)
    return home_win, draw, away_win, over_2_5


def solve_lambdas(target_home_p, target_draw_p, target_away_p, target_over_p=None,
                   fallback_total_goals=2.6):
    """Numerically solve for (lambda_home, lambda_away).

    If target_over_p is provided, both parameters are free and fit against
    (home, draw, away, over) jointly.

    If target_over_p is None, total goal expectancy is fixed at
    fallback_total_goals (a league-average prior you should calibrate per
    league/season) and only the home/away split is fit against (home, draw,
    away).
    """

    if target_over_p is not None:
        def loss(params):
            lam_h, lam_a = params
            if lam_h <= 0 or lam_a <= 0:
                return 1e6
            m = score_matrix(lam_h, lam_a)
            h, d, a, o = probs_from_matrix(m)
            return ((h - target_home_p) ** 2 + (d - target_draw_p) ** 2 +
                     (a - target_away_p) ** 2 + (o - target_over_p) ** 2)

        result = minimize(loss, x0=[1.4, 1.2], method="Nelder-Mead")
        return result.x[0], result.x[1]

    else:
        def loss(split):
            # split = lambda_home / (lambda_home + lambda_away), in (0,1)
            lam_h = fallback_total_goals * split[0]
            lam_a = fallback_total_goals * (1 - split[0])
            if lam_h <= 0 or lam_a <= 0:
                return 1e6
            m = score_matrix(lam_h, lam_a)
            h, d, a, _ = probs_from_matrix(m)
            return (h - target_home_p) ** 2 + (d - target_draw_p) ** 2 + (a - target_away_p) ** 2

        result = minimize(loss, x0=[0.55], method="Nelder-Mead",
                           bounds=[(0.05, 0.95)])
        split = result.x[0]
        return fallback_total_goals * split, fallback_total_goals * (1 - split)


def process_file(input_path, output_path, fallback_total_goals=2.6):
    df = pd.read_csv(input_path)

    # Handle a couple of common column-name variants across football-data.co.uk vintages
    over_col = next((c for c in ["BbAv>2.5", "Avg>2.5", "AvgU2.5"] if c in df.columns), None)
    under_col = next((c for c in ["BbAv<2.5", "Avg<2.5", "AvgU2.5"] if c in df.columns), None)

    rows = []
    for _, row in df.iterrows():
        try:
            h_odds, d_odds, a_odds = row["B365H"], row["B365D"], row["B365A"]
            if pd.isna(h_odds) or pd.isna(d_odds) or pd.isna(a_odds):
                continue

            p_home, p_draw, p_away = devig_proportional([h_odds, d_odds, a_odds])

            p_over = None
            if over_col and under_col and not pd.isna(row.get(over_col)) and not pd.isna(row.get(under_col)):
                p_over_raw, p_under_raw = devig_proportional([row[over_col], row[under_col]])
                p_over = p_over_raw

            lam_home, lam_away = solve_lambdas(
                p_home, p_draw, p_away, target_over_p=p_over,
                fallback_total_goals=fallback_total_goals
            )

            # football-data.co.uk dates are DD/MM/YYYY. Convert to ISO
            # (YYYY-MM-DD) here, once, so nothing downstream has to guess
            # day-first vs month-first and risk misreading e.g. 03/04/2026.
            raw_date = row.get("Date")
            parsed_date = pd.to_datetime(raw_date, dayfirst=True).strftime("%Y-%m-%d")

            rows.append({
                "date": parsed_date,
                "home_team": row.get("HomeTeam"),
                "away_team": row.get("AwayTeam"),
                "xg_pre_market_home": round(lam_home, 3),
                "xg_pre_market_away": round(lam_away, 3),
                "xg_pre_market_total": round(lam_home + lam_away, 3),
                "used_ou_line": p_over is not None,
                "market_p_home_win": round(p_home, 3),
                "market_p_draw": round(p_draw, 3),
                "market_p_away_win": round(p_away, 3),
            })
        except Exception as e:
            print(f"Skipping row due to error: {e}")
            continue

    OUTPUT_COLUMNS = [
        "date", "home_team", "away_team", "xg_pre_market_home", "xg_pre_market_away",
        "xg_pre_market_total", "used_ou_line", "market_p_home_win", "market_p_draw",
        "market_p_away_win",
    ]
    # pd.DataFrame([]) has NO columns at all (not even empty ones), so an
    # empty `rows` list - e.g. an upcoming-fixtures file with nothing in it
    # between rounds - would otherwise make every later `out_df[...]`
    # reference below raise KeyError instead of just writing an empty,
    # correctly-headered CSV.
    out_df = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)
    out_df.to_csv(output_path, index=False)
    print(f"Wrote {len(out_df)} matches to {output_path}")
    if len(out_df) > 0:
        print(f"  -> {out_df['used_ou_line'].sum()} used an O/U line for the split; "
              f"{(~out_df['used_ou_line']).sum()} used the fallback total-goals prior.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", required=True, help="Path to football-data.co.uk style odds CSV")
    parser.add_argument("--output", required=True, help="Path to write derived pre-match xG CSV")
    parser.add_argument("--league", default="championship",
                         choices=league_choices(),
                         help="Used only to pick the default --fallback-total-goals for this division.")
    parser.add_argument("--fallback-total-goals", type=float, default=None,
                         help="League-average total goals prior, used only when no O/U line is present. "
                              "Defaults to a per-league approximation from leagues.py.")
    args = parser.parse_args()

    fallback = args.fallback_total_goals
    if fallback is None:
        fallback = get_league(args.league)["fallback_total_goals"]

    process_file(args.input, args.output, fallback_total_goals=fallback)
