# EFL xG Form Model

Tracks each team's shot-based xG performance vs. the pre-match market-implied
expectation, to surface opponent-adjusted attacking/defensive form — for the
**Championship**, **League One**, or **League Two**. Pick the division with
`--league championship|league-one|league-two` on every script; everything
else (football-data.co.uk code, FotMob league id, output filenames) follows
automatically from `leagues.py`.

## Pipeline

Run each step by hand, or see **Automated weekly update** below to run (and
schedule) all of them in one go, for one division at a time.

1. **Get odds/results data**: download the current season file from
   [football-data.co.uk](https://www.football-data.co.uk/englandm.php)
   (`E1.csv`/`E2.csv`/`E3.csv` for Championship/League One/League Two) into
   `data/`. Automated by `download_results.py`:
   ```
   python download_results.py --league league-one --output data/E2.csv
   ```

2. **Derive pre-match expected goals from odds**:
   ```
   python derive_prematch_xg.py --league league-one --input data/E2.csv --output data/prematch_xg_league-one.csv
   ```

3. **Scrape post-match shot-based xG from FotMob**:
   ```
   python scrape_fotmob.py --league league-one --output data/fotmob_league-one_xg.csv
   # or, for a fast weekly rerun that only fetches recent matches:
   python scrape_fotmob.py --league league-one --output data/fotmob_league-one_xg.csv --since 2026-08-25
   ```

4. **Merge and compute form deltas** (league-agnostic — just points at
   whichever prematch/postmatch CSVs you give it):
   ```
   python build_deltas.py \
       --prematch data/prematch_xg_league-one.csv \
       --postmatch data/fotmob_league-one_xg.csv \
       --output data/team_match_deltas_league-one.csv \
       --rolling-window 6
   ```

5. **Build the Excel report** (one tab per team, plus a league Overview —
   also league-agnostic):
   ```
   python build_team_workbook.py \
       --input data/team_match_deltas_league-one.csv \
       --results data/E2.csv \
       --output data/league-one_xg_report.xlsx \
       --rolling-window 6
   ```

## Automated weekly update

`run_weekly_update.py --league <league>` runs steps 1-5 above in order for
that division, and stops with a clear error if any step fails rather than
silently producing a stale report:
```
python run_weekly_update.py --league championship
python run_weekly_update.py --league league-one
python run_weekly_update.py --league league-two
```

Three separate GitHub Actions workflows in `.github/workflows/` each run one
league's update every Tuesday (06:00/06:15/06:30 UTC — staggered so they
don't all hit football-data.co.uk/FotMob at once, the morning after each
weekend's gameweek), and commit the refreshed `data/` files — including that
league's `_xg_report.xlsx` — straight back to the repo:
- `weekly_update_championship.yml`
- `weekly_update_league_one.yml`
- `weekly_update_league_two.yml`

You don't need to do anything for them to run once they're on GitHub; you
can also trigger any one manually from the repo's **Actions** tab (pick the
workflow by name -> **Run workflow**).

First-time setup on GitHub:
1. Push this repo (including the `.github/workflows/` folder) to GitHub.
2. In the repo's **Settings -> Actions -> General -> Workflow permissions**,
   make sure **Read and write permissions** is selected, so the workflows can
   commit their output back.
3. That's it — each will run on its own schedule from then on. Check the
   **Actions** tab to see run history and logs if a step ever fails (e.g. if
   football-data.co.uk or FotMob change their page format). If you only want
   one or two of the three divisions, just delete the workflow file(s) you
   don't need — or edit its `cron:` line to change the schedule.

## Excel report (`<league>_xg_report.xlsx`)

Built by `build_team_workbook.py`:
- **One tab per team** with its full match log (opponent, venue, goals, pre-match
  vs. actual xG, and the delta/rolling-delta columns), colour-scaled so
  over-/under-performance is visible at a glance, plus a season-summary block
  at the bottom using real formulas (`AVERAGE`, `SUM`) over that team's table.
- **An "Overview" tab** ranking every team by season and current-form xG
  deltas, pulled live from each team's tab via cell-reference formulas.

## Output schema (`team_match_deltas_<league>.csv`)

One row per team per match:

| column | meaning |
|---|---|
| `xg_pre_market_for` | market-implied expected goals scored, pre-match |
| `xg_pre_market_against` | market-implied expected goals conceded, pre-match |
| `xg_actual_for` | shot-based xG actually generated (FotMob) |
| `xg_actual_against` | shot-based xG actually conceded (FotMob) |
| `xg_delta_for` | actual − expected (attacking overperformance) |
| `xg_delta_against` | actual conceded − expected conceded (negative = defended better than expected) |
| `xg_delta_for_rolling` / `xg_delta_against_rolling` | rolling average over `--rolling-window` matches — this is your "form" signal |

## Known gaps / next steps

- **`scrape_fbref.py` doesn't currently work for League One or League Two**:
  FBref's Scores & Fixtures table has xG columns for the Championship, but
  not (as of this writing) for League One or League Two. This is exactly why
  the pipeline's real post-match xG source is `scrape_fotmob.py` (FotMob does
  carry xG for all three divisions) for every league here, not just the
  Championship — `scrape_fbref.py` is kept only for the Championship / in
  case FBref adds lower-league xG coverage later. See its docstring.
- **Team name mismatches**: football-data.co.uk and FotMob use different
  spellings for some clubs. `build_deltas.py` has a `TEAM_NAME_MAP` dict,
  now with some starter entries for common League One/Two short-name vs.
  full-name mismatches (e.g. "Doncaster" vs "Doncaster Rovers") — but since
  which clubs sit in which division changes every season with promotion/
  relegation, treat that list as a starting point, not a verified one for
  the current season. Extend it as mismatches surface (the script warns you
  about unmatched fixtures).
- **1X2-only odds files won't pin down the split precisely**. Get an
  Over/Under 2.5 average odds column in your football-data.co.uk download if
  you can — the derive script uses it when present, and falls back to a
  fixed total-goals prior otherwise (materially less accurate). Lower
  divisions are more likely to be missing this column than the Championship,
  since fewer bookmakers post O/U lines that far down.
- **`--fallback-total-goals` defaults are rough per-league guesses**
  (`leagues.py`) — recalibrate from a season of your own data once you have
  one, same as the Dixon-Coles rho note below.
- **Dixon-Coles rho is fixed at -0.1** as a placeholder. Once you have a
  full season of results, refit rho (and possibly the whole model) by
  maximum likelihood against actual scorelines rather than trusting the
  default.
- **`download_results.py` and `scrape_fotmob.py` need real network access to
  test** — this dev environment's sandbox couldn't reach either site
  (network egress is restricted here), so run `python run_weekly_update.py
  --league <league>` locally at least once per division before trusting the
  scheduled GitHub Actions runs. Expect to need small fixes over time for
  FotMob page-structure changes (the scraper is explicitly best-effort —
  see its docstring) or football-data.co.uk season-code/column changes.
- **FotMob rate limiting**: `scrape_fotmob.py` fetches one page per match and
  sleeps between requests on purpose — don't remove that delay, or you risk
  getting temporarily blocked.
