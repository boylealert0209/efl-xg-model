# xG Form Model

Tracks each team's shot-based xG performance vs. the pre-match market-implied
expectation, to surface opponent-adjusted attacking/defensive form — for the
**Championship**, **League One**, **League Two**, **Premier League**,
**Scottish Premiership ("SPL")**, or **National League**. Pick the league with
`--league championship|league-one|league-two|premier-league|spl|national-league`
on every script; everything else (football-data.co.uk code, FotMob league id,
output filenames) follows automatically from `leagues.py`. Run any script
with `--help` to see the full up-to-date list of `--league` choices — it's
read straight from `leagues.py`, so it never goes stale.

**National League note:** whether FotMob actually publishes shot-based xG
for this tier (English football's 5th) could not be confirmed while building
this — see the long comment in `leagues.py`. Run `python scrape_fotmob.py
--league national-league --output data/fotmob_national-league_xg.csv` and
check how many matches come back with an xG value before trusting the
scheduled workflow for it.

## Pipeline

Run each step by hand, or see **Automated weekly update** below to run (and
schedule) all of them in one go, for one league at a time.

1. **Get odds/results data**: download the current season file from
   [football-data.co.uk](https://www.football-data.co.uk/englandm.php)
   (`E0`/`E1`/`E2`/`E3`/`EC` for England, `SC0` for Scotland - see
   `leagues.py`) into `data/`. Automated by `download_results.py`:
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

5. **Get upcoming fixtures + odds** (for the "Upcoming Fixtures" sheet -
   optional; skip steps 5-6 and drop `--fixtures` from step 7 if you don't
   want that sheet):
   ```
   python download_fixtures.py --league league-one --output data/fixtures_league-one.csv
   ```

6. **Derive pre-match expected goals for those fixtures** (same script as
   step 2 - it only ever reads odds columns, so it works unchanged on a
   file of matches that haven't been played yet):
   ```
   python derive_prematch_xg.py --league league-one --input data/fixtures_league-one.csv --output data/prematch_xg_fixtures_league-one.csv
   ```

7. **Build the Excel report** (one tab per team, plus League Table,
   Overview, Home & Away Splits, the Supremacy/Total xG Delta grids, and
   Upcoming Fixtures if step 6's file is passed — also league-agnostic):
   ```
   python build_team_workbook.py \
       --input data/team_match_deltas_league-one.csv \
       --results data/E2.csv \
       --fixtures data/prematch_xg_fixtures_league-one.csv \
       --output data/league-one_xg_report.xlsx \
       --rolling-window 6
   ```

## Automated weekly update

`run_weekly_update.py --league <league>` runs steps 1-7 above in order for
that league (fixtures included), and stops with a clear error if any step
fails rather than silently producing a stale report:
```
python run_weekly_update.py --league championship
python run_weekly_update.py --league league-one
python run_weekly_update.py --league league-two
python run_weekly_update.py --league premier-league
python run_weekly_update.py --league spl
python run_weekly_update.py --league national-league
```

Six separate GitHub Actions workflows in `.github/workflows/` each run one
league's update every Tuesday (06:00/06:15/06:30/06:45/07:00/07:15 UTC —
staggered so they don't all hit football-data.co.uk/FotMob at once, the
morning after each weekend's gameweek), and commit the refreshed `data/`
files — including that league's `_xg_report.xlsx` — straight back to the
repo:
- `weekly_update_championship.yml`
- `weekly_update_league_one.yml`
- `weekly_update_league_two.yml`
- `weekly_update_premier_league.yml`
- `weekly_update_spl.yml`
- `weekly_update_national_league.yml` (experimental - see the National
  League note above; delete this file if you'd rather not run it
  unattended)

You don't need to do anything for them to run once they're on GitHub; you
can also trigger any one manually from the repo's **Actions** tab (pick the
workflow by name -> **Run workflow**). If a league currently has no
upcoming fixtures with posted odds (e.g. between rounds), step 5 writes an
empty file rather than failing, and the report is built without an
Upcoming Fixtures sheet that week.

First-time setup on GitHub:
1. Push this repo (including the `.github/workflows/` folder) to GitHub.
2. In the repo's **Settings -> Actions -> General -> Workflow permissions**,
   make sure **Read and write permissions** is selected, so the workflows can
   commit their output back.
3. That's it — each will run on its own schedule from then on. Check the
   **Actions** tab to see run history and logs if a step ever fails (e.g. if
   football-data.co.uk or FotMob change their page format). If you only want
   some of the six leagues, just delete the workflow file(s) you don't need
   — or edit its `cron:` line to change the schedule.

## Excel report (`<league>_xg_report.xlsx`)

Built by `build_team_workbook.py`. Every sheet is explained in the
workbook's own **"About This Report"** tab (the first sheet) — including
the exact judgement calls made anywhere the brief for this report was
ambiguous (sort orders, what "expected supremacy" means for upcoming
fixtures, etc.) — but in short:

- **League Table** - standard standings from every played match.
- **Overview** - every team's season xG form (average, median, and
  current rolling form), sorted by Avg xG Delta (For), best first.
- **Home & Away Splits** - league-wide Total/Home/Away average + median
  supremacy, and Total/Home/Away goals for and against, one row per team.
- **Supremacy Delta** / **(Home)** / **(Away)** - match-by-match grids of
  actual vs. expected xG supremacy, colour-scaled green (better than
  expected) to red (worse), zero-anchored so the colour always means the
  same thing regardless of that column's own spread. Sorted the same way
  as Overview.
- **Total xG Delta** - match-by-match grid of how much higher/lower-
  scoring each match was than the market's pre-match total-goals
  expectation, same colour-scaling, sorted by its own average (highest
  first).
- **Upcoming Fixtures** - only present if a fixtures pre-match xG file was
  passed with `--fixtures`. Every fixture ranked by a "Combined Expected
  Supremacy" that blends the market's own pre-match view of that specific
  fixture with how much each side has been over/under-performing its own
  expectation all season - see the About sheet for the exact formula and
  its limits (it's a transparent blend, not a fitted prediction model).
- **One tab per team** with its full match log (opponent, venue, goals,
  pre-match vs. actual xG, and the delta/rolling-delta columns), a
  season-summary block using real formulas (average AND median of every
  delta stat) over that team's table, and a Home/Away split mini-table.

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

- **`scrape_fbref.py` doesn't currently work for League One, League Two,
  Premier League, the Scottish Premiership, or the National League**:
  FBref's Scores & Fixtures table has an xG column for the Championship,
  but this was never confirmed/wired up as a real data source for any
  other league. This is exactly why the pipeline's real post-match xG
  source is `scrape_fotmob.py` for every league here, not just the
  Championship — `scrape_fbref.py` is kept only for the Championship / in
  case FBref adds coverage for the others later. See its docstring.
- **National League xG coverage on FotMob is unconfirmed** — see the note
  at the top of this file and in `leagues.py`.
- **Team name mismatches**: football-data.co.uk and FotMob use different
  spellings for some clubs. `build_deltas.py` has a `TEAM_NAME_MAP` dict,
  with starter entries for common short-name vs. full-name mismatches
  across all six leagues — but since which clubs sit in which division
  changes every season with promotion/relegation, treat that list as a
  starting point, not a verified one for the current season. Extend it as
  mismatches surface (the script warns you about unmatched fixtures). The
  National League entries in particular are thin — it's a 24-club
  division drawn from a much bigger non-league pyramid that changes a lot
  season to season.
- **1X2-only odds files won't pin down the split precisely**. Get an
  Over/Under 2.5 average odds column in your football-data.co.uk download if
  you can — the derive script uses it when present, and falls back to a
  fixed total-goals prior otherwise (materially less accurate). Lower
  divisions are more likely to be missing this column than the Premier
  League/Championship, since fewer bookmakers post O/U lines that far down.
- **`--fallback-total-goals` defaults are rough per-league guesses**
  (`leagues.py`) — recalibrate from a season of your own data once you have
  one, same as the Dixon-Coles rho note below.
- **Dixon-Coles rho is fixed at -0.1** as a placeholder. Once you have a
  full season of results, refit rho (and possibly the whole model) by
  maximum likelihood against actual scorelines rather than trusting the
  default.
- **Upcoming Fixtures data only covers the next round or two**:
  football-data.co.uk's fixtures.csv only carries odds for matches that
  have already had lines posted (typically a few days out), so
  `download_fixtures.py` won't give you a full season of future fixtures
  — just what's coming up soon. Re-run it (the weekly workflow does) to
  keep the sheet current.
- **Network access needed to test `download_results.py`, `scrape_fotmob.py`
  and `download_fixtures.py`** — this dev environment's sandbox couldn't
  reach any of those sites (network egress is restricted here), so run
  `python run_weekly_update.py --league <league>` locally at least once per
  league before trusting the scheduled GitHub Actions runs. `leagues.py`'s
  fd_code/FotMob id/slug values for the three newly-added leagues were
  checked directly against football-data.co.uk's own file listing and
  FotMob's own league pages, but the actual scrape/download scripts
  themselves could not be run end-to-end here — `build_team_workbook.py`
  was tested locally against the Championship/League One data already in
  this repo (including a full formula-recalculation pass), but not against
  live Premier League/SPL/National League data. Expect to need small fixes
  over time for FotMob page-structure changes (the scraper is explicitly
  best-effort — see its docstring) or football-data.co.uk
  season-code/column changes.
- **FotMob rate limiting**: `scrape_fotmob.py` fetches one page per match and
  sleeps between requests on purpose — don't remove that delay, or you risk
  getting temporarily blocked.
