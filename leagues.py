"""
leagues.py

Purpose
-------
Single source of truth for the per-league constants the rest of the
pipeline needs, so every script can do `--league league-one` instead of
having the Championship hardcoded everywhere. Add a new league by adding
one entry here.

Where these numbers come from
------------------------------
- `fd_code`: football-data.co.uk's league code (the "E1.csv" part of
  https://www.football-data.co.uk/mmz4281/<season>/<code>.csv).
  Confirmed from football-data.co.uk's own key: E0 = Premier League,
  E1 = Championship, E2 = League One, E3 = League Two.
- `fotmob_id` / `fotmob_slug`: FotMob's internal league id and the slug
  it uses in URLs like https://www.fotmob.com/leagues/<id>/matches/<slug>.
  Confirmed by checking FotMob's own league pages (Championship=48,
  League One=108/"league-one", League Two=109/"league-two").
- `fbref_comp_id` / `fbref_slug`: FBref's competition id and the name
  used in its Scores & Fixtures URL
  (https://fbref.com/en/comps/<id>/schedule/<slug>-Scores-and-Fixtures).
  Championship=10, League One=15, League Two=16.
- `fallback_total_goals`: the league-average total-goals prior
  derive_prematch_xg.py falls back to when a match has no Over/Under 2.5
  line. These are rough, easily-outdated approximations (lower EFL
  divisions have historically averaged a bit fewer goals/game than the
  Championship) — recalibrate from a season of your own data once you
  have one, per derive_prematch_xg.py's own notes on refitting.

Known gap: FBref's Scores & Fixtures table has an xG column for the
Championship (Opta/StatsBomb-sourced) but, as of this writing, does NOT
have one for League One or League Two — so scrape_fbref.py will find no
matches there for those two leagues. FotMob's match pages do carry xG
for League One/Two (confirmed via FotMob's own League One/League Two
stats pages), which is why scrape_fotmob.py is the pipeline's actual
post-match xG source for every league, not just the Championship.

Premier League / Scottish Premiership / National League (added later)
------------------------------------------------------------------------
- Premier League: fd_code E0, FotMob id 47 / slug "premier-league" -
  both confirmed directly against football-data.co.uk's own file
  listing and FotMob's own Premier League pages.
- Scottish Premiership (what most people still call "the SPL", even
  though that name was retired in 2013): fd_code SC0 (football-data.co.uk
  files the top 4 Scottish tiers as SC0-SC3), FotMob id 64 / slug
  "premiership" - confirmed against FotMob's own Scottish Premiership
  pages (FotMob calls it just "Premiership", not "Scottish Premiership",
  in its URLs/labels, since it's Scotland's own top flight).
- National League (English football's 5th tier, still commonly called
  "the Conference"): fd_code EC (football-data.co.uk's own key lists
  "England (E0, E1, E2, E3 & EC)"). FotMob id 117 / slug
  "national-league" is confirmed to exist as a real league on FotMob.

  IMPORTANT / UNCONFIRMED: whether FotMob's match pages actually carry
  a shot-based "Expected goals (xG)" stat for National League matches
  (the way they demonstrably do for the Championship/League One/League
  Two/Premier League/Scottish Premiership) could NOT be confirmed from
  here - that needs an actual FotMob National League match page open in
  a browser, or scrape_fotmob.py run against it, to check. Third-party
  xG sites that list "National League xG" (e.g. FootyStats) use their
  own possession/shot-frequency-based model, not FotMob's shot-based
  xG, so their coverage existing doesn't confirm FotMob's does. Treat
  national-league as experimental: try `python scrape_fotmob.py
  --league national-league --output data/fotmob_national-league_xg.csv`
  first and check how many matches actually come back with an xG value
  before building a weekly workflow around it. If FotMob has no xG for
  this tier, the rest of the pipeline (results/League Table/fixtures)
  still works - you'd just have an empty/near-empty
  team_match_deltas_national-league.csv and report.
- fbref_comp_id/fbref_slug are left as None for these three - fbref
  scraping was never wired up as the real post-match source for any
  league here anyway (see scrape_fbref.py's docstring); FotMob is.
- fallback_total_goals are rough starting priors (Premier League and
  the Scottish Premiership tend to average a bit more goals/game than
  the Championship; National League's is a guess) - recalibrate once
  you have a season of real data, same advice as the original three.
"""

LEAGUES = {
    "championship": {
        "label": "EFL Championship",
        "fd_code": "E1",
        "fotmob_id": 48,
        "fotmob_slug": "championship",
        "fbref_comp_id": 10,
        "fbref_slug": "Championship",
        "fallback_total_goals": 2.6,
    },
    "league-one": {
        "label": "EFL League One",
        "fd_code": "E2",
        "fotmob_id": 108,
        "fotmob_slug": "league-one",
        "fbref_comp_id": 15,
        "fbref_slug": "League-One",
        "fallback_total_goals": 2.5,
    },
    "league-two": {
        "label": "EFL League Two",
        "fd_code": "E3",
        "fotmob_id": 109,
        "fotmob_slug": "league-two",
        "fbref_comp_id": 16,
        "fbref_slug": "League-Two",
        "fallback_total_goals": 2.5,
    },
    "premier-league": {
        "label": "Premier League",
        "fd_code": "E0",
        "fotmob_id": 47,
        "fotmob_slug": "premier-league",
        "fbref_comp_id": None,
        "fbref_slug": None,
        "fallback_total_goals": 2.8,
    },
    "spl": {
        "label": "Scottish Premiership",
        "fd_code": "SC0",
        "fotmob_id": 64,
        "fotmob_slug": "premiership",
        "fbref_comp_id": None,
        "fbref_slug": None,
        "fallback_total_goals": 2.7,
    },
    "national-league": {
        "label": "National League",
        "fd_code": "EC",
        "fotmob_id": 117,
        "fotmob_slug": "national-league",
        "fbref_comp_id": None,
        "fbref_slug": None,
        "fallback_total_goals": 2.6,
        # See the "IMPORTANT / UNCONFIRMED" note above - flagged here too
        # so any script can warn about it programmatically if useful.
        "xg_coverage_unconfirmed": True,
    },
}


def get_league(key: str) -> dict:
    try:
        return LEAGUES[key]
    except KeyError:
        valid = ", ".join(LEAGUES)
        raise SystemExit(f"Unknown --league '{key}'. Valid options: {valid}")


def league_choices() -> list:
    """All valid --league keys, for argparse choices= - one place to add a
    league (above) instead of editing every script's argparse block."""
    return list(LEAGUES.keys())
