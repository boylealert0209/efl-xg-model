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
}


def get_league(key: str) -> dict:
    try:
        return LEAGUES[key]
    except KeyError:
        valid = ", ".join(LEAGUES)
        raise SystemExit(f"Unknown --league '{key}'. Valid options: {valid}")
