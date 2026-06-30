"""Major-league teams lookup from data/major_league_teams.xlsx.

Used by the team-rename vendor rule: if a row's Vendor is one of the
TEAM_RENAME_VENDORS (e.g. 'Ticketmaster AM', 'Tickets.com', 'Ballpark')
and Team/Performer is one of these teams, rename the Vendor to the
Team/Performer. Otherwise leave Vendor unchanged.

File schema: two columns 'League', 'Team'. We only care about Team.
Membership is case-insensitive — we lowercase both the lookup set and
the input.
"""
from __future__ import annotations

import logging
import os
from functools import lru_cache
from pathlib import Path

import pandas as pd

log = logging.getLogger(__name__)

DEFAULT_PATH = Path("data") / "major_league_teams.xlsx"

# Vendors that get rewritten to the Team/Performer when the performer is
# a known major-league team. Stored lowercased for case-insensitive match.
# To add a vendor to this rule, add its lowercased name here.
TEAM_RENAME_VENDORS: frozenset[str] = frozenset({
    "ticketmaster am",
    "tickets.com",
    "ballpark",
})


def load_teams(path: Path | str | None = None) -> set[str]:
    """Return the set of team names (lowercased) for fast lookup."""
    path = Path(path) if path else Path(os.getenv("TEAMS_PATH", DEFAULT_PATH))
    if not path.exists():
        log.warning("Teams file not found at %s; using empty set.", path)
        return set()

    df = pd.read_excel(path)
    if "Team" not in df.columns:
        log.warning("Teams file at %s missing 'Team' column; using empty set.", path)
        return set()

    teams = {
        str(t).strip().lower()
        for t in df["Team"].dropna().tolist()
        if str(t).strip()
    }
    log.info("Loaded %d teams from %s", len(teams), path)
    return teams


@lru_cache(maxsize=1)
def get_teams() -> set[str]:
    """Cached accessor used by the processor."""
    return load_teams()


# Keywords that mark a (non-major-league) Team/Performer as a college team, so
# season-ticket groups for college sports get a "College" label. Mirrors the
# Purchase Details app.
COLLEGE_KEYWORDS = ["college", "university", "football", "basketball", "hockey", "baseball"]


def load_team_leagues(path: Path | str | None = None) -> dict[str, str]:
    """Return a {team_lowercased: league} map from the major-league teams file.

    The file's 'League' column holds the league (MLB, NBA, NFL, NHL, MLS). Used
    to tag detected season-ticket groups with their league.
    """
    path = Path(path) if path else Path(os.getenv("TEAMS_PATH", DEFAULT_PATH))
    if not path.exists():
        log.warning("Teams file not found at %s; using empty league map.", path)
        return {}
    df = pd.read_excel(path)
    if "Team" not in df.columns or "League" not in df.columns:
        log.warning("Teams file at %s missing 'Team'/'League'; using empty league map.", path)
        return {}
    out: dict[str, str] = {}
    for _, row in df.iterrows():
        team = str(row["Team"]).strip()
        league = str(row["League"]).strip()
        if team and league and league.lower() != "nan":
            out[team.lower()] = league
    log.info("Loaded %d team→league entries from %s", len(out), path)
    return out


@lru_cache(maxsize=1)
def get_team_leagues() -> dict[str, str]:
    """Cached {team_lowercased: league} accessor."""
    return load_team_leagues()


def league_for_team(team_performer, leagues: dict[str, str] | None = None) -> str:
    """Return the league label for a team ('MLB'/'NBA'/'NFL'/'NHL'/'MLS'), or
    'College' for college teams (by keyword), or '' if neither. Case-insensitive."""
    if not isinstance(team_performer, str):
        return ""
    if leagues is None:
        leagues = get_team_leagues()
    key = team_performer.strip().lower()
    if key in leagues:
        return leagues[key]
    if any(kw in key for kw in COLLEGE_KEYWORDS):
        return "College"
    return ""


def reset_cache() -> None:
    """Clear the caches (used in tests)."""
    get_teams.cache_clear()
    get_team_leagues.cache_clear()


def rename_team_vendor(
    vendor: str | None,
    team_performer: str | None,
    teams: set[str] | None = None,
) -> str | None:
    """Apply the team-vendor rename rule:
       - If Vendor is in TEAM_RENAME_VENDORS (e.g. 'Ticketmaster AM',
         'Tickets.com', 'Ballpark') AND Team/Performer matches a known
         major-league team → return Team/Performer.
       - Otherwise → return Vendor unchanged.

    Case-insensitive on both the vendor check and the team lookup.

    NOTE: The processor no longer calls this — vendor renaming is handled by
    the full multi-stage pipeline in vendor_rules.apply_vendor_pipeline (which
    also keys off VenueName and Original Company). This single-stage helper is
    retained for reference / standalone use; get_teams() above is still the
    shared teams accessor used by both.
    """
    if vendor is None or (isinstance(vendor, float) and pd.isna(vendor)):
        return vendor
    if str(vendor).strip().lower() not in TEAM_RENAME_VENDORS:
        return vendor
    if team_performer is None or (isinstance(team_performer, float) and pd.isna(team_performer)):
        return vendor
    if teams is None:
        teams = get_teams()
    if str(team_performer).strip().lower() in teams:
        return str(team_performer).strip()
    return vendor


# Backwards-compatible alias for the original single-vendor function.
rename_ticketmaster_am = rename_team_vendor
