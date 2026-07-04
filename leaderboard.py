"""
leaderboard.py
==============

Tournament leaderboard generation and standings management.

This module provides the foundational types, configuration objects, ranking
definitions, and helper abstractions required to build tournament standings.

Responsibilities
----------------
* Team standings representation
* Ranking configuration
* Sorting priorities
* Net Run Rate (NRR) support
* Points table configuration
* Tournament ranking metadata
* Logging infrastructure
* Public API definitions

Workbook interaction is intentionally excluded from this part and will be
implemented in later sections.

The design aims to be deterministic, reusable, unit-test friendly, and
independent from Excel-specific implementation details.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from enum import Enum
from typing import (
    Any,
    Callable,
    Dict,
    Iterable,
    List,
    Mapping,
    MutableMapping,
    NamedTuple,
    Optional,
    Sequence,
    Tuple,
)

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Internal Constants
# ---------------------------------------------------------------------------

_DEFAULT_LOGGER_NAME = __name__

_INTERNAL_FLOAT_TOLERANCE = 1e-9

_INTERNAL_DEFAULT_DECIMAL_PLACES = 3

_INTERNAL_EMPTY_STRING = ""

_INTERNAL_UNKNOWN_TEAM = "<Unknown Team>"

# ---------------------------------------------------------------------------
# Public Constants
# ---------------------------------------------------------------------------

DEFAULT_POINTS_FOR_WIN = 2

DEFAULT_POINTS_FOR_TIE = 1

DEFAULT_POINTS_FOR_NO_RESULT = 1

DEFAULT_POINTS_FOR_LOSS = 0

DEFAULT_NRR_DECIMAL_PLACES = 3

DEFAULT_RANK_START = 1

DEFAULT_MAX_DISPLAY_TEAMS = 1000

# ---------------------------------------------------------------------------
# Enumerations
# ---------------------------------------------------------------------------


class MatchResult(Enum):
    """Supported match outcomes."""

    WIN = "WIN"
    LOSS = "LOSS"
    TIE = "TIE"
    NO_RESULT = "NO_RESULT"
    ABANDONED = "ABANDONED"


class RankingCriterion(Enum):
    """Supported leaderboard ranking criteria."""

    POINTS = "points"
    NET_RUN_RATE = "net_run_rate"
    WINS = "wins"
    HEAD_TO_HEAD = "head_to_head"
    FEWEST_LOSSES = "fewest_losses"
    RUNS_FOR = "runs_for"
    RUNS_AGAINST = "runs_against"
    ALPHABETICAL = "alphabetical"


class SortDirection(Enum):
    """Sort direction."""

    ASCENDING = "ascending"
    DESCENDING = "descending"


# ---------------------------------------------------------------------------
# NamedTuple Definitions
# ---------------------------------------------------------------------------


class SortRule(NamedTuple):
    """Defines a single sorting rule."""

    field_name: str
    direction: SortDirection


class RankingRule(NamedTuple):
    """Represents one ranking priority."""

    criterion: RankingCriterion
    direction: SortDirection


# ---------------------------------------------------------------------------
# Dataclasses
# ---------------------------------------------------------------------------


@dataclass(slots=True)
class LeaderboardConfiguration:
    """
    Global leaderboard configuration.
    """

    points_for_win: int = DEFAULT_POINTS_FOR_WIN
    points_for_tie: int = DEFAULT_POINTS_FOR_TIE
    points_for_no_result: int = DEFAULT_POINTS_FOR_NO_RESULT
    points_for_loss: int = DEFAULT_POINTS_FOR_LOSS
    nrr_decimal_places: int = DEFAULT_NRR_DECIMAL_PLACES
    rank_start: int = DEFAULT_RANK_START
    max_display_teams: int = DEFAULT_MAX_DISPLAY_TEAMS


@dataclass(slots=True)
class LeaderboardEntry:
    """
    Represents a single team's standings entry.
    """

    team_name: str

    matches: int = 0

    wins: int = 0

    losses: int = 0

    ties: int = 0

    no_results: int = 0

    points: int = 0

    runs_for: int = 0

    overs_faced: float = 0.0

    runs_against: int = 0

    overs_bowled: float = 0.0

    net_run_rate: float = 0.0

    rank: Optional[int] = None

    metadata: Dict[str, Any] = field(default_factory=dict)


@dataclass(slots=True)
class Leaderboard:
    """
    Container for leaderboard entries.
    """

    entries: List[LeaderboardEntry] = field(default_factory=list)

    configuration: LeaderboardConfiguration = field(
        default_factory=LeaderboardConfiguration
    )


# ---------------------------------------------------------------------------
# Default Sorting Rules
# ---------------------------------------------------------------------------

DEFAULT_SORT_RULES: Tuple[SortRule, ...] = (
    SortRule("points", SortDirection.DESCENDING),
    SortRule("net_run_rate", SortDirection.DESCENDING),
    SortRule("wins", SortDirection.DESCENDING),
    SortRule("team_name", SortDirection.ASCENDING),
)

# ---------------------------------------------------------------------------
# Default Ranking Priorities
# ---------------------------------------------------------------------------

DEFAULT_RANKING_PRIORITIES: Tuple[RankingRule, ...] = (
    RankingRule(
        RankingCriterion.POINTS,
        SortDirection.DESCENDING,
    ),
    RankingRule(
        RankingCriterion.NET_RUN_RATE,
        SortDirection.DESCENDING,
    ),
    RankingRule(
        RankingCriterion.WINS,
        SortDirection.DESCENDING,
    ),
    RankingRule(
        RankingCriterion.ALPHABETICAL,
        SortDirection.ASCENDING,
    ),
)

# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

__all__ = [
    "DEFAULT_MAX_DISPLAY_TEAMS",
    "DEFAULT_NRR_DECIMAL_PLACES",
    "DEFAULT_POINTS_FOR_LOSS",
    "DEFAULT_POINTS_FOR_NO_RESULT",
    "DEFAULT_POINTS_FOR_TIE",
    "DEFAULT_POINTS_FOR_WIN",
    "DEFAULT_RANK_START",
    "DEFAULT_RANKING_PRIORITIES",
    "DEFAULT_SORT_RULES",
    "Leaderboard",
    "LeaderboardConfiguration",
    "LeaderboardEntry",
    "MatchResult",
    "RankingCriterion",
    "RankingRule",
    "SortDirection",
    "SortRule",
]