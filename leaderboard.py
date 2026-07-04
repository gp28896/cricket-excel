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



# ---------------------------------------------------------------------------
# Core Data Models
# ---------------------------------------------------------------------------


@dataclass(slots=True)
class NetRunRateComponents:
    """
    Raw values used to calculate Net Run Rate (NRR).

    NRR = (Runs For / Overs Faced) - (Runs Against / Overs Bowled)

    Overs are expected to be expressed as legal-over decimal values
    produced by the shared formulas module in later integration.
    """

    runs_for: int = 0
    overs_faced: float = 0.0
    runs_against: int = 0
    overs_bowled: float = 0.0

    def __post_init__(self) -> None:
        if self.runs_for < 0:
            raise ValueError("runs_for cannot be negative.")

        if self.runs_against < 0:
            raise ValueError("runs_against cannot be negative.")

        if self.overs_faced < 0:
            raise ValueError("overs_faced cannot be negative.")

        if self.overs_bowled < 0:
            raise ValueError("overs_bowled cannot be negative.")

    @property
    def batting_run_rate(self) -> float:
        """Return batting run rate."""

        if self.overs_faced <= 0:
            return 0.0

        return self.runs_for / self.overs_faced

    @property
    def bowling_run_rate(self) -> float:
        """Return bowling run rate conceded."""

        if self.overs_bowled <= 0:
            return 0.0

        return self.runs_against / self.overs_bowled

    @property
    def net_run_rate(self) -> float:
        """Return calculated NRR."""

        return self.batting_run_rate - self.bowling_run_rate

    def to_dict(self) -> Dict[str, Any]:
        """Serialize."""

        return {
            "runs_for": self.runs_for,
            "overs_faced": self.overs_faced,
            "runs_against": self.runs_against,
            "overs_bowled": self.overs_bowled,
            "net_run_rate": self.net_run_rate,
        }

    @classmethod
    def from_dict(cls, data: Mapping[str, Any]) -> "NetRunRateComponents":
        """Deserialize."""

        return cls(
            runs_for=int(data.get("runs_for", 0)),
            overs_faced=float(data.get("overs_faced", 0.0)),
            runs_against=int(data.get("runs_against", 0)),
            overs_bowled=float(data.get("overs_bowled", 0.0)),
        )


@dataclass(slots=True)
class MatchResultSummary:
    """
    Summary of match outcomes for one team.
    """

    played: int = 0
    wins: int = 0
    losses: int = 0
    ties: int = 0
    no_results: int = 0

    def __post_init__(self) -> None:
        for value in (
            self.played,
            self.wins,
            self.losses,
            self.ties,
            self.no_results,
        ):
            if value < 0:
                raise ValueError("Match counts cannot be negative.")

        total = (
            self.wins
            + self.losses
            + self.ties
            + self.no_results
        )

        if self.played != total:
            raise ValueError(
                "played must equal wins + losses + ties + no_results."
            )

    @property
    def win_percentage(self) -> float:
        """Return win percentage."""

        if self.played == 0:
            return 0.0

        return (self.wins / self.played) * 100.0

    def to_dict(self) -> Dict[str, Any]:
        return {
            "played": self.played,
            "wins": self.wins,
            "losses": self.losses,
            "ties": self.ties,
            "no_results": self.no_results,
        }

    @classmethod
    def from_dict(cls, data: Mapping[str, Any]) -> "MatchResultSummary":
        return cls(
            played=int(data.get("played", 0)),
            wins=int(data.get("wins", 0)),
            losses=int(data.get("losses", 0)),
            ties=int(data.get("ties", 0)),
            no_results=int(data.get("no_results", 0)),
        )


@dataclass(slots=True)
class TeamStanding:
    """
    Complete standings record for one team.
    """

    team_name: str

    results: MatchResultSummary = field(
        default_factory=MatchResultSummary
    )

    nrr: NetRunRateComponents = field(
        default_factory=NetRunRateComponents
    )

    points: int = 0

    rank: Optional[int] = None

    metadata: Dict[str, Any] = field(default_factory=dict)

    def __post_init__(self) -> None:
        self.team_name = self.team_name.strip()

        if not self.team_name:
            raise ValueError("team_name cannot be empty.")

        if self.points < 0:
            raise ValueError("points cannot be negative.")

    @property
    def net_run_rate(self) -> float:
        """Computed NRR."""

        return self.nrr.net_run_rate

    @property
    def matches_played(self) -> int:
        return self.results.played

    @property
    def wins(self) -> int:
        return self.results.wins

    @property
    def losses(self) -> int:
        return self.results.losses

    @property
    def ties(self) -> int:
        return self.results.ties

    @property
    def no_results(self) -> int:
        return self.results.no_results

    def to_dict(self) -> Dict[str, Any]:
        return {
            "team_name": self.team_name,
            "points": self.points,
            "rank": self.rank,
            "results": self.results.to_dict(),
            "nrr": self.nrr.to_dict(),
            "metadata": dict(self.metadata),
        }

    @classmethod
    def from_dict(cls, data: Mapping[str, Any]) -> "TeamStanding":
        return cls(
            team_name=str(data["team_name"]),
            points=int(data.get("points", 0)),
            rank=data.get("rank"),
            results=MatchResultSummary.from_dict(
                data.get("results", {})
            ),
            nrr=NetRunRateComponents.from_dict(
                data.get("nrr", {})
            ),
            metadata=dict(data.get("metadata", {})),
        )


@dataclass(slots=True)
class TournamentStanding:
    """
    Standings for an entire tournament.
    """

    tournament_name: str = ""

    teams: List[TeamStanding] = field(default_factory=list)

    generated_by: str = ""

    generated_at: Optional[str] = None

    def __post_init__(self) -> None:
        names = [team.team_name for team in self.teams]

        if len(names) != len(set(names)):
            raise ValueError(
                "Duplicate team names are not permitted."
            )

    @property
    def total_teams(self) -> int:
        return len(self.teams)

    @property
    def ranked_teams(self) -> List[TeamStanding]:
        return sorted(
            self.teams,
            key=lambda t: (
                t.rank if t.rank is not None else 999999,
                t.team_name,
            ),
        )

    def to_dict(self) -> Dict[str, Any]:
        return {
            "tournament_name": self.tournament_name,
            "generated_by": self.generated_by,
            "generated_at": self.generated_at,
            "teams": [team.to_dict() for team in self.teams],
        }

    @classmethod
    def from_dict(
        cls,
        data: Mapping[str, Any],
    ) -> "TournamentStanding":
        return cls(
            tournament_name=data.get("tournament_name", ""),
            generated_by=data.get("generated_by", ""),
            generated_at=data.get("generated_at"),
            teams=[
                TeamStanding.from_dict(item)
                for item in data.get("teams", [])
            ],
        )


@dataclass(slots=True)
class LeaderboardSettings:
    """
    Runtime settings controlling leaderboard generation.
    """

    sort_rules: Tuple[SortRule, ...] = DEFAULT_SORT_RULES

    ranking_rules: Tuple[
        RankingRule,
        ...
    ] = DEFAULT_RANKING_PRIORITIES

    decimal_places: int = DEFAULT_NRR_DECIMAL_PLACES

    assign_ranks: bool = True

    include_metadata: bool = False

    def __post_init__(self) -> None:
        if self.decimal_places < 0:
            raise ValueError(
                "decimal_places cannot be negative."
            )


@dataclass(slots=True)
class RankingResult:
    """
    Output produced by leaderboard ranking operations.
    """

    standings: List[TeamStanding] = field(default_factory=list)

    ties_detected: bool = False

    ranking_rules_used: Tuple[
        RankingRule,
        ...
    ] = DEFAULT_RANKING_PRIORITIES

    def __len__(self) -> int:
        return len(self.standings)

    @property
    def is_empty(self) -> bool:
        return not self.standings

    def to_dict(self) -> Dict[str, Any]:
        return {
            "ties_detected": self.ties_detected,
            "ranking_rules_used": [
                {
                    "criterion": rule.criterion.value,
                    "direction": rule.direction.value,
                }
                for rule in self.ranking_rules_used
            ],
            "standings": [
                standing.to_dict()
                for standing in self.standings
            ],
        }

    @classmethod
    def from_dict(
        cls,
        data: Mapping[str, Any],
    ) -> "RankingResult":
        return cls(
            ties_detected=bool(
                data.get("ties_detected", False)
            ),
            standings=[
                TeamStanding.from_dict(item)
                for item in data.get("standings", [])
            ],
        )



# ---------------------------------------------------------------------------
# Internal Statistical Helpers
# ---------------------------------------------------------------------------


def safe_divide(
    numerator: float,
    denominator: float,
    *,
    default: float = 0.0,
) -> float:
    """
    Safely divide two numbers.

    This helper is intentionally tolerant of divide-by-zero situations and is
    used throughout leaderboard calculations to avoid repeated defensive code.

    Parameters
    ----------
    numerator:
        Value being divided.

    denominator:
        Divisor.

    default:
        Value returned when the denominator is effectively zero.

    Returns
    -------
    float
        Division result or ``default``.
    """
    if abs(denominator) <= _INTERNAL_FLOAT_TOLERANCE:
        return default

    return numerator / denominator


def percentage(
    value: float,
    total: float,
    *,
    default: float = 0.0,
) -> float:
    """
    Return the percentage represented by ``value`` of ``total``.

    Examples
    --------
    >>> percentage(8, 10)
    80.0

    >>> percentage(0, 0)
    0.0
    """
    return safe_divide(value * 100.0, total, default=default)


def points_per_game(
    points: int,
    matches: int,
) -> float:
    """
    Calculate average points earned per completed match.

    Parameters
    ----------
    points:
        Total tournament points.

    matches:
        Number of completed matches.

    Returns
    -------
    float
        Average points per match.
    """
    return safe_divide(float(points), float(matches))


def overs_to_balls(overs: float) -> int:
    """
    Convert cricket overs into legal balls.

    The decimal component is interpreted as completed legal deliveries,
    not mathematical decimal fractions.

    Examples
    --------
    20.0 -> 120
    18.3 -> 111
    49.5 -> 299

    Raises
    ------
    ValueError
        If the delivery component is outside 0-5.
    """
    if overs < 0:
        raise ValueError("Overs cannot be negative.")

    whole = int(overs)
    balls = int(round((overs - whole) * 10))

    if not 0 <= balls <= 5:
        raise ValueError(
            "Invalid over notation. Ball component must be between 0 and 5."
        )

    return (whole * 6) + balls


def balls_to_overs(balls: int) -> float:
    """
    Convert legal deliveries into cricket over notation.

    Examples
    --------
    120 -> 20.0
    111 -> 18.3
    299 -> 49.5
    """
    if balls < 0:
        raise ValueError("Ball count cannot be negative.")

    overs = balls // 6
    remaining = balls % 6

    return float(f"{overs}.{remaining}")


def legal_run_rate(
    runs: int,
    legal_balls: int,
) -> float:
    """
    Calculate run rate using legal deliveries.

    Formula
    -------
        Run Rate = Runs / (Balls / 6)

    This implementation avoids inaccuracies associated with treating
    cricket over notation (e.g. 19.4) as a decimal number.
    """
    if legal_balls <= 0:
        return 0.0

    return runs / (legal_balls / 6.0)


def safe_average(
    values: Sequence[float],
    *,
    default: float = 0.0,
) -> float:
    """
    Return the arithmetic mean of a sequence.

    Empty sequences return ``default`` rather than raising an exception.
    """
    if not values:
        return default

    return sum(values) / len(values)


def clamp(
    value: float,
    minimum: float,
    maximum: float,
) -> float:
    """
    Clamp a numeric value into the supplied inclusive range.

    Parameters
    ----------
    value:
        Value to clamp.

    minimum:
        Lowest permitted value.

    maximum:
        Highest permitted value.

    Returns
    -------
    float
        Clamped value.
    """
    if minimum > maximum:
        raise ValueError("minimum cannot exceed maximum.")

    return max(minimum, min(value, maximum))


def compare_float(
    left: float,
    right: float,
    *,
    tolerance: float = _INTERNAL_FLOAT_TOLERANCE,
) -> int:
    """
    Compare two floating-point values using a tolerance.

    Returns
    -------
    int

    -1
        left < right

     0
        Values are effectively equal.

     1
        left > right
    """
    difference = left - right

    if abs(difference) <= tolerance:
        return 0

    return -1 if difference < 0 else 1


# ---------------------------------------------------------------------------
# Internal Statistical Helpers
# ---------------------------------------------------------------------------


def safe_divide(
    numerator: float,
    denominator: float,
    *,
    default: float = 0.0,
) -> float:
    """
    Safely divide two numbers.

    This helper is intentionally tolerant of divide-by-zero situations and is
    used throughout leaderboard calculations to avoid repeated defensive code.

    Parameters
    ----------
    numerator:
        Value being divided.

    denominator:
        Divisor.

    default:
        Value returned when the denominator is effectively zero.

    Returns
    -------
    float
        Division result or ``default``.
    """
    if abs(denominator) <= _INTERNAL_FLOAT_TOLERANCE:
        return default

    return numerator / denominator


def percentage(
    value: float,
    total: float,
    *,
    default: float = 0.0,
) -> float:
    """
    Return the percentage represented by ``value`` of ``total``.

    Examples
    --------
    >>> percentage(8, 10)
    80.0

    >>> percentage(0, 0)
    0.0
    """
    return safe_divide(value * 100.0, total, default=default)


def points_per_game(
    points: int,
    matches: int,
) -> float:
    """
    Calculate average points earned per completed match.

    Parameters
    ----------
    points:
        Total tournament points.

    matches:
        Number of completed matches.

    Returns
    -------
    float
        Average points per match.
    """
    return safe_divide(float(points), float(matches))


def overs_to_balls(overs: float) -> int:
    """
    Convert cricket overs into legal balls.

    The decimal component is interpreted as completed legal deliveries,
    not mathematical decimal fractions.

    Examples
    --------
    20.0 -> 120
    18.3 -> 111
    49.5 -> 299

    Raises
    ------
    ValueError
        If the delivery component is outside 0-5.
    """
    if overs < 0:
        raise ValueError("Overs cannot be negative.")

    whole = int(overs)
    balls = int(round((overs - whole) * 10))

    if not 0 <= balls <= 5:
        raise ValueError(
            "Invalid over notation. Ball component must be between 0 and 5."
        )

    return (whole * 6) + balls


def balls_to_overs(balls: int) -> float:
    """
    Convert legal deliveries into cricket over notation.

    Examples
    --------
    120 -> 20.0
    111 -> 18.3
    299 -> 49.5
    """
    if balls < 0:
        raise ValueError("Ball count cannot be negative.")

    overs = balls // 6
    remaining = balls % 6

    return float(f"{overs}.{remaining}")


def legal_run_rate(
    runs: int,
    legal_balls: int,
) -> float:
    """
    Calculate run rate using legal deliveries.

    Formula
    -------
        Run Rate = Runs / (Balls / 6)

    This implementation avoids inaccuracies associated with treating
    cricket over notation (e.g. 19.4) as a decimal number.
    """
    if legal_balls <= 0:
        return 0.0

    return runs / (legal_balls / 6.0)


def safe_average(
    values: Sequence[float],
    *,
    default: float = 0.0,
) -> float:
    """
    Return the arithmetic mean of a sequence.

    Empty sequences return ``default`` rather than raising an exception.
    """
    if not values:
        return default

    return sum(values) / len(values)


def clamp(
    value: float,
    minimum: float,
    maximum: float,
) -> float:
    """
    Clamp a numeric value into the supplied inclusive range.

    Parameters
    ----------
    value:
        Value to clamp.

    minimum:
        Lowest permitted value.

    maximum:
        Highest permitted value.

    Returns
    -------
    float
        Clamped value.
    """
    if minimum > maximum:
        raise ValueError("minimum cannot exceed maximum.")

    return max(minimum, min(value, maximum))


def compare_float(
    left: float,
    right: float,
    *,
    tolerance: float = _INTERNAL_FLOAT_TOLERANCE,
) -> int:
    """
    Compare two floating-point values using a tolerance.

    Returns
    -------
    int

    -1
        left < right

     0
        Values are effectively equal.

     1
        left > right
    """
    difference = left - right

    if abs(difference) <= tolerance:
        return 0

    return -1 if difference < 0 else 1


# ---------------------------------------------------------------------------
# Tournament Statistics Engine
# ---------------------------------------------------------------------------


def calculate_points(
    result: MatchResult,
    *,
    settings: Optional[LeaderboardConfiguration] = None,
) -> int:
    """
    Calculate tournament points awarded for a match result.

    Parameters
    ----------
    result
        Official match outcome.

    settings
        Optional leaderboard configuration containing the tournament's
        points system. If omitted, default ICC-style values are used.

    Returns
    -------
    int
        Points awarded.
    """
    settings = settings or LeaderboardConfiguration()

    if result is MatchResult.WIN:
        return settings.points_for_win

    if result is MatchResult.LOSS:
        return settings.points_for_loss

    if result is MatchResult.TIE:
        return settings.points_for_tie

    if result in (
        MatchResult.NO_RESULT,
        MatchResult.ABANDONED,
    ):
        return settings.points_for_no_result

    raise ValueError(f"Unsupported match result: {result}")


def calculate_bonus_points(
    *,
    batting_bonus: int = 0,
    bowling_bonus: int = 0,
    fair_play_bonus: int = 0,
) -> int:
    """
    Calculate bonus points.

    Many domestic tournaments award additional points for batting,
    bowling or fair-play achievements.

    ICC tournaments typically return zero.
    """
    for value in (
        batting_bonus,
        bowling_bonus,
        fair_play_bonus,
    ):
        if value < 0:
            raise ValueError("Bonus points cannot be negative.")

    return (
        batting_bonus
        + bowling_bonus
        + fair_play_bonus
    )


def calculate_penalty_points(
    *,
    over_rate_penalty: int = 0,
    disciplinary_penalty: int = 0,
    administrative_penalty: int = 0,
) -> int:
    """
    Calculate points deducted from a team.

    Returned value is always non-negative.
    """
    for value in (
        over_rate_penalty,
        disciplinary_penalty,
        administrative_penalty,
    ):
        if value < 0:
            raise ValueError("Penalty points cannot be negative.")

    return (
        over_rate_penalty
        + disciplinary_penalty
        + administrative_penalty
    )


def calculate_matches_played(
    *,
    wins: int = 0,
    losses: int = 0,
    ties: int = 0,
    no_results: int = 0,
    walkovers: int = 0,
    forfeits: int = 0,
) -> int:
    """
    Calculate the number of completed fixtures.

    Walkovers and forfeits count as official fixtures for standings.
    """
    values = (
        wins,
        losses,
        ties,
        no_results,
        walkovers,
        forfeits,
    )

    if any(v < 0 for v in values):
        raise ValueError("Match counts cannot be negative.")

    return sum(values)


def apply_match_result(
    standing: TeamStanding,
    *,
    result: MatchResult,
    settings: Optional[LeaderboardConfiguration] = None,
    walkover: bool = False,
    forfeit: bool = False,
    super_over: bool = False,
    bonus_points: int = 0,
    penalty_points: int = 0,
) -> TeamStanding:
    """
    Apply a single match result to a team's standing.

    Notes
    -----
    * Super Overs are considered a WIN/LOSS once resolved.
    * Walkovers are treated as official wins/losses.
    * Forfeits are treated as official wins/losses.
    * Bonus and penalty points are applied after match points.
    """
    settings = settings or LeaderboardConfiguration()

    if result is MatchResult.WIN:
        standing.results.wins += 1

    elif result is MatchResult.LOSS:
        standing.results.losses += 1

    elif result is MatchResult.TIE:
        standing.results.ties += 1

    elif result in (
        MatchResult.NO_RESULT,
        MatchResult.ABANDONED,
    ):
        standing.results.no_results += 1

    else:
        raise ValueError(f"Unsupported result: {result}")

    # Metadata flags retained for future reporting.
    if walkover:
        standing.metadata["walkover"] = (
            standing.metadata.get("walkover", 0) + 1
        )

    if forfeit:
        standing.metadata["forfeit"] = (
            standing.metadata.get("forfeit", 0) + 1
        )

    if super_over:
        standing.metadata["super_over"] = (
            standing.metadata.get("super_over", 0) + 1
        )

    standing.results.played = calculate_matches_played(
        wins=standing.results.wins,
        losses=standing.results.losses,
        ties=standing.results.ties,
        no_results=standing.results.no_results,
        walkovers=standing.metadata.get("walkover", 0),
        forfeits=standing.metadata.get("forfeit", 0),
    )

    standing.points += calculate_points(
        result,
        settings=settings,
    )

    standing.points += bonus_points

    standing.points -= penalty_points

    return standing


def update_team_statistics(
    standing: TeamStanding,
    *,
    result: MatchResult,
    runs_scored: int,
    overs_faced: float,
    runs_conceded: int,
    overs_bowled: float,
    scheduled_overs: Optional[float] = None,
    batting_all_out: bool = False,
    bowling_all_out: bool = False,
    abandoned: bool = False,
    no_result: bool = False,
    walkover: bool = False,
    forfeit: bool = False,
    super_over: bool = False,
    bonus_points: int = 0,
    penalty_points: int = 0,
    settings: Optional[LeaderboardConfiguration] = None,
) -> TeamStanding:
    """
    Update all tournament statistics for a team after one fixture.

    Responsibilities
    ----------------
    * Apply official match result.
    * Update cumulative runs and overs.
    * Recalculate Net Run Rate.
    * Apply bonus and penalty points.
    * Update matches played.

    Workbook integration is intentionally excluded.
    """
    apply_match_result(
        standing,
        result=result,
        settings=settings,
        walkover=walkover,
        forfeit=forfeit,
        super_over=super_over,
        bonus_points=calculate_bonus_points(
            batting_bonus=bonus_points
        ),
        penalty_points=calculate_penalty_points(
            administrative_penalty=penalty_points
        ),
    )

    if not (abandoned or no_result):
        standing.nrr.runs_for += runs_scored
        standing.nrr.runs_against += runs_conceded

        standing.nrr.overs_faced += calculate_overs_faced(
            overs_faced,
            all_out=batting_all_out,
            scheduled_overs=scheduled_overs,
        )

        standing.nrr.overs_bowled += calculate_overs_bowled(
            overs_bowled,
            opposition_all_out=bowling_all_out,
            scheduled_overs=scheduled_overs,
        )

        standing.net_run_rate = standing.nrr.net_run_rate

    return standing


# ---------------------------------------------------------------------------
# Ranking Engine
# ---------------------------------------------------------------------------

from functools import cmp_to_key


def sort_by_points(
    standings: Sequence[TeamStanding],
    *,
    descending: bool = True,
) -> List[TeamStanding]:
    """
    Return standings sorted by tournament points.
    """
    return sorted(
        standings,
        key=lambda team: team.points,
        reverse=descending,
    )


def sort_by_nrr(
    standings: Sequence[TeamStanding],
    *,
    descending: bool = True,
) -> List[TeamStanding]:
    """
    Return standings sorted by Net Run Rate.
    """
    return sorted(
        standings,
        key=lambda team: team.net_run_rate,
        reverse=descending,
    )


def sort_by_runs(
    standings: Sequence[TeamStanding],
    *,
    descending: bool = True,
) -> List[TeamStanding]:
    """
    Sort by cumulative runs scored.
    """
    return sorted(
        standings,
        key=lambda team: team.nrr.runs_for,
        reverse=descending,
    )


def sort_by_wins(
    standings: Sequence[TeamStanding],
    *,
    descending: bool = True,
) -> List[TeamStanding]:
    """
    Sort by number of wins.
    """
    return sorted(
        standings,
        key=lambda team: team.results.wins,
        reverse=descending,
    )


def sort_by_alphabetical(
    standings: Sequence[TeamStanding],
) -> List[TeamStanding]:
    """
    Sort alphabetically by team name.
    """
    return sorted(
        standings,
        key=lambda team: team.team_name.casefold(),
    )


def sort_by_head_to_head(
    standings: Sequence[TeamStanding],
    head_to_head: Optional[
        Mapping[Tuple[str, str], int]
    ] = None,
) -> List[TeamStanding]:
    """
    Sort using head-to-head results.

    Parameters
    ----------
    standings
        Teams to sort.

    head_to_head
        Mapping::

            (winner, loser) -> 1

        Example
        -------
        {
            ("India", "Australia"): 1,
            ("England", "Pakistan"): 1,
        }

    Notes
    -----
    Head-to-head only applies when teams are otherwise tied.
    When no head-to-head data is available the original ordering is
    preserved.
    """
    if not head_to_head:
        return list(standings)

    teams = list(standings)

    def compare(
        left: TeamStanding,
        right: TeamStanding,
    ) -> int:

        if left.team_name == right.team_name:
            return 0

        if head_to_head.get(
            (left.team_name, right.team_name)
        ):
            return -1

        if head_to_head.get(
            (right.team_name, left.team_name)
        ):
            return 1

        return 0

    return sorted(
        teams,
        key=cmp_to_key(compare),
    )


# ---------------------------------------------------------------------------
# Internal Ranking Keys
# ---------------------------------------------------------------------------

_RANKING_KEY_FUNCTIONS: Dict[
    RankingCriterion,
    Callable[[TeamStanding], Any],
] = {
    RankingCriterion.POINTS:
        lambda t: t.points,

    RankingCriterion.NET_RUN_RATE:
        lambda t: t.net_run_rate,

    RankingCriterion.WINS:
        lambda t: t.results.wins,

    RankingCriterion.RUNS_FOR:
        lambda t: t.nrr.runs_for,

    RankingCriterion.RUNS_AGAINST:
        lambda t: t.nrr.runs_against,

    RankingCriterion.ALPHABETICAL:
        lambda t: t.team_name.casefold(),
}


def rank_teams(
    standings: Sequence[TeamStanding],
    *,
    ranking_rules: Optional[
        Sequence[RankingRule]
    ] = None,
    head_to_head: Optional[
        Mapping[Tuple[str, str], int]
    ] = None,
    dense_ranking: bool = True,
) -> List[TeamStanding]:
    """
    Rank teams according to configurable priority rules.

    Parameters
    ----------
    standings
        Teams to rank.

    ranking_rules
        Ordered ranking priorities.

    head_to_head
        Optional head-to-head lookup used when the corresponding
        ranking criterion is encountered.

    dense_ranking
        If True:

            1,2,2,3

        otherwise:

            1,2,2,4

    Returns
    -------
    list[TeamStanding]
        Ranked standings.

    Examples
    --------
    Default ordering::

        Points
        ↓
        NRR
        ↓
        Wins
        ↓
        Alphabetical

    Any tournament may override this simply by supplying a different
    RankingRule sequence.
    """
    rules = tuple(
        ranking_rules or DEFAULT_RANKING_PRIORITIES
    )

    ranked = list(standings)

    #
    # Stable multi-pass sorting.
    #
    # Python's sort is stable, therefore we sort from the
    # lowest priority to the highest priority.
    #
    for rule in reversed(rules):

        if (
            rule.criterion
            is RankingCriterion.HEAD_TO_HEAD
        ):
            ranked = sort_by_head_to_head(
                ranked,
                head_to_head=head_to_head,
            )
            continue

        key = _RANKING_KEY_FUNCTIONS.get(
            rule.criterion
        )

        if key is None:
            continue

        ranked.sort(
            key=key,
            reverse=(
                rule.direction
                is SortDirection.DESCENDING
            ),
        )

    #
    # Assign ranks
    #
    current_rank = 1

    previous_keys = None

    for index, team in enumerate(ranked):

        current_keys = tuple(
            (
                _RANKING_KEY_FUNCTIONS[r.criterion](team)
                if r.criterion
                in _RANKING_KEY_FUNCTIONS
                else None
            )
            for r in rules
            if r.criterion
            is not RankingCriterion.HEAD_TO_HEAD
        )

        if previous_keys is None:
            team.rank = current_rank

        elif current_keys == previous_keys:
            team.rank = current_rank

        else:
            if dense_ranking:
                current_rank += 1
            else:
                current_rank = index + 1

            team.rank = current_rank

        previous_keys = current_keys

    return ranked