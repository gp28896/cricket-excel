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


# ---------------------------------------------------------------------------
# Tie Breakers
# ---------------------------------------------------------------------------

from typing import Protocol


# ---------------------------------------------------------------------------
# Tie Break Strategy Protocol
# ---------------------------------------------------------------------------


class TieBreakerStrategy(Protocol):
    """
    Protocol implemented by tournament-specific tie-break strategies.

    A strategy receives only the tied teams and returns them in their
    resolved order.

    Examples
    --------
    ICC Events
    Domestic League
    Franchise League
    School Tournament
    """

    def __call__(
        self,
        teams: Sequence[TeamStanding],
    ) -> List[TeamStanding]:
        ...


# ---------------------------------------------------------------------------
# Head-to-Head Helpers
# ---------------------------------------------------------------------------


def head_to_head_lookup(
    team_a: str,
    team_b: str,
    results: Mapping[Tuple[str, str], MatchResult],
) -> Optional[MatchResult]:
    """
    Lookup the official head-to-head result.

    Parameters
    ----------
    team_a
        First team.

    team_b
        Second team.

    results
        Mapping::

            (winner, loser) -> MatchResult

    Returns
    -------
    MatchResult | None

    Notes
    -----
    None is returned when the teams have never met or when the
    tournament has not supplied head-to-head information.
    """
    if (team_a, team_b) in results:
        return results[(team_a, team_b)]

    if (team_b, team_a) in results:
        result = results[(team_b, team_a)]

        if result is MatchResult.WIN:
            return MatchResult.LOSS

        if result is MatchResult.LOSS:
            return MatchResult.WIN

        return result

    return None


# ---------------------------------------------------------------------------
# Mini League
# ---------------------------------------------------------------------------


def mini_league(
    tied_teams: Sequence[TeamStanding],
    head_to_head_results: Mapping[
        Tuple[str, str],
        MatchResult,
    ],
) -> Dict[str, int]:
    """
    Construct a mini-league table.

    Only matches involving tied teams are considered.

    Returns
    -------
    dict

        {
            "India": 4,
            "Australia": 2,
            "England": 0,
        }

    The returned value represents mini-league points only.
    """
    table = {
        team.team_name: 0
        for team in tied_teams
    }

    names = list(table.keys())

    for i, first in enumerate(names):

        for second in names[i + 1:]:

            result = head_to_head_lookup(
                first,
                second,
                head_to_head_results,
            )

            if result is MatchResult.WIN:
                table[first] += DEFAULT_POINTS_FOR_WIN

            elif result is MatchResult.LOSS:
                table[second] += DEFAULT_POINTS_FOR_WIN

            elif result is MatchResult.TIE:
                table[first] += DEFAULT_POINTS_FOR_TIE
                table[second] += DEFAULT_POINTS_FOR_TIE

            elif result is MatchResult.NO_RESULT:
                table[first] += DEFAULT_POINTS_FOR_NO_RESULT
                table[second] += DEFAULT_POINTS_FOR_NO_RESULT

    return table


# ---------------------------------------------------------------------------
# Two-Team Tie
# ---------------------------------------------------------------------------


def break_two_team_tie(
    team_a: TeamStanding,
    team_b: TeamStanding,
    *,
    head_to_head_results: Optional[
        Mapping[Tuple[str, str], MatchResult]
    ] = None,
    fallback: Optional[
        TieBreakerStrategy
    ] = None,
) -> List[TeamStanding]:
    """
    Resolve a two-team tie.

    Resolution order

    1. Head-to-head
    2. Custom fallback strategy
    3. Alphabetical order
    """
    if head_to_head_results:

        result = head_to_head_lookup(
            team_a.team_name,
            team_b.team_name,
            head_to_head_results,
        )

        if result is MatchResult.WIN:
            return [team_a, team_b]

        if result is MatchResult.LOSS:
            return [team_b, team_a]

    if fallback is not None:
        return fallback([team_a, team_b])

    return sorted(
        [team_a, team_b],
        key=lambda t: t.team_name.casefold(),
    )


# ---------------------------------------------------------------------------
# Multi-Team Tie
# ---------------------------------------------------------------------------


def break_multi_team_tie(
    teams: Sequence[TeamStanding],
    *,
    head_to_head_results: Optional[
        Mapping[Tuple[str, str], MatchResult]
    ] = None,
    strategy: Optional[
        TieBreakerStrategy
    ] = None,
) -> List[TeamStanding]:
    """
    Resolve a tie involving three or more teams.

    Default behaviour

    1. Mini-league
    2. Overall NRR
    3. Overall wins
    4. Alphabetical

    A tournament may completely replace this behaviour by supplying a
    custom strategy.
    """
    teams = list(teams)

    if strategy is not None:
        return strategy(teams)

    if head_to_head_results:

        mini_table = mini_league(
            teams,
            head_to_head_results,
        )

        teams.sort(
            key=lambda team: (
                mini_table.get(team.team_name, 0),
                team.net_run_rate,
                team.results.wins,
                team.team_name.casefold(),
            ),
            reverse=True,
        )

        return teams

    return sorted(
        teams,
        key=lambda team: (
            team.points,
            team.net_run_rate,
            team.results.wins,
            team.team_name.casefold(),
        ),
        reverse=True,
    )


# ---------------------------------------------------------------------------
# Coin Toss Placeholder
# ---------------------------------------------------------------------------


def coin_toss_placeholder(
    teams: Sequence[TeamStanding],
) -> List[TeamStanding]:
    """
    Placeholder for tournaments that ultimately resolve ties
    by drawing lots or performing a coin toss.

    Raises
    ------
    NotImplementedError

    Tournament implementations should replace this function with an
    official resolver appropriate for their playing conditions.
    """
    raise NotImplementedError(
        "Coin toss resolution must be implemented by the "
        "tournament-specific tie-break strategy."
    )


# ---------------------------------------------------------------------------
# Default Tie-Break Registry
# ---------------------------------------------------------------------------

DEFAULT_TIE_BREAK_STRATEGY: TieBreakerStrategy = (
    break_multi_team_tie
)


# ---------------------------------------------------------------------------
# Match Processing
# ---------------------------------------------------------------------------


@dataclass(slots=True)
class MatchSummary:
    """
    Lightweight representation of a completed match.

    This object forms the interface between scorecard.py and
    leaderboard.py.

    scorecard.py is responsible for producing MatchSummary objects.
    leaderboard.py consumes them to update tournament standings.
    """

    team1: str
    team2: str

    team1_runs: int
    team1_overs: float
    team1_all_out: bool = False

    team2_runs: int = 0
    team2_overs: float = 0.0
    team2_all_out: bool = False

    winner: Optional[str] = None

    tied: bool = False

    no_result: bool = False

    abandoned: bool = False

    walkover: bool = False

    forfeit: bool = False

    super_over: bool = False

    scheduled_overs: Optional[float] = None

    dls_adjustment_team1: int = 0
    dls_adjustment_team2: int = 0

    metadata: Dict[str, Any] = field(default_factory=dict)


# ---------------------------------------------------------------------------
# Extraction
# ---------------------------------------------------------------------------


def extract_match_summary(
    scorecard: Mapping[str, Any],
) -> MatchSummary:
    """
    Extract a MatchSummary from scorecard output.

    Parameters
    ----------
    scorecard

        Dictionary produced by scorecard.py.

    Notes
    -----
    This function deliberately understands only the public scorecard API.
    Internal worksheet layout is ignored.
    """

    return MatchSummary(
        team1=scorecard["team1"],
        team2=scorecard["team2"],

        team1_runs=scorecard["team1_runs"],
        team1_overs=scorecard["team1_overs"],
        team1_all_out=scorecard.get(
            "team1_all_out",
            False,
        ),

        team2_runs=scorecard["team2_runs"],
        team2_overs=scorecard["team2_overs"],
        team2_all_out=scorecard.get(
            "team2_all_out",
            False,
        ),

        winner=scorecard.get("winner"),

        tied=scorecard.get("tied", False),

        no_result=scorecard.get(
            "no_result",
            False,
        ),

        abandoned=scorecard.get(
            "abandoned",
            False,
        ),

        walkover=scorecard.get(
            "walkover",
            False,
        ),

        forfeit=scorecard.get(
            "forfeit",
            False,
        ),

        super_over=scorecard.get(
            "super_over",
            False,
        ),

        scheduled_overs=scorecard.get(
            "scheduled_overs",
        ),

        dls_adjustment_team1=scorecard.get(
            "dls_adjustment_team1",
            0,
        ),

        dls_adjustment_team2=scorecard.get(
            "dls_adjustment_team2",
            0,
        ),

        metadata=dict(
            scorecard.get("metadata", {})
        ),
    )


# ---------------------------------------------------------------------------
# Match Processing
# ---------------------------------------------------------------------------


def process_completed_match(
    standings: MutableMapping[str, TeamStanding],
    match: MatchSummary,
    *,
    settings: Optional[
        LeaderboardConfiguration
    ] = None,
) -> None:
    """
    Apply one completed match to tournament standings.

    Parameters
    ----------
    standings

        Dictionary keyed by team name.

    match

        Match summary extracted from scorecard.py.

    Notes
    -----
    Missing teams are automatically created.
    """

    settings = settings or LeaderboardConfiguration()

    team1 = standings.setdefault(
        match.team1,
        TeamStanding(match.team1),
    )

    team2 = standings.setdefault(
        match.team2,
        TeamStanding(match.team2),
    )

    #
    # Determine match results.
    #
    if match.abandoned:

        result1 = MatchResult.ABANDONED
        result2 = MatchResult.ABANDONED

    elif match.no_result:

        result1 = MatchResult.NO_RESULT
        result2 = MatchResult.NO_RESULT

    elif match.tied:

        result1 = MatchResult.TIE
        result2 = MatchResult.TIE

    elif match.winner == match.team1:

        result1 = MatchResult.WIN
        result2 = MatchResult.LOSS

    elif match.winner == match.team2:

        result1 = MatchResult.LOSS
        result2 = MatchResult.WIN

    else:

        raise ValueError(
            "Unable to determine match winner."
        )

    update_team_statistics(
        team1,
        result=result1,
        runs_scored=match.team1_runs,
        overs_faced=match.team1_overs,
        runs_conceded=match.team2_runs,
        overs_bowled=match.team2_overs,
        scheduled_overs=match.scheduled_overs,
        batting_all_out=match.team1_all_out,
        bowling_all_out=match.team2_all_out,
        abandoned=match.abandoned,
        no_result=match.no_result,
        walkover=match.walkover,
        forfeit=match.forfeit,
        super_over=match.super_over,
        settings=settings,
    )

    update_team_statistics(
        team2,
        result=result2,
        runs_scored=match.team2_runs,
        overs_faced=match.team2_overs,
        runs_conceded=match.team1_runs,
        overs_bowled=match.team1_overs,
        scheduled_overs=match.scheduled_overs,
        batting_all_out=match.team2_all_out,
        bowling_all_out=match.team1_all_out,
        abandoned=match.abandoned,
        no_result=match.no_result,
        walkover=match.walkover,
        forfeit=match.forfeit,
        super_over=match.super_over,
        settings=settings,
    )


# ---------------------------------------------------------------------------
# Tournament Processing
# ---------------------------------------------------------------------------


def process_all_matches(
    matches: Iterable[MatchSummary],
    *,
    settings: Optional[
        LeaderboardConfiguration
    ] = None,
) -> TournamentStanding:
    """
    Process every completed match in chronological order.

    Returns
    -------
    TournamentStanding
    """

    standings: Dict[
        str,
        TeamStanding,
    ] = {}

    for match in matches:

        process_completed_match(
            standings,
            match,
            settings=settings,
        )

    ranked = rank_teams(
        list(standings.values())
    )

    return TournamentStanding(
        teams=ranked,
    )


# ---------------------------------------------------------------------------
# Statistics Rebuild
# ---------------------------------------------------------------------------


def rebuild_statistics(
    matches: Iterable[MatchSummary],
    *,
    settings: Optional[
        LeaderboardConfiguration
    ] = None,
) -> TournamentStanding:
    """
    Completely rebuild tournament standings.

    Existing statistics are ignored.

    This function is useful when:

    * a scorecard is edited
    * a match is deleted
    * tournament rules change
    * rankings require recalculation

    The implementation intentionally performs a full rebuild rather than
    incremental repair to guarantee deterministic results.
    """

    return process_all_matches(
        matches,
        settings=settings,
    )


# ---------------------------------------------------------------------------
# Leaderboard Table Builder
# ---------------------------------------------------------------------------

@dataclass(slots=True, frozen=True)
class LeaderboardRow:
    """
    Represents one logical row of the leaderboard.

    This class is intentionally independent of Excel. It contains only the
    information required to render a leaderboard in any presentation layer
    (Excel, HTML, PDF, terminal, etc.).
    """

    rank: Optional[int]

    team_name: str

    matches: int

    wins: int

    losses: int

    ties: int

    no_results: int

    points: int

    net_run_rate: float

    runs_for: int

    runs_against: int

    overs_faced: float

    overs_bowled: float

    metadata: Mapping[str, Any] = field(default_factory=dict)


@dataclass(slots=True, frozen=True)
class SummaryRow:
    """
    Tournament summary row.

    These rows are typically displayed above or below the leaderboard.
    """

    label: str

    value: Any


@dataclass(slots=True, frozen=True)
class FooterRow:
    """
    Footer information displayed beneath the leaderboard.

    Examples
    --------
    • Ranking rules
    • NRR calculation notes
    • Generated timestamp
    """

    label: str

    value: Any


# ---------------------------------------------------------------------------
# Leaderboard Builder
# ---------------------------------------------------------------------------


def build_leaderboard_rows(
    standings: Sequence[TeamStanding],
    *,
    decimal_places: int = DEFAULT_NRR_DECIMAL_PLACES,
) -> List[LeaderboardRow]:
    """
    Convert TeamStanding objects into presentation-neutral leaderboard rows.

    Parameters
    ----------
    standings
        Ranked team standings.

    decimal_places
        Number of decimal places used for displaying Net Run Rate.

    Returns
    -------
    list[LeaderboardRow]
    """

    rows: List[LeaderboardRow] = []

    for team in standings:

        rows.append(
            LeaderboardRow(
                rank=team.rank,
                team_name=team.team_name,
                matches=team.results.played,
                wins=team.results.wins,
                losses=team.results.losses,
                ties=team.results.ties,
                no_results=team.results.no_results,
                points=team.points,
                net_run_rate=round(
                    team.net_run_rate,
                    decimal_places,
                ),
                runs_for=team.nrr.runs_for,
                runs_against=team.nrr.runs_against,
                overs_faced=team.nrr.overs_faced,
                overs_bowled=team.nrr.overs_bowled,
                metadata=dict(team.metadata),
            )
        )

    return rows


# ---------------------------------------------------------------------------
# Summary Builder
# ---------------------------------------------------------------------------


def build_summary_rows(
    standings: Sequence[TeamStanding],
) -> List[SummaryRow]:
    """
    Build tournament summary rows.

    Returns summary information independent of any output format.
    """

    standings = list(standings)

    total_matches = sum(
        team.results.played
        for team in standings
    ) // 2

    total_points = sum(
        team.points
        for team in standings
    )

    highest_points = (
        max(
            (team.points for team in standings),
            default=0,
        )
    )

    best_nrr = (
        max(
            (team.net_run_rate for team in standings),
            default=0.0,
        )
    )

    return [
        SummaryRow(
            "Teams",
            len(standings),
        ),
        SummaryRow(
            "Completed Matches",
            total_matches,
        ),
        SummaryRow(
            "Total Points Awarded",
            total_points,
        ),
        SummaryRow(
            "Highest Points",
            highest_points,
        ),
        SummaryRow(
            "Best Net Run Rate",
            round(
                best_nrr,
                DEFAULT_NRR_DECIMAL_PLACES,
            ),
        ),
    ]


# ---------------------------------------------------------------------------
# Footer Builder
# ---------------------------------------------------------------------------


def build_footer_rows(
    *,
    ranking_rules: Sequence[
        RankingRule
    ] = DEFAULT_RANKING_PRIORITIES,
    include_notes: bool = True,
) -> List[FooterRow]:
    """
    Build footer rows describing leaderboard generation.

    These rows contain informational metadata only and are suitable for
    rendering beneath the standings table.
    """

    rows: List[FooterRow] = []

    ranking_description = " → ".join(
        rule.criterion.value.replace("_", " ").title()
        for rule in ranking_rules
    )

    rows.append(
        FooterRow(
            label="Ranking Priority",
            value=ranking_description,
        )
    )

    if include_notes:

        rows.extend(
            [
                FooterRow(
                    label="Net Run Rate",
                    value=(
                        "Calculated using cumulative runs and "
                        "legal deliveries faced/bowled."
                    ),
                ),
                FooterRow(
                    label="Tie Handling",
                    value=(
                        "Resolved according to configured "
                        "ranking and tie-break rules."
                    ),
                ),
                FooterRow(
                    label="Data Source",
                    value="Generated from processed scorecards.",
                ),
            ]
        )

    return rows


# ---------------------------------------------------------------------------
# Writing Standings to Excel
# ---------------------------------------------------------------------------

from datetime import datetime

try:
    #
    # Preferred architecture:
    #
    # workbook.py exposes a single helper/facade rather than many functions.
    #
    from workbook import WorkbookHelper

except ImportError:  # pragma: no cover
    WorkbookHelper = None

try:
    from styles import (
        apply_header_style,
        apply_table_style,
        apply_footer_style,
        apply_timestamp_style,
    )
except ImportError:  # pragma: no cover

    def apply_header_style(*args, **kwargs):
        return None

    def apply_table_style(*args, **kwargs):
        return None

    def apply_footer_style(*args, **kwargs):
        return None

    def apply_timestamp_style(*args, **kwargs):
        return None


# ---------------------------------------------------------------------------
# Column Definitions
# ---------------------------------------------------------------------------

LEADERBOARD_COLUMNS: Tuple[Tuple[str, str], ...] = (
    ("Rank", "rank"),
    ("Team", "team_name"),
    ("P", "matches"),
    ("W", "wins"),
    ("L", "losses"),
    ("T", "ties"),
    ("NR", "no_results"),
    ("Pts", "points"),
    ("NRR", "net_run_rate"),
)


# ---------------------------------------------------------------------------
# Header Writer
# ---------------------------------------------------------------------------


def write_headers(
    worksheet,
    *,
    start_row: int = 1,
    start_column: int = 1,
) -> int:
    """
    Write leaderboard table headers.

    Parameters
    ----------
    worksheet
        Target worksheet.

    start_row
        Header row.

    start_column
        Starting column.

    Returns
    -------
    int
        Next writable row.
    """

    column = start_column

    for title, _ in LEADERBOARD_COLUMNS:

        cell = worksheet.cell(
            row=start_row,
            column=column,
        )

        cell.value = title

        apply_header_style(cell)

        column += 1

    return start_row + 1


# ---------------------------------------------------------------------------
# Team Row Writer
# ---------------------------------------------------------------------------


def write_team_row(
    worksheet,
    row: LeaderboardRow,
    *,
    row_number: int,
    start_column: int = 1,
) -> int:
    """
    Write one leaderboard row.

    Parameters
    ----------
    worksheet
        Target worksheet.

    row
        LeaderboardRow instance.

    row_number
        Excel row number.

    Returns
    -------
    int
        Next writable row.
    """

    values = (
        row.rank,
        row.team_name,
        row.matches,
        row.wins,
        row.losses,
        row.ties,
        row.no_results,
        row.points,
        row.net_run_rate,
    )

    for offset, value in enumerate(values):

        cell = worksheet.cell(
            row=row_number,
            column=start_column + offset,
        )

        cell.value = value

        apply_table_style(cell)

    return row_number + 1


# ---------------------------------------------------------------------------
# Footer Writer
# ---------------------------------------------------------------------------


def write_footer(
    worksheet,
    footer_rows: Sequence[FooterRow],
    *,
    start_row: int,
    start_column: int = 1,
) -> int:
    """
    Write informational footer rows.

    Parameters
    ----------
    footer_rows
        FooterRow objects.

    Returns
    -------
    int
        Next writable row.
    """

    row = start_row

    for footer in footer_rows:

        label_cell = worksheet.cell(
            row=row,
            column=start_column,
        )

        value_cell = worksheet.cell(
            row=row,
            column=start_column + 1,
        )

        label_cell.value = footer.label
        value_cell.value = footer.value

        apply_footer_style(label_cell)
        apply_footer_style(value_cell)

        row += 1

    return row


# ---------------------------------------------------------------------------
# Timestamp Writer
# ---------------------------------------------------------------------------


def write_timestamp(
    worksheet,
    *,
    row: int,
    column: int = 1,
    timestamp: Optional[datetime] = None,
) -> None:
    """
    Write workbook generation timestamp.

    Parameters
    ----------
    worksheet
        Destination worksheet.

    row
        Target row.

    column
        Target column.

    timestamp
        Optional timestamp. Defaults to current local time.
    """

    timestamp = timestamp or datetime.now()

    cell = worksheet.cell(
        row=row,
        column=column,
    )

    cell.value = (
        f"Generated: "
        f"{timestamp:%Y-%m-%d %H:%M:%S}"
    )

    apply_timestamp_style(cell)


# ---------------------------------------------------------------------------
# Leaderboard Formatting Helpers
# ---------------------------------------------------------------------------

try:
    from styles import (
        apply_qualified_style,
        apply_eliminated_style,
        apply_position_style,
        apply_alternate_row_style,
    )
except ImportError:  # pragma: no cover

    def apply_qualified_style(*args, **kwargs):
        return None

    def apply_eliminated_style(*args, **kwargs):
        return None

    def apply_position_style(*args, **kwargs):
        return None

    def apply_alternate_row_style(*args, **kwargs):
        return None


def highlight_qualified(
    worksheet,
    *,
    data_start_row: int,
    qualified_count: int,
    data_rows: int,
) -> None:
    """
    Highlight teams that have qualified.

    Parameters
    ----------
    worksheet
        Leaderboard worksheet.

    data_start_row
        First team row.

    qualified_count
        Number of qualifying teams.

    data_rows
        Total number of team rows.
    """

    if qualified_count <= 0:
        return

    last_row = min(
        data_start_row + qualified_count - 1,
        data_start_row + data_rows - 1,
    )

    for row in range(data_start_row, last_row + 1):
        for cell in worksheet[row]:
            apply_qualified_style(cell)


def highlight_eliminated(
    worksheet,
    *,
    data_start_row: int,
    data_rows: int,
    eliminated_count: int,
) -> None:
    """
    Highlight eliminated teams.
    """

    if eliminated_count <= 0:
        return

    first_row = max(
        data_start_row,
        data_start_row + data_rows - eliminated_count,
    )

    last_row = data_start_row + data_rows - 1

    for row in range(first_row, last_row + 1):
        for cell in worksheet[row]:
            apply_eliminated_style(cell)


def highlight_top_positions(
    worksheet,
    *,
    data_start_row: int,
    positions: int = 3,
) -> None:
    """
    Highlight the top leaderboard positions.

    Typical usage:

    * Champion
    * Runner-up
    * Third place
    """

    if positions <= 0:
        return

    for offset in range(positions):

        row = data_start_row + offset

        for cell in worksheet[row]:
            apply_position_style(
                cell,
                position=offset + 1,
            )


def alternate_row_colors(
    worksheet,
    *,
    data_start_row: int,
    data_rows: int,
) -> None:
    """
    Apply alternating row colours (banded rows).
    """

    for index in range(data_rows):

        row = data_start_row + index

        if index % 2 == 1:

            for cell in worksheet[row]:
                apply_alternate_row_style(cell)


def freeze_panes(
    worksheet,
    *,
    header_row: int = 1,
    first_data_column: str = "A",
) -> None:
    """
    Freeze the header row.

    Example
    -------
    Header row = 1

    Freeze point becomes A2.
    """

    worksheet.freeze_panes = (
        f"{first_data_column}{header_row + 1}"
    )


def auto_filter(
    worksheet,
    *,
    header_row: int,
    first_column: int,
    last_column: int,
    last_row: int,
) -> None:
    """
    Enable Excel AutoFilter over the leaderboard table.
    """

    worksheet.auto_filter.ref = (
        f"{worksheet.cell(header_row, first_column).coordinate}:"
        f"{worksheet.cell(last_row, last_column).coordinate}"
    )



# ---------------------------------------------------------------------------
# Tournament Summary
# ---------------------------------------------------------------------------

from dataclasses import dataclass, field
from typing import Any, Dict, Iterable, Mapping, Optional


# ---------------------------------------------------------------------------
# Summary Models
# ---------------------------------------------------------------------------


@dataclass(slots=True, frozen=True)
class TournamentSummary:
    """
    Aggregated tournament statistics.

    This structure is presentation-independent and may later be rendered
    to Excel, HTML, PDF or other output formats.
    """

    matches_played: int

    teams: int

    highest_team_score: int

    lowest_team_score: int

    highest_successful_chase: int

    most_runs: int

    most_wickets: int

    best_net_run_rate: float

    metadata: Dict[str, Any] = field(default_factory=dict)


# ---------------------------------------------------------------------------
# Score Helpers
# ---------------------------------------------------------------------------


def highest_team_score(
    scorecards: Iterable[Mapping[str, Any]],
) -> int:
    """
    Return the highest innings total recorded.
    """

    highest = 0

    for scorecard in scorecards:

        highest = max(
            highest,
            int(scorecard.get("team1_runs", 0)),
            int(scorecard.get("team2_runs", 0)),
        )

    return highest


def lowest_team_score(
    scorecards: Iterable[Mapping[str, Any]],
) -> int:
    """
    Return the lowest completed innings total.

    Zero scores are ignored since they typically represent incomplete or
    unused innings.
    """

    scores = []

    for scorecard in scorecards:

        for key in ("team1_runs", "team2_runs"):

            runs = int(scorecard.get(key, 0))

            if runs > 0:
                scores.append(runs)

    return min(scores, default=0)


def highest_chase(
    scorecards: Iterable[Mapping[str, Any]],
) -> int:
    """
    Return the highest successful run chase.
    """

    highest = 0

    for scorecard in scorecards:

        winner = scorecard.get("winner")

        if winner == scorecard.get("team2"):

            highest = max(
                highest,
                int(scorecard.get("team2_runs", 0)),
            )

        elif winner == scorecard.get("team1"):

            target = int(scorecard.get("team2_runs", 0))

            if target > int(scorecard.get("team1_runs", 0)):
                continue

    return highest


def most_runs(
    player_summaries: Iterable[Mapping[str, Any]],
) -> int:
    """
    Return the highest individual tournament run tally.

    Expected input
    --------------
    Sequence containing player summary dictionaries with a ``runs`` field.
    """

    return max(
        (
            int(player.get("runs", 0))
            for player in player_summaries
        ),
        default=0,
    )


def most_wickets(
    player_summaries: Iterable[Mapping[str, Any]],
) -> int:
    """
    Return the highest individual wicket tally.
    """

    return max(
        (
            int(player.get("wickets", 0))
            for player in player_summaries
        ),
        default=0,
    )


def best_nrr(
    standings: Iterable[TeamStanding],
) -> float:
    """
    Return the best tournament Net Run Rate.
    """

    return max(
        (
            team.net_run_rate
            for team in standings
        ),
        default=0.0,
    )


# ---------------------------------------------------------------------------
# Summary Builder
# ---------------------------------------------------------------------------


def build_tournament_summary(
    scorecards: Iterable[Mapping[str, Any]],
    standings: Iterable[TeamStanding],
    *,
    player_summaries: Optional[
        Iterable[Mapping[str, Any]]
    ] = None,
    metadata: Optional[Mapping[str, Any]] = None,
) -> TournamentSummary:
    """
    Build an aggregate tournament summary.

    Parameters
    ----------
    scorecards
        Match summaries produced by scorecard.py.

    standings
        Final ranked team standings.

    player_summaries
        Optional batting/bowling aggregates.

    metadata
        Additional tournament metadata.

    Returns
    -------
    TournamentSummary
    """

    scorecards = list(scorecards)
    standings = list(standings)

    player_summaries = list(player_summaries or [])

    return TournamentSummary(
        matches_played=len(scorecards),
        teams=len(standings),
        highest_team_score=highest_team_score(
            scorecards
        ),
        lowest_team_score=lowest_team_score(
            scorecards
        ),
        highest_successful_chase=highest_chase(
            scorecards
        ),
        most_runs=most_runs(
            player_summaries
        ),
        most_wickets=most_wickets(
            player_summaries
        ),
        best_net_run_rate=best_nrr(
            standings
        ),
        metadata=dict(metadata or {}),
    )

# ---------------------------------------------------------------------------
# Chart Support
# ---------------------------------------------------------------------------

from dataclasses import dataclass, field


# ---------------------------------------------------------------------------
# Chart Models
# ---------------------------------------------------------------------------


@dataclass(slots=True, frozen=True)
class ChartSeries:
    """
    Represents one data series suitable for chart generation.

    This structure is intentionally independent of Excel. It may later be
    consumed by openpyxl, XlsxWriter, matplotlib, Plotly, or web frontends.
    """

    title: str

    categories: List[Any]

    values: List[float]


@dataclass(slots=True, frozen=True)
class ChartRanges:
    """
    Logical ranges describing where leaderboard chart data resides.

    These are worksheet-independent descriptions. Later workbook integration
    can translate them into Excel cell references.
    """

    category_field: str

    value_fields: Tuple[str, ...]

    first_data_row: int

    last_data_row: int


# ---------------------------------------------------------------------------
# Progress Data Builders
# ---------------------------------------------------------------------------


def points_progress_data(
    standings: Sequence[TeamStanding],
) -> ChartSeries:
    """
    Build chart data representing tournament points.
    """

    return ChartSeries(
        title="Tournament Points",
        categories=[
            team.team_name
            for team in standings
        ],
        values=[
            float(team.points)
            for team in standings
        ],
    )


def nrr_progress_data(
    standings: Sequence[TeamStanding],
) -> ChartSeries:
    """
    Build chart data representing Net Run Rate.
    """

    return ChartSeries(
        title="Net Run Rate",
        categories=[
            team.team_name
            for team in standings
        ],
        values=[
            team.net_run_rate
            for team in standings
        ],
    )


def wins_progress_data(
    standings: Sequence[TeamStanding],
) -> ChartSeries:
    """
    Build chart data representing total wins.
    """

    return ChartSeries(
        title="Wins",
        categories=[
            team.team_name
            for team in standings
        ],
        values=[
            float(team.results.wins)
            for team in standings
        ],
    )


# ---------------------------------------------------------------------------
# Chart Range Builder
# ---------------------------------------------------------------------------


def chart_ranges(
    leaderboard_rows: Sequence[LeaderboardRow],
    *,
    first_data_row: int = 2,
) -> ChartRanges:
    """
    Build logical chart ranges.

    Parameters
    ----------
    leaderboard_rows
        Leaderboard rows that will eventually be written to Excel.

    first_data_row
        First row containing team data.

    Returns
    -------
    ChartRanges

    Notes
    -----
    This function deliberately returns logical field names rather than
    Excel cell references. Workbook/chart modules can later translate
    these into worksheet ranges.
    """

    return ChartRanges(
        category_field="team_name",
        value_fields=(
            "points",
            "net_run_rate",
            "wins",
        ),
        first_data_row=first_data_row,
        last_data_row=(
            first_data_row
            + len(leaderboard_rows)
            - 1
        ),
    )


# ---------------------------------------------------------------------------
# Validation & Consistency
# ---------------------------------------------------------------------------

from dataclasses import dataclass, field


# ---------------------------------------------------------------------------
# Validation Models
# ---------------------------------------------------------------------------


@dataclass(slots=True, frozen=True)
class ValidationIssue:
    """
    Represents a single leaderboard validation issue.
    """

    severity: str
    team: Optional[str]
    field: str
    message: str


@dataclass(slots=True)
class ValidationReport:
    """
    Validation results produced by leaderboard auditing.
    """

    issues: List[ValidationIssue] = field(default_factory=list)

    @property
    def has_errors(self) -> bool:
        return any(
            issue.severity == "ERROR"
            for issue in self.issues
        )

    @property
    def has_warnings(self) -> bool:
        return any(
            issue.severity == "WARNING"
            for issue in self.issues
        )

    def add_error(
        self,
        team: Optional[str],
        field: str,
        message: str,
    ) -> None:
        self.issues.append(
            ValidationIssue(
                severity="ERROR",
                team=team,
                field=field,
                message=message,
            )
        )

    def add_warning(
        self,
        team: Optional[str],
        field: str,
        message: str,
    ) -> None:
        self.issues.append(
            ValidationIssue(
                severity="WARNING",
                team=team,
                field=field,
                message=message,
            )
        )

    def raise_if_invalid(self) -> None:
        """
        Raise a ValueError summarizing all validation errors.
        """

        if not self.has_errors:
            return

        lines = [
            f"[{issue.team or 'Tournament'}] "
            f"{issue.field}: {issue.message}"
            for issue in self.issues
            if issue.severity == "ERROR"
        ]

        raise ValueError(
            "Leaderboard validation failed:\n"
            + "\n".join(lines)
        )


# ---------------------------------------------------------------------------
# Points Validation
# ---------------------------------------------------------------------------


def validate_points(
    standing: TeamStanding,
    *,
    settings: Optional[
        LeaderboardConfiguration
    ] = None,
) -> None:
    """
    Validate tournament points.

    Raises
    ------
    ValueError
        If calculated points differ from stored points.
    """

    settings = settings or LeaderboardConfiguration()

    expected = (
        standing.results.wins * settings.points_for_win
        + standing.results.losses * settings.points_for_loss
        + standing.results.ties * settings.points_for_tie
        + standing.results.no_results
        * settings.points_for_no_result
    )

    expected += getattr(
        standing,
        "bonus_points",
        0,
    )

    expected -= getattr(
        standing,
        "penalty_points",
        0,
    )

    if standing.points != expected:
        raise ValueError(
            f"{standing.team_name}: "
            f"expected {expected} points, "
            f"found {standing.points}."
        )


# ---------------------------------------------------------------------------
# Match Validation
# ---------------------------------------------------------------------------


def validate_matches(
    standing: TeamStanding,
) -> None:
    """
    Validate match totals.
    """

    calculated = (
        standing.results.wins
        + standing.results.losses
        + standing.results.ties
        + standing.results.no_results
    )

    if calculated != standing.results.played:
        raise ValueError(
            f"{standing.team_name}: "
            "matches played do not equal the sum of "
            "wins, losses, ties and no results."
        )


# ---------------------------------------------------------------------------
# Net Run Rate Validation
# ---------------------------------------------------------------------------


def validate_nrr(
    standing: TeamStanding,
    *,
    tolerance: float = _INTERNAL_FLOAT_TOLERANCE,
) -> None:
    """
    Validate stored Net Run Rate.

    Raises
    ------
    ValueError
        If the stored value differs from the calculated value.
    """

    calculated = calculate_net_run_rate(
        runs_scored=standing.nrr.runs_for,
        overs_faced=standing.nrr.overs_faced,
        runs_conceded=standing.nrr.runs_against,
        overs_bowled=standing.nrr.overs_bowled,
    )

    if (
        abs(
            calculated
            - standing.net_run_rate
        )
        > tolerance
    ):
        raise ValueError(
            f"{standing.team_name}: "
            "Net Run Rate validation failed "
            f"(expected {calculated:.6f}, "
            f"found {standing.net_run_rate:.6f})."
        )


# ---------------------------------------------------------------------------
# Standings Validation
# ---------------------------------------------------------------------------


def validate_standings(
    standings: Sequence[TeamStanding],
    *,
    settings: Optional[
        LeaderboardConfiguration
    ] = None,
) -> None:
    """
    Validate every team in the standings.

    Raises
    ------
    ValueError
        If any inconsistency is detected.
    """

    names = set()

    for standing in standings:

        if standing.team_name in names:
            raise ValueError(
                f"Duplicate team detected: "
                f"{standing.team_name}"
            )

        names.add(standing.team_name)

        validate_matches(
            standing,
        )

        validate_points(
            standing,
            settings=settings,
        )

        validate_nrr(
            standing,
        )


# ---------------------------------------------------------------------------
# Leaderboard Audit
# ---------------------------------------------------------------------------


def audit_leaderboard(
    standings: Sequence[TeamStanding],
    *,
    settings: Optional[
        LeaderboardConfiguration
    ] = None,
) -> ValidationReport:
    """
    Perform a complete leaderboard audit.

    Unlike the validation helpers, this function continues collecting
    issues instead of stopping at the first error.

    Returns
    -------
    ValidationReport
    """

    report = ValidationReport()

    seen = set()

    for standing in standings:

        if standing.team_name in seen:
            report.add_error(
                standing.team_name,
                "team_name",
                "Duplicate team detected.",
            )

        seen.add(
            standing.team_name,
        )

        try:
            validate_matches(
                standing,
            )
        except ValueError as exc:
            report.add_error(
                standing.team_name,
                "matches",
                str(exc),
            )

        try:
            validate_points(
                standing,
                settings=settings,
            )
        except ValueError as exc:
            report.add_error(
                standing.team_name,
                "points",
                str(exc),
            )

        try:
            validate_nrr(
                standing,
            )
        except ValueError as exc:
            report.add_error(
                standing.team_name,
                "net_run_rate",
                str(exc),
            )

    return report


# ---------------------------------------------------------------------------
# High-Level Public API
# ---------------------------------------------------------------------------

logger = logging.getLogger(__name__)


def generate_leaderboard(
    workbook,
    scorecards: Iterable[MatchSummary],
    *,
    sheet_name: str = DEFAULT_LEADERBOARD_SHEET_NAME,
    settings: Optional[LeaderboardConfiguration] = None,
    include_summary: bool = True,
    include_footer: bool = True,
    apply_formatting: bool = True,
) -> TournamentStanding:
    """
    Generate a complete leaderboard worksheet.

    This is the primary public entry point for leaderboard generation.

    Workflow
    --------
    1. Process all completed matches.
    2. Validate standings.
    3. Prepare worksheet.
    4. Write leaderboard.
    5. Optionally write summary.
    6. Optionally write footer.
    7. Apply formatting.
    8. Write generation timestamp.

    Parameters
    ----------
    workbook
        Open workbook instance.

    scorecards
        Iterable of MatchSummary objects.

    Returns
    -------
    TournamentStanding
    """

    logger.info("Generating leaderboard.")

    tournament = process_all_matches(
        scorecards,
        settings=settings,
    )

    validate_standings(
        tournament.teams,
        settings=settings,
    )

    worksheet = prepare_leaderboard(
        workbook,
        sheet_name=sheet_name,
        recreate=True,
    )

    leaderboard_rows = build_leaderboard_rows(
        tournament.teams,
    )

    row = write_headers(worksheet)

    for leaderboard_row in leaderboard_rows:
        row = write_team_row(
            worksheet,
            leaderboard_row,
            row_number=row,
        )

    if include_summary:
        summary = build_summary_rows(
            tournament.teams,
        )

        row += 1

        for summary_row in summary:
            worksheet.cell(row=row, column=1).value = summary_row.label
            worksheet.cell(row=row, column=2).value = summary_row.value
            row += 1

    if include_footer:
        footer = build_footer_rows()

        row += 1

        row = write_footer(
            worksheet,
            footer,
            start_row=row,
        )

    write_timestamp(
        worksheet,
        row=row + 1,
    )

    if apply_formatting:

        data_start = 2
        data_rows = len(leaderboard_rows)

        alternate_row_colors(
            worksheet,
            data_start_row=data_start,
            data_rows=data_rows,
        )

        freeze_panes(
            worksheet,
        )

        auto_filter(
            worksheet,
            header_row=1,
            first_column=1,
            last_column=len(LEADERBOARD_COLUMNS),
            last_row=data_rows + 1,
        )

    logger.info("Leaderboard generation complete.")

    return tournament


def update_leaderboard(
    workbook,
    scorecards: Iterable[MatchSummary],
    *,
    sheet_name: str = DEFAULT_LEADERBOARD_SHEET_NAME,
    settings: Optional[LeaderboardConfiguration] = None,
) -> TournamentStanding:
    """
    Update an existing leaderboard.

    This is currently implemented as a complete regeneration to ensure
    deterministic results.
    """

    logger.info("Updating leaderboard.")

    return generate_leaderboard(
        workbook,
        scorecards,
        sheet_name=sheet_name,
        settings=settings,
    )


def refresh_leaderboard(
    workbook,
    scorecards: Iterable[MatchSummary],
    *,
    sheet_name: str = DEFAULT_LEADERBOARD_SHEET_NAME,
    settings: Optional[LeaderboardConfiguration] = None,
) -> TournamentStanding:
    """
    Refresh leaderboard after scorecard modifications.
    """

    logger.info("Refreshing leaderboard.")

    return generate_leaderboard(
        workbook,
        scorecards,
        sheet_name=sheet_name,
        settings=settings,
    )


def recalculate_leaderboard(
    workbook,
    scorecards: Iterable[MatchSummary],
    *,
    sheet_name: str = DEFAULT_LEADERBOARD_SHEET_NAME,
    settings: Optional[LeaderboardConfiguration] = None,
) -> TournamentStanding:
    """
    Recalculate all standings from raw match summaries.

    Intended for situations where tournament rules change or historical
    matches have been edited.
    """

    logger.info("Recalculating leaderboard.")

    tournament = rebuild_statistics(
        scorecards,
        settings=settings,
    )

    validate_standings(
        tournament.teams,
        settings=settings,
    )

    generate_leaderboard(
        workbook,
        scorecards,
        sheet_name=sheet_name,
        settings=settings,
    )

    return tournament


def load_existing_leaderboard(
    workbook,
    *,
    sheet_name: str = DEFAULT_LEADERBOARD_SHEET_NAME,
):
    """
    Load the existing leaderboard worksheet.

    Returns
    -------
    Worksheet

    Raises
    ------
    ValueError
        If the leaderboard worksheet does not exist.
    """

    worksheet = find_leaderboard_sheet(
        workbook,
        sheet_name=sheet_name,
    )

    if worksheet is None:
        raise ValueError(
            f"Leaderboard worksheet '{sheet_name}' was not found."
        )

    return worksheet

# ---------------------------------------------------------------------------
# Convenience Functions
# ---------------------------------------------------------------------------

try:
    import pandas as pd

    _PANDAS_AVAILABLE = True
except ImportError:  # pragma: no cover
    pd = None
    _PANDAS_AVAILABLE = False


def team_position(
    standings: Sequence[TeamStanding],
    team_name: str,
) -> Optional[int]:
    """
    Return the ranking position of a team.

    Parameters
    ----------
    standings
        Ranked tournament standings.

    team_name
        Team to locate.

    Returns
    -------
    int | None
        Team rank if found, otherwise None.
    """
    for team in standings:
        if team.team_name.casefold() == team_name.casefold():
            return team.rank

    return None


def qualified_teams(
    standings: Sequence[TeamStanding],
    *,
    qualification_slots: int = 4,
) -> List[TeamStanding]:
    """
    Return teams that qualify for the next stage.

    Parameters
    ----------
    qualification_slots
        Number of qualifying positions.
    """
    if qualification_slots < 0:
        raise ValueError(
            "qualification_slots cannot be negative."
        )

    ranked = sorted(
        standings,
        key=lambda team: team.rank
        if team.rank is not None
        else float("inf"),
    )

    return ranked[:qualification_slots]


def eliminated_teams(
    standings: Sequence[TeamStanding],
    *,
    qualification_slots: int = 4,
) -> List[TeamStanding]:
    """
    Return teams that failed to qualify.
    """
    if qualification_slots < 0:
        raise ValueError(
            "qualification_slots cannot be negative."
        )

    ranked = sorted(
        standings,
        key=lambda team: team.rank
        if team.rank is not None
        else float("inf"),
    )

    return ranked[qualification_slots:]


def points_table(
    standings: Sequence[TeamStanding],
) -> List[LeaderboardRow]:
    """
    Build a presentation-independent points table.

    Equivalent to the rows displayed in the leaderboard.
    """
    return build_leaderboard_rows(
        standings,
    )


def export_dataframe(
    standings: Sequence[TeamStanding],
):
    """
    Export standings as a pandas DataFrame.

    Returns
    -------
    pandas.DataFrame

    Raises
    ------
    ImportError
        If pandas is unavailable.
    """
    if not _PANDAS_AVAILABLE:
        raise ImportError(
            "pandas is required for export_dataframe()."
        )

    rows = build_leaderboard_rows(
        standings,
    )

    return pd.DataFrame(
        [
            {
                "Rank": row.rank,
                "Team": row.team_name,
                "Played": row.matches,
                "Wins": row.wins,
                "Losses": row.losses,
                "Ties": row.ties,
                "No Result": row.no_results,
                "Points": row.points,
                "Net Run Rate": row.net_run_rate,
                "Runs For": row.runs_for,
                "Runs Against": row.runs_against,
                "Overs Faced": row.overs_faced,
                "Overs Bowled": row.overs_bowled,
            }
            for row in rows
        ]
    )


def standings_dictionary(
    standings: Sequence[TeamStanding],
) -> Dict[str, Dict[str, Any]]:
    """
    Export standings as a dictionary keyed by team name.

    Returns
    -------
    dict

    Example
    -------
    {
        "India": {
            "rank": 1,
            "points": 12,
            "net_run_rate": 1.425,
            ...
        }
    }
    """
    return {
        team.team_name: {
            "rank": team.rank,
            "played": team.results.played,
            "wins": team.results.wins,
            "losses": team.results.losses,
            "ties": team.results.ties,
            "no_results": team.results.no_results,
            "points": team.points,
            "net_run_rate": team.net_run_rate,
            "runs_for": team.nrr.runs_for,
            "runs_against": team.nrr.runs_against,
            "overs_faced": team.nrr.overs_faced,
            "overs_bowled": team.nrr.overs_bowled,
        }
        for team in standings
    }