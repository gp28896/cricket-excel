"""
===============================================================================
Project : Cricket Excel Tournament Scoring Engine
Module  : formulae.py

Author  : <Your Name>
Version : 0.2.0

Description
-------------------------------------------------------------------------------

This module contains ONLY pure cricket calculation functions.

It intentionally has **NO dependency** on:

    • openpyxl
    • Workbook
    • Worksheet
    • Excel formulas
    • File system
    • GUI
    • Database
    • Web framework

The purpose of this module is to provide deterministic cricket calculations
that can be reused by every other component of the project.

The module never reads from or writes to Excel.

Instead, higher-level modules (score_ball.py, scorecard.py,
leaderboard.py, workbook.py, etc.) call these functions and then write
their returned values into the workbook.

Design Goals
-------------------------------------------------------------------------------

✓ Pure functions
✓ No side effects
✓ Easy unit testing
✓ Deterministic behaviour
✓ Reusable by future desktop, mobile and web applications
✓ Independent of Excel implementation

Internal Design
-------------------------------------------------------------------------------

Cricket overs are **NOT decimal numbers**.

For example,

    19.5

does NOT mean

    nineteen point five overs

It means

    19 overs
    +
    5 legal deliveries

Internally, almost every cricket calculation should therefore be performed
using **legal deliveries (balls)** rather than decimal overs.

Example

Instead of

    18.4 + 1.2

convert to

    112 balls + 8 balls

perform the arithmetic

then convert back into cricket notation.

This approach eliminates floating-point errors and keeps calculations
consistent throughout the scoring engine.

Future Sections
-------------------------------------------------------------------------------

Part 1
    ✓ Constants
    ✓ Type aliases
    ✓ Validation helpers
    ✓ Conversion helpers

Part 2
    • Remaining balls
    • Remaining overs
    • Run rate
    • Required run rate

Part 3
    • Batting calculations

Part 4
    • Bowling calculations

Part 5
    • Match calculations

===============================================================================
"""

from __future__ import annotations

###############################################################################
# Imports
###############################################################################

from math import isclose
from typing import Final
from typing import Literal
from typing import NewType
from typing import overload

###############################################################################
# Type Aliases
###############################################################################

# -----------------------------------------------------------------------------
# These aliases improve readability without changing runtime behaviour.
#
# Example
#
#     def calculate_run_rate(
#         runs: Runs,
#         balls: Balls,
#     ) -> float:
#
# is much clearer than
#
#     def calculate_run_rate(
#         runs: int,
#         balls: int,
#     ) -> float:
# -----------------------------------------------------------------------------

Runs = NewType("Runs", int)
Balls = NewType("Balls", int)
Overs = NewType("Overs", float)
Wickets = NewType("Wickets", int)

###############################################################################
# Match Constants
###############################################################################

#: Number of legal deliveries in one completed over.
BALLS_PER_OVER: Final[int] = 6

#: Maximum wickets that can fall in a completed innings.
MAX_WICKETS: Final[int] = 10

###############################################################################
# Match Status
###############################################################################

MatchStatus = Literal[
    "Scheduled",
    "Live",
    "Completed",
    "Abandoned",
    "No Result",
]

###############################################################################
# Toss Decision
###############################################################################

TossDecision = Literal[
    "Bat",
    "Bowl",
]

###############################################################################
# Validation Helpers
###############################################################################


def _validate_non_negative(
    value: int | float,
    field_name: str,
) -> None:
    """
    Validate that a numeric value is non-negative.

    Parameters
    ----------
    value
        Numeric value to validate.

    field_name
        Human-readable field name used in error messages.

    Raises
    ------
    ValueError
        If value is negative.

    Examples
    --------
    >>> _validate_non_negative(5, "Runs")

    >>> _validate_non_negative(-1, "Runs")
    Traceback (most recent call last):
        ...
    ValueError: Runs cannot be negative.
    """

    if value < 0:
        raise ValueError(f"{field_name} cannot be negative.")


def _validate_over_notation(
    overs: int | float,
) -> None:
    """
    Validate that a cricket over uses valid notation.

    Cricket notation allows only one decimal digit.

    Valid examples

        0
        0.0
        5.2
        19.5
        50.0

    Invalid examples

        10.6
        12.9
        5.75
        20.25

    Parameters
    ----------
    overs
        Cricket overs.

    Raises
    ------
    ValueError
        If the notation is invalid.
    """

    _validate_non_negative(overs, "Overs")

    whole_overs = int(overs)

    fractional_part = overs - whole_overs

    fractional_balls = round(fractional_part * 10)

    #
    # Reject values having more than one decimal digit.
    #
    # Example
    #
    #     20.25
    #
    # becomes
    #
    #     fractional_part * 10 = 2.5
    #
    # which is invalid.
    #
    if not isclose(
        fractional_part * 10,
        fractional_balls,
        abs_tol=1e-9,
    ):
        raise ValueError(
            "Overs must contain at most one decimal digit."
        )

    if fractional_balls >= BALLS_PER_OVER:
        raise ValueError(
            "Invalid cricket over notation. "
            "Decimal part cannot exceed 5."
        )


###############################################################################
# Conversion Function Overloads
###############################################################################

#
# These overloads improve IDE auto-completion and static type checking.
# Runtime behaviour remains unchanged.
#

@overload
def overs_to_balls(
    overs: int,
) -> Balls:
    ...


@overload
def overs_to_balls(
    overs: float,
) -> Balls:
    ...


@overload
def balls_to_overs(
    balls: Balls,
) -> Overs:
    ...


@overload
def balls_to_overs(
    balls: int,
) -> Overs:
    ...


###############################################################################
# Conversion Functions
###############################################################################


def overs_to_balls(
    overs: int | float,
) -> Balls:
    """
    Convert cricket overs into legal deliveries.

    Cricket overs are not decimal values.

    Example

        10.3

    means

        10 overs
        +
        3 legal deliveries

    and NOT

        10.3 × 6

    Examples
    --------

    >>> overs_to_balls(0)
    0

    >>> overs_to_balls(1)
    6

    >>> overs_to_balls(10.3)
    63

    >>> overs_to_balls(19.5)
    119

    Parameters
    ----------
    overs
        Cricket over notation.

    Returns
    -------
    Balls
        Equivalent legal deliveries.

    Raises
    ------
    ValueError

        If

        • overs is negative

        • notation contains more than one decimal digit

        • decimal portion exceeds five
    """

    _validate_over_notation(overs)

    whole_overs = int(overs)

    balls = round((overs - whole_overs) * 10)

    total_balls = (
        whole_overs * BALLS_PER_OVER
    ) + balls

    return Balls(total_balls)


def balls_to_overs(
    balls: Balls | int,
) -> Overs:
    """
    Convert legal deliveries into cricket over notation.

    Cricket notation is represented as

        completed_overs.remaining_balls

    rather than as a decimal quantity.

    Examples
    --------

    >>> balls_to_overs(0)
    0.0

    >>> balls_to_overs(6)
    1.0

    >>> balls_to_overs(7)
    1.1

    >>> balls_to_overs(63)
    10.3

    >>> balls_to_overs(119)
    19.5

    Parameters
    ----------
    balls
        Number of legal deliveries.

    Returns
    -------
    Overs
        Cricket over notation.

    Raises
    ------
    ValueError
        If balls is negative.

    Notes
    -----
    This function intentionally returns cricket notation rather than
    elapsed decimal overs.

    For example

        19.5

    represents

        19 overs
        +
        5 balls

    not

        19.5 decimal overs.
    """

    _validate_non_negative(balls, "Balls")

    completed_overs = balls // BALLS_PER_OVER

    remaining_balls = balls % BALLS_PER_OVER

    #
    # Construct cricket notation.
    #
    # Example
    #
    #     19 overs
    #     5 balls
    #
    # becomes
    #
    #     19.5
    #
    over_value = completed_overs + (remaining_balls / 10)

    return Overs(over_value)

###############################################################################
# Delivery Constants
###############################################################################

#
# The scoring engine classifies every recorded delivery into one of the
# following categories.
#
# Legal deliveries
# ----------------
#
# These deliveries count towards:
#
#     • Ball number
#     • Over progression
#
# Illegal deliveries
# ------------------
#
# These award runs where applicable but DO NOT count as one of the six legal
# deliveries in an over.
#
# The constants below are intentionally implemented as immutable sets for
# efficient membership testing throughout the scoring engine.
#

#: Delivery types that consume one legal ball.
LEGAL_DELIVERY_TYPES: Final[frozenset[str]] = frozenset({
    "Normal",
    "Bye",
    "Leg Bye",
})

#: Delivery types that do not consume a legal ball.
ILLEGAL_DELIVERY_TYPES: Final[frozenset[str]] = frozenset({
    "Wide",
    "No Ball",
})

#: Every delivery type recognised by the scoring engine.
ALL_DELIVERY_TYPES: Final[frozenset[str]] = (
    LEGAL_DELIVERY_TYPES |
    ILLEGAL_DELIVERY_TYPES
)

###############################################################################
# Delivery Validation Functions
###############################################################################


def is_legal_delivery(
    delivery_type: str,
) -> bool:
    """
    Determine whether a delivery counts as a legal ball.

    Cricket laws distinguish between legal and illegal deliveries.

    Legal deliveries advance the over count.

    Illegal deliveries (Wide and No Ball) award extra runs where applicable
    but do not count as one of the six legal balls in an over.

    Parameters
    ----------
    delivery_type
        Delivery type recorded by the scoring engine.

    Returns
    -------
    bool

        True
            Delivery counts towards the over.

        False
            Delivery does not count towards the over.

    Raises
    ------
    ValueError

        If the supplied delivery type is unknown.

    Examples
    --------

    >>> is_legal_delivery("Normal")
    True

    >>> is_legal_delivery("Bye")
    True

    >>> is_legal_delivery("Leg Bye")
    True

    >>> is_legal_delivery("Wide")
    False

    >>> is_legal_delivery("No Ball")
    False
    """

    if delivery_type not in ALL_DELIVERY_TYPES:
        valid_types = ", ".join(sorted(ALL_DELIVERY_TYPES))

        raise ValueError(
            f"Unknown delivery type '{delivery_type}'. "
            f"Expected one of: {valid_types}."
        )

    return delivery_type in LEGAL_DELIVERY_TYPES


def legal_deliveries(
    delivery_types: list[str],
) -> Balls:
    """
    Count the number of legal deliveries in an ordered delivery sequence.

    This helper is frequently used while rebuilding an innings from an event
    log or validating scorecards.

    Every delivery is first validated using ``is_legal_delivery()`` to ensure
    unknown delivery types are rejected immediately.

    Parameters
    ----------
    delivery_types
        Ordered list of delivery types.

    Returns
    -------
    Balls
        Number of legal deliveries.

    Raises
    ------
    ValueError
        If any delivery type is not recognised.

    Examples
    --------

    >>> legal_deliveries([
    ...     "Normal",
    ...     "Wide",
    ...     "Normal",
    ...     "No Ball",
    ... ])
    2

    >>> legal_deliveries([])
    0

    Notes
    -----
    This function intentionally delegates validation to
    ``is_legal_delivery()`` so that the delivery rules are maintained in one
    location only.
    """

    count = 0

    for delivery in delivery_types:

        if is_legal_delivery(delivery):
            count += 1

    return Balls(count)


###############################################################################
# Module Exports
###############################################################################

__all__ = [
    # Type aliases
    "Runs",
    "Balls",
    "Overs",
    "Wickets",

    # Match constants
    "BALLS_PER_OVER",
    "MAX_WICKETS",

    # Match metadata
    "MatchStatus",
    "TossDecision",

    # Delivery constants
    "LEGAL_DELIVERY_TYPES",
    "ILLEGAL_DELIVERY_TYPES",
    "ALL_DELIVERY_TYPES",

    # Conversion helpers
    "overs_to_balls",
    "balls_to_overs",

    # Delivery helpers
    "is_legal_delivery",
    "legal_deliveries",

    # Innings helpers
    "remaining_balls",
    "remaining_overs",

    # Run rate calculations     
    "calculate_run_rate",
    "calculate_required_run_rate",
]

###############################################################################
# Internal Helper Functions
###############################################################################


def _validate_positive_balls(
    balls: int,
    field_name: str,
) -> None:
    """
    Validate that a ball count is greater than zero.

    This helper is used by rate calculations where division by zero would
    otherwise occur.

    Parameters
    ----------
    balls
        Number of legal deliveries.

    field_name
        Human-readable field name.

    Raises
    ------
    ValueError
        If the supplied value is zero or negative.

    Examples
    --------
    >>> _validate_positive_balls(1, "Balls")
    >>> _validate_positive_balls(0, "Balls")
    Traceback (most recent call last):
        ...
    ValueError: Balls must be greater than zero.
    """

    if balls <= 0:
        raise ValueError(
            f"{field_name} must be greater than zero."
        )


###############################################################################
# Run Rate Calculations
###############################################################################


def calculate_run_rate(
    runs: Runs | int,
    overs_faced: Overs | int | float,
    *,
    precision: int = 2,
) -> float:
    """
    Calculate a team's run rate.

    Unlike many implementations, this function NEVER performs arithmetic
    directly on cricket over notation.

    Instead,

        overs
            ↓
        legal deliveries
            ↓
        run rate

    Formula
    -------

        Run Rate = Runs × 6 / Legal Deliveries

    Parameters
    ----------
    runs
        Runs scored.

    overs_faced
        Overs faced using cricket notation.

    precision
        Decimal places to round to.

    Returns
    -------
    float
        Run rate.

    Raises
    ------
    ValueError

        If

        * runs is negative

        * over notation is invalid

        * zero legal deliveries have been bowled

    Examples
    --------

    >>> calculate_run_rate(150, 20)
    7.5

    >>> calculate_run_rate(150, 18.3)
    8.11

    >>> calculate_run_rate(84, 10.4)
    7.88

    Notes
    -----
    Cricket notation such as

        18.3

    does NOT represent eighteen point three overs.

    It represents

        18 overs
        +
        3 legal deliveries

    Therefore calculations are always performed using legal deliveries.
    """

    _validate_non_negative(runs, "Runs")

    legal_balls = overs_to_balls(overs_faced)

    _validate_positive_balls(
        legal_balls,
        "Overs faced",
    )

    run_rate = (runs * BALLS_PER_OVER) / legal_balls

    return round(run_rate, precision)


def calculate_required_run_rate(
    target_runs: Runs | int,
    current_runs: Runs | int,
    total_overs: Overs | int | float,
    overs_bowled: Overs | int | float,
    *,
    precision: int = 2,
) -> float:
    """
    Calculate the required run rate.

    Formula
    -------

        Runs Required × 6
        -----------------
        Balls Remaining

    Runs Required is calculated according to cricket's chase rule:

        Target - Current Score

    Parameters
    ----------
    target_runs
        Target score to win.

    current_runs
        Current batting score.

    total_overs
        Scheduled innings length.

    overs_bowled
        Overs already faced.

    precision
        Decimal places.

    Returns
    -------
    float
        Required run rate.

    Raises
    ------
    ValueError

        If

        * target is negative

        * current score is negative

        * innings has no remaining deliveries

        * over notation is invalid

    Examples
    --------

    Twenty-over chase

    >>> calculate_required_run_rate(
    ...     target_runs=181,
    ...     current_runs=120,
    ...     total_overs=20,
    ...     overs_bowled=15,
    ... )
    12.2

    Fifty-over chase

    >>> calculate_required_run_rate(
    ...     target_runs=301,
    ...     current_runs=200,
    ...     total_overs=50,
    ...     overs_bowled=35,
    ... )
    6.73

    Notes
    -----
    If the batting side has already reached or exceeded the target,
    the required run rate is zero.

    Internally this function performs every calculation using legal
    deliveries instead of decimal over notation.
    """

    _validate_non_negative(
        target_runs,
        "Target runs",
    )

    _validate_non_negative(
        current_runs,
        "Current runs",
    )

    #
    # Target achieved.
    #
    if current_runs >= target_runs:
        return 0.0

    balls_left = remaining_balls(
        total_overs=total_overs,
        overs_bowled=overs_bowled,
    )

    _validate_positive_balls(
        balls_left,
        "Remaining balls",
    )

    runs_required = target_runs - current_runs

    required_rate = (
        runs_required * BALLS_PER_OVER
    ) / balls_left

    return round(required_rate, precision)