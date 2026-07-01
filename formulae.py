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