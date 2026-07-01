"""
===============================================================================
Project : Cricket Excel Tournament Scoring Engine
Module  : formulas.py

Author  : <Your Name>
Version : 0.1.0

Description
-------------------------------------------------------------------------------

This module contains ONLY pure cricket calculation functions.

It intentionally has **NO dependency** on:

    - openpyxl
    - Workbook
    - Worksheet
    - Excel formulas
    - File system
    - GUI

The purpose of this module is to provide deterministic cricket calculations
that can be reused by every other component of the project.

The module should never directly modify workbook contents.

Instead, other modules (score_ball.py, leaderboard.py, scorecard.py, etc.)
will call these functions and then write the returned values into Excel.

Benefits
--------

✓ Easy to test
✓ Easy to reuse
✓ Independent of Excel
✓ Future web application can reuse the exact same logic
✓ Suitable for unit testing

-------------------------------------------------------------------------------
Future Sections

Part 1
    ✓ Constants
    ✓ Type aliases
    ✓ Helper conversion functions

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

from typing import Final, Literal, NewType

###############################################################################
#                               TYPE ALIASES
###############################################################################

# Using NewType improves readability without affecting runtime performance.
# These aliases make function signatures self-documenting.

Runs = NewType("Runs", int)
Balls = NewType("Balls", int)
Overs = NewType("Overs", float)
Wickets = NewType("Wickets", int)

###############################################################################
#                            MATCH CONSTANTS
###############################################################################

# Number of legal deliveries in a standard over.
BALLS_PER_OVER: Final[int] = 6

# Maximum wickets available in an innings.
MAX_WICKETS: Final[int] = 10

# Common match status values.
MatchStatus = Literal[
    "Scheduled",
    "Live",
    "Completed",
    "Abandoned",
    "No Result",
]

# Toss decision values.
TossDecision = Literal[
    "Bat",
    "Bowl",
]

###############################################################################
#                          DELIVERY CONSTANTS
###############################################################################

# Delivery types recognised by the scoring engine.
#
# A legal delivery counts towards the over.
#
# Wide and No Ball DO NOT count as legal deliveries.
#
# These constants will be referenced throughout the scoring engine.

LEGAL_DELIVERY_TYPES: Final[set[str]] = {
    "Normal",
    "Bye",
    "Leg Bye",
}

ILLEGAL_DELIVERY_TYPES: Final[set[str]] = {
    "Wide",
    "No Ball",
}

ALL_DELIVERY_TYPES: Final[set[str]] = (
    LEGAL_DELIVERY_TYPES |
    ILLEGAL_DELIVERY_TYPES
)

###############################################################################
#                         VALIDATION HELPER FUNCTIONS
###############################################################################


def _validate_non_negative(value: int | float, field_name: str) -> None:
    """
    Validate that a numeric value is not negative.

    Parameters
    ----------
    value
        Numeric value to validate.

    field_name
        Human-readable field name.

    Raises
    ------
    ValueError
        If the supplied value is negative.
    """

    if value < 0:
        raise ValueError(f"{field_name} cannot be negative.")


###############################################################################
#                         CONVERSION FUNCTIONS
###############################################################################


def overs_to_balls(overs: float) -> Balls:
    """
    Convert cricket overs into legal deliveries.

    Cricket overs are NOT decimal numbers.

    Example
    -------

    10.3 overs

    means

        10 overs
        +
        3 legal deliveries

    NOT

        10.3 * 6

    Examples
    --------

    >>> overs_to_balls(0.0)
    0

    >>> overs_to_balls(1.0)
    6

    >>> overs_to_balls(10.3)
    63

    >>> overs_to_balls(19.5)
    119

    Parameters
    ----------
    overs
        Cricket overs.

    Returns
    -------
    Balls

        Total legal deliveries.

    Raises
    ------
    ValueError

        If

        • overs is negative

        • decimal portion exceeds 5
    """

    _validate_non_negative(overs, "Overs")

    whole_overs = int(overs)

    balls = int(round((overs - whole_overs) * 10))

    if balls >= BALLS_PER_OVER:
        raise ValueError(
            "Invalid cricket over notation. "
            "Decimal part cannot exceed 5."
        )

    total = (whole_overs * BALLS_PER_OVER) + balls

    return Balls(total)


def balls_to_overs(balls: int) -> Overs:
    """
    Convert legal deliveries into cricket over notation.

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

        Cricket notation.

    Raises
    ------
    ValueError

        If balls is negative.
    """

    _validate_non_negative(balls, "Balls")

    completed_overs = balls // BALLS_PER_OVER

    remaining_balls = balls % BALLS_PER_OVER

    over_value = float(f"{completed_overs}.{remaining_balls}")

    return Overs(over_value)


###############################################################################
#                      DELIVERY VALIDATION FUNCTIONS
###############################################################################


def is_legal_delivery(delivery_type: str) -> bool:
    """
    Determine whether a delivery counts as a legal ball.

    Legal deliveries increment

        • over
        • ball number

    Illegal deliveries

        • Wide
        • No Ball

    do NOT count towards the over.

    Parameters
    ----------
    delivery_type

        Type of delivery.

    Returns
    -------
    bool

        True if legal.

    Raises
    ------
    ValueError

        Unknown delivery type.
    """

    if delivery_type not in ALL_DELIVERY_TYPES:
        raise ValueError(
            f"Unknown delivery type '{delivery_type}'."
        )

    return delivery_type in LEGAL_DELIVERY_TYPES


def legal_deliveries(delivery_types: list[str]) -> Balls:
    """
    Count legal deliveries in a sequence.

    This helper is commonly used while rebuilding an innings from
    the event log.

    Example
    -------

    >>> legal_deliveries([
    ...     "Normal",
    ...     "Wide",
    ...     "Normal",
    ...     "No Ball",
    ... ])
    2

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

        If an unknown delivery type is encountered.
    """

    count = 0

    for delivery in delivery_types:

        if is_legal_delivery(delivery):
            count += 1

    return Balls(count)


###############################################################################
#                         MODULE EXPORTS
###############################################################################

__all__ = [
    "Runs",
    "Balls",
    "Overs",
    "Wickets",
    "BALLS_PER_OVER",
    "MAX_WICKETS",
    "MatchStatus",
    "TossDecision",
    "LEGAL_DELIVERY_TYPES",
    "ILLEGAL_DELIVERY_TYPES",
    "ALL_DELIVERY_TYPES",
    "overs_to_balls",
    "balls_to_overs",
    "is_legal_delivery",
    "legal_deliveries",
]