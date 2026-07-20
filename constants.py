"""
constants.py
============

Centralized application-wide constants for the Cricket Tournament Workbook.

Overview
--------
This module serves as the single authoritative location for immutable values
that are shared across the entire application.

Rather than scattering repeated string literals, version numbers, formatting
rules, application names, and other configuration values throughout the code
base, they are defined once in this module and imported wherever needed.

This design provides several important advantages:

* Eliminates duplicated literals.
* Prevents accidental inconsistencies.
* Makes future maintenance significantly easier.
* Improves IDE auto-completion.
* Improves readability.
* Simplifies testing.
* Encourages strong separation of concerns.

The values defined here are intended to remain immutable throughout the
lifetime of the application.

No code within this module should modify any exported constant.

Relationship with Other Modules
-------------------------------

This module is intentionally lightweight and dependency-free so that it can be
imported safely from anywhere in the project without creating circular imports.

The following modules depend on this file.

---------------------------------------------------------------------------
workbook.py
---------------------------------------------------------------------------

The workbook builder imports shared metadata such as:

* application name
* workbook version
* workbook metadata
* formatting conventions
* logging namespace

Future workbook-related constants will also be defined here once the workbook
structure becomes stable.

---------------------------------------------------------------------------
validation.py
---------------------------------------------------------------------------

Validation logic imports shared values to avoid hard-coded literals.

Examples include:

* application version
* logging names
* formatting conventions
* future validation configuration

Keeping these values centralized ensures all validation rules remain
consistent across worksheets.

---------------------------------------------------------------------------
formulas.py
---------------------------------------------------------------------------

Formula helper functions use common application metadata where appropriate.

Although formulas primarily contain mathematical logic, importing shared
constants prevents duplicated strings and maintains consistency across modules.

---------------------------------------------------------------------------
create_match.py
---------------------------------------------------------------------------

The match creation module uses common metadata for workbook generation,
logging, timestamps, and future configuration values.

By importing constants instead of defining them locally, the module remains
clean and easier to maintain.

---------------------------------------------------------------------------
score_ball.py
---------------------------------------------------------------------------

Ball-by-ball scoring imports shared application metadata and future scoring
configuration values.

This avoids repeated literals throughout scoring logic while keeping business
rules centralized.

Design Principles
-----------------

This module follows several architectural principles.

Single Source of Truth
    Every application-wide immutable value should exist only once.

Immutability
    All exported values are typed using ``typing.Final`` to communicate that
    they must never be reassigned.

Dependency Safety
    This module should never import any project module.

Low Coupling
    Other modules depend on this module, but this module depends only upon the
    Python standard library.

High Cohesion
    Only globally reusable constants belong here.

Predictability
    Constants should never change during execution.

Future Expansion
----------------

As the project grows, this module will gradually become the home for:

* workbook metadata
* worksheet names
* table names
* named ranges
* cricket terminology
* match constants
* tournament constants
* formatting constants
* colours
* fonts
* borders
* validation limits
* scoring constants
* statistics constants
* application defaults

Only values that are reused across multiple modules should be placed here.

Values that belong exclusively to one module should remain inside that module.

Implementation Notes
--------------------

This module intentionally contains:

* no executable statements
* no helper functions
* no workbook creation logic
* no validation logic
* no formulas
* no side effects

Its sole responsibility is defining immutable shared values.
"""

from __future__ import annotations

from typing import Final, Literal

__author__: Final[str] = "Gaurav R Panchal"
__version__: Final[str] = "1.0.0"

# ============================================================================
# Application Metadata
# ============================================================================

APPLICATION_NAME: Final[str] = "Cricket Tournament Workbook"

WORKBOOK_VERSION: Final[str] = "1.0.0"

# ============================================================================
# Formatting Standards
# ============================================================================

DATE_FORMAT: Final[str] = "YYYY-MM-DD"

TIME_FORMAT: Final[str] = "HH:MM:SS"

# ============================================================================
# Logging
# ============================================================================

LOGGER_NAME: Final[str] = "cricket_workbook"

# ============================================================================
# General Application Literals
# ============================================================================

APPLICATION_MODE: Final[
    Literal["development", "testing", "production"]
] = "development"


# ============================================================================
# Workbook Metadata
# ============================================================================
#
# The constants in this section describe workbook-level metadata rather than
# worksheet-specific configuration.
#
# These values may be written into the Excel workbook properties when the
# workbook is created. Keeping them here provides a single source of truth for
# every module that needs workbook identification or metadata.
#
# Future versions of workbook.py can import these values directly when
# populating workbook properties without duplicating string literals.
# ============================================================================

WORKBOOK_TITLE: Final[str] = "Cricket Tournament Manager"

WORKBOOK_DESCRIPTION: Final[str] = (
    "An Excel-based cricket tournament management system supporting "
    "tournaments, teams, players, fixtures, live scoring, scorecards, "
    "statistics, and reporting."
)

WORKBOOK_AUTHOR: Final[str] = "Your Name"

WORKBOOK_COMPANY: Final[str] = "Your Organization"

WORKBOOK_SUBJECT: Final[str] = (
    "Cricket Tournament Management"
)

WORKBOOK_CATEGORY: Final[str] = (
    "Sports Management"
)

WORKBOOK_KEYWORDS: Final[str] = (
    "cricket, tournament, scoring, scorebook, excel, sports, statistics"
)

# ============================================================================
# Default Workbook File
# ============================================================================
#
# Default filename suggested when saving a newly created workbook.
#
# This value intentionally excludes any directory path so that callers remain
# free to decide where the workbook should be stored.
# ============================================================================

DEFAULT_WORKBOOK_FILENAME: Final[str] = (
    "Cricket_Tournament_Manager.xlsx"
)

# ============================================================================
# Workbook Appearance
# ============================================================================
#
# Workbook-wide presentation defaults.
#
# The theme name is intended as a logical identifier used by the application.
# It is not tied to any specific Office XML theme implementation and may later
# be mapped to a concrete Excel theme if required.
# ============================================================================

WORKBOOK_THEME_NAME: Final[str] = "Default"

# ============================================================================
# Worksheet Ordering
# ============================================================================
#
# Defines the preferred ordering of worksheets immediately after workbook
# creation.
#
# workbook.py should create worksheets in this order whenever possible.
#
# Only worksheet names belong here.
#
# Sheet-specific settings (tab colours, protection, dimensions, validations,
# etc.) should be defined elsewhere in future sections of constants.py.
# ============================================================================

DEFAULT_WORKSHEET_ORDER: Final[tuple[str, ...]] = (
    "Tournament",
    "Teams",
    "Players",
    "Fixtures",
    "Match",
    "Ball by Ball",
    "Scorecard",
    "Batting Stats",
    "Bowling Stats",
    "Points Table",
    "Reports",
    "Lookup",
)

# ============================================================================
# Workbook Protection Defaults
# ============================================================================
#
# These values represent the application's default behaviour regarding workbook
# protection.
#
# workbook.py may override these values for specialised use cases, but these
# serve as the standard defaults throughout the application.
# ============================================================================

DEFAULT_WORKBOOK_STRUCTURE_PROTECTION: Final[bool] = False

DEFAULT_WORKBOOK_WINDOWS_PROTECTION: Final[bool] = False

DEFAULT_WORKBOOK_PASSWORD: Final[str] = ""

# ============================================================================
# Freeze Pane Defaults
# ============================================================================
#
# Default freeze-pane location applied to newly created worksheets unless a
# sheet explicitly specifies otherwise.
#
# "A1" indicates that no rows or columns are frozen.
# ============================================================================

DEFAULT_FREEZE_PANE: Final[str] = "A1"

# ============================================================================
# Auto-Fit Defaults
# ============================================================================
#
# Excel does not provide native automatic column sizing through the workbook
# file format. workbook.py will therefore implement its own sizing logic.
#
# These constants indicate whether that sizing should be attempted by default.
# ============================================================================

DEFAULT_AUTO_FIT_COLUMNS: Final[bool] = True

DEFAULT_AUTO_FIT_ROWS: Final[bool] = False

# ============================================================================
# Excel Limits
# ============================================================================
#
# Microsoft Excel limits worksheet names to 31 characters.
#
# Defining the limit here avoids "magic numbers" throughout the project and
# makes validation.py, workbook.py, and create_match.py consistent.
# ============================================================================

MAX_SHEET_NAME_LENGTH: Final[int] = 31



# ============================================================================
# Worksheet Names
# ============================================================================
#
# This section defines the canonical names of every worksheet used throughout
# the Cricket Tournament Workbook.
#
# IMPORTANT
# ---------
# Every worksheet should be referenced using these constants rather than
# embedding literal sheet names throughout the project.
#
# Benefits:
#
# • Eliminates duplicated string literals.
# • Prevents spelling inconsistencies.
# • Makes worksheet renaming a single-point change.
# • Improves IDE refactoring support.
# • Simplifies workbook generation.
# • Ensures validation.py, workbook.py, create_match.py, and score_ball.py
#   always refer to the exact same worksheet names.
#
# Worksheet-specific behaviour (column layouts, formatting, protection,
# validations, etc.) should be defined elsewhere.
# ============================================================================

SHEET_HOME: Final[str] = "Home"

SHEET_TOURNAMENT: Final[str] = "Tournament"

SHEET_TEAMS: Final[str] = "Teams"

SHEET_PLAYERS: Final[str] = "Players"

SHEET_FIXTURES: Final[str] = "Fixtures"

SHEET_MATCH_SUMMARY: Final[str] = "Match Summary"

SHEET_SCORECARD: Final[str] = "Scorecard"

SHEET_BATTING: Final[str] = "Batting"

SHEET_BOWLING: Final[str] = "Bowling"

SHEET_BALL_BY_BALL: Final[str] = "Ball by Ball"

SHEET_EXTRAS: Final[str] = "Extras"

SHEET_STATISTICS: Final[str] = "Statistics"

SHEET_STANDINGS: Final[str] = "Standings"

SHEET_SETTINGS: Final[str] = "Settings"

SHEET_VALIDATION: Final[str] = "Validation"

SHEET_HIDDEN_LISTS: Final[str] = "Hidden Lists"

SHEET_NAVIGATION: Final[str] = "Navigation"

SHEET_REPORTS: Final[str] = "Reports"

# ============================================================================
# Workbook Worksheet Order
# ============================================================================
#
# Defines the canonical worksheet ordering for the workbook.
#
# workbook.py should create worksheets in this sequence unless an explicit
# reason exists to override it.
#
# Navigation-oriented sheets are intentionally placed near the beginning,
# operational sheets occupy the middle, reporting sheets follow, and internal
# configuration sheets are placed at the end.
# ============================================================================

WORKSHEET_ORDER: Final[tuple[str, ...]] = (
    SHEET_HOME,
    SHEET_TOURNAMENT,
    SHEET_TEAMS,
    SHEET_PLAYERS,
    SHEET_FIXTURES,
    SHEET_MATCH_SUMMARY,
    SHEET_SCORECARD,
    SHEET_BATTING,
    SHEET_BOWLING,
    SHEET_BALL_BY_BALL,
    SHEET_EXTRAS,
    SHEET_STATISTICS,
    SHEET_STANDINGS,
    SHEET_REPORTS,
    SHEET_NAVIGATION,
    SHEET_SETTINGS,
    SHEET_VALIDATION,
    SHEET_HIDDEN_LISTS,
)

# ============================================================================
# Visible Worksheets
# ============================================================================
#
# These worksheets are intended to be visible to end users during normal
# workbook operation.
#
# workbook.py may use this collection when configuring worksheet visibility.
# ============================================================================

VISIBLE_SHEETS: Final[tuple[str, ...]] = (
    SHEET_HOME,
    SHEET_TOURNAMENT,
    SHEET_TEAMS,
    SHEET_PLAYERS,
    SHEET_FIXTURES,
    SHEET_MATCH_SUMMARY,
    SHEET_SCORECARD,
    SHEET_BATTING,
    SHEET_BOWLING,
    SHEET_BALL_BY_BALL,
    SHEET_EXTRAS,
    SHEET_STATISTICS,
    SHEET_STANDINGS,
    SHEET_REPORTS,
    SHEET_NAVIGATION,
    SHEET_SETTINGS,
)

# ============================================================================
# Hidden Worksheets
# ============================================================================
#
# These worksheets contain supporting data, lookup lists, validation ranges,
# or other implementation details that are not intended for routine user
# interaction.
#
# workbook.py should mark these worksheets as hidden immediately after
# creation. Other modules should continue referring to them using these
# constants regardless of visibility.
# ============================================================================

HIDDEN_SHEETS: Final[tuple[str, ...]] = (
    SHEET_VALIDATION,
    SHEET_HIDDEN_LISTS,
)

# ============================================================================
# Cricket Match Configuration
# ============================================================================
#
# This section contains the core constants that define the structure of a
# cricket match. These values are intended to be format-independent defaults
# suitable for most limited-overs competitions.
#
# Centralising these constants ensures that workbook.py, create_match.py,
# score_ball.py, validation.py and formulas.py all operate using the same
# assumptions.
#
# Future versions may introduce tournament-specific rule sets (e.g. T20,
# One Day, The Hundred, custom leagues). Those configurations should build
# upon these defaults rather than redefining literal values throughout the
# application.
# ============================================================================

# ============================================================================
# Innings Configuration
# ============================================================================

MAX_INNINGS: Final[int] = 2

# ============================================================================
# Over Configuration
# ============================================================================
#
# Default maximum overs per innings.
#
# Individual tournaments may override this value (e.g. 5, 10, 20, 50 overs),
# but 20 overs is used as the application default.
# ============================================================================

DEFAULT_MAX_OVERS: Final[int] = 20

# ============================================================================
# Ball Configuration
# ============================================================================
#
# Modern cricket uses six legal deliveries per over.
# ============================================================================

BALLS_PER_OVER: Final[int] = 6

# ============================================================================
# Wicket Configuration
# ============================================================================
#
# A batting innings normally concludes after ten wickets have fallen.
# ============================================================================

MAX_WICKETS: Final[int] = 10

# ============================================================================
# Toss Options
# ============================================================================
#
# Valid outcomes of the coin toss.
# ============================================================================

TOSS_CHOICES: Final[tuple[str, ...]] = (
    "Heads",
    "Tails",
)

# ============================================================================
# Toss Decision Options
# ============================================================================
#
# Decision made by the toss-winning captain.
# ============================================================================

TOSS_DECISION_CHOICES: Final[tuple[str, ...]] = (
    "Bat",
    "Bowl",
)

# ============================================================================
# Match Result Types
# ============================================================================
#
# Standard match outcomes recognised by the application.
#
# Additional tournament-specific outcomes may be appended in future releases.
# ============================================================================

MATCH_RESULT_TYPES: Final[tuple[str, ...]] = (
    "Team 1 Won",
    "Team 2 Won",
    "Tie",
    "No Result",
    "Abandoned",
    "Cancelled",
)

# ============================================================================
# Super Over Configuration
# ============================================================================
#
# Default values governing Super Overs.
#
# These defaults are based on the commonly adopted ICC playing conditions.
# ============================================================================

SUPER_OVER_ENABLED: Final[bool] = True

SUPER_OVER_MAX_OVERS: Final[int] = 1

SUPER_OVER_MAX_WICKETS: Final[int] = 2

SUPER_OVER_BALLS_PER_OVER: Final[int] = 6

# ============================================================================
# Rain Interruption Configuration
# ============================================================================
#
# These values are used by future modules implementing reduced-over matches,
# interruptions, and Duckworth-Lewis-Stern (DLS) calculations.
#
# They intentionally represent application states rather than mathematical
# calculations.
# ============================================================================

RAIN_INTERRUPTION_VALUES: Final[tuple[str, ...]] = (
    "None",
    "Delayed Start",
    "Match Interrupted",
    "Reduced Overs",
    "DLS Applied",
    "Abandoned",
)

# ============================================================================
# Powerplay Defaults
# ============================================================================
#
# Default powerplay structure for a standard T20 innings.
#
# Tournament-specific configurations may override these values.
# ============================================================================

DEFAULT_POWERPLAY_START_OVER: Final[int] = 1

DEFAULT_POWERPLAY_END_OVER: Final[int] = 6

POWERPLAY_NAMES: Final[tuple[str, ...]] = (
    "Powerplay",
    "Middle Overs",
    "Death Overs",
)

# ============================================================================
# Field Restriction Phases
# ============================================================================
#
# Human-readable names describing the current field restriction phase.
# ============================================================================

FIELD_RESTRICTION_NAMES: Final[tuple[str, ...]] = (
    "Powerplay",
    "Normal Field",
    "Death Overs",
)

# ============================================================================
# Match Status Values
# ============================================================================
#
# Lifecycle states recognised by the application.
#
# score_ball.py and create_match.py should update match progression using
# these values instead of embedding literal strings.
# ============================================================================

MATCH_STATUS_VALUES: Final[tuple[str, ...]] = (
    "Scheduled",
    "Toss Pending",
    "First Innings",
    "Innings Break",
    "Second Innings",
    "Super Over",
    "Completed",
    "Abandoned",
    "No Result",
    "Cancelled",
)

# ============================================================================
# Player Configuration
# ============================================================================
#
# This section defines all player-related constants used throughout the
# application.
#
# These constants provide a single source of truth for player metadata and
# should be referenced by:
#
#   • workbook.py
#   • validation.py
#   • create_match.py
#   • score_ball.py
#   • formulas.py
#
# Centralising these values avoids duplicated string literals, simplifies
# dropdown generation, and ensures consistent terminology throughout the
# workbook.
#
# Future versions may extend these tuples to support additional competition
# rules, player classifications, or custom tournament requirements.
# ============================================================================

# ============================================================================
# Player Roles
# ============================================================================
#
# Defines the primary cricketing role of a player.
#
# Batter
#     Specialist batter.
#
# Bowler
#     Specialist bowler.
#
# All-rounder
#     Contributes significantly with both bat and ball.
#
# Wicket Keeper
#     Specialist wicket keeper.
#
# Wicket Keeper Batter
#     Modern wicket keeper whose primary role also includes batting.
# ============================================================================

PLAYER_ROLES: Final[tuple[str, ...]] = (
    "Batter",
    "Bowler",
    "All-rounder",
    "Wicket Keeper",
    "Wicket Keeper Batter",
)

# ============================================================================
# Leadership Roles
# ============================================================================
#
# Team leadership designations.
#
# These values are intentionally independent of PLAYER_ROLES because any
# eligible player may serve as captain or vice captain.
# ============================================================================

TEAM_LEADERSHIP_ROLES: Final[tuple[str, ...]] = (
    "Captain",
    "Vice Captain",
)

# ============================================================================
# Captaincy Labels
# ============================================================================
#
# Frequently used labels throughout scorecards and reports.
# ============================================================================

CAPTAIN_LABEL: Final[str] = "Captain"

VICE_CAPTAIN_LABEL: Final[str] = "Vice Captain"

# ============================================================================
# Wicket Keeper Labels
# ============================================================================
#
# Used when identifying designated wicket keepers in reports, scorecards,
# team sheets and match summaries.
# ============================================================================

WICKET_KEEPER_LABEL: Final[str] = "Wicket Keeper"

# ============================================================================
# Batting Styles
# ============================================================================
#
# Describes the player's batting handedness.
# ============================================================================

BATTING_STYLES: Final[tuple[str, ...]] = (
    "Right-hand Bat",
    "Left-hand Bat",
)

# ============================================================================
# Bowling Styles
# ============================================================================
#
# Describes the bowling discipline.
#
# The bowling arm is intentionally stored separately so applications may
# combine arm and style programmatically (e.g. Right-arm Fast Medium).
# ============================================================================

BOWLING_STYLES: Final[tuple[str, ...]] = (
    "Fast",
    "Fast Medium",
    "Medium Fast",
    "Medium",
    "Off Break",
    "Leg Break",
    "Leg Spin",
    "Left-arm Orthodox",
    "Left-arm Wrist Spin",
    "Slow Left-arm",
    "None",
)

# ============================================================================
# Bowling Arms
# ============================================================================
#
# Specifies the arm with which the player bowls.
# ============================================================================

BOWLING_ARMS: Final[tuple[str, ...]] = (
    "Right-arm",
    "Left-arm",
    "None",
)

# ============================================================================
# Player Status
# ============================================================================
#
# Indicates the player's current registration or squad status.
#
# Active
#     Available for normal selection.
#
# Injured
#     Temporarily unavailable due to injury.
#
# Suspended
#     Unavailable because of disciplinary action.
#
# Retired
#     No longer participating.
#
# Unavailable
#     Temporarily unavailable for reasons other than injury.
# ============================================================================

PLAYER_STATUS: Final[tuple[str, ...]] = (
    "Active",
    "Injured",
    "Suspended",
    "Retired",
    "Unavailable",
)

# ============================================================================
# Substitute Types
# ============================================================================
#
# Defines recognised substitute categories.
#
# These values accommodate traditional substitute fielders as well as modern
# ICC replacement regulations.
# ============================================================================

SUBSTITUTE_TYPES: Final[tuple[str, ...]] = (
    "Substitute Fielder",
    "Concussion Substitute",
    "COVID Replacement",
    "Impact Player",
    "None",
)

# ============================================================================
# Player Availability
# ============================================================================
#
# Used during squad selection and match preparation.
#
# Available
#     Eligible for selection.
#
# Selected
#     Included in the playing XI.
#
# Rested
#     Available but intentionally not selected.
#
# Injured
#     Unable to participate.
#
# Unavailable
#     Not available for selection.
#
# On Leave
#     Temporarily absent for personal or official reasons.
# ============================================================================

PLAYER_AVAILABILITY: Final[tuple[str, ...]] = (
    "Available",
    "Selected",
    "Rested",
    "Injured",
    "Unavailable",
    "On Leave",
)

# ============================================================================
# Ball Outcome Configuration
# ============================================================================
#
# This section defines the canonical constants describing the outcome of an
# individual ball.
#
# These values are shared across:
#
#   • score_ball.py
#   • formulas.py
#   • validation.py
#   • workbook.py
#   • create_match.py
#
# Every delivery recorded in the workbook should use these constants rather
# than embedded string literals. This ensures consistent terminology, reliable
# validation, and easier maintenance.
#
# Terminology follows the ICC Laws of Cricket wherever practical.
# ============================================================================

# ============================================================================
# Delivery Types
# ============================================================================
#
# Legal deliveries count towards the over.
#
# Illegal deliveries do NOT count as one of the six legal balls and therefore
# require an additional delivery to complete the over.
# ============================================================================

LEGAL_DELIVERIES: Final[tuple[str, ...]] = (
    "Legal Delivery",
)

ILLEGAL_DELIVERIES: Final[tuple[str, ...]] = (
    "No Ball",
    "Wide",
)

# ============================================================================
# Ball Run Values
# ============================================================================
#
# These represent runs scored from the striker's bat or completed running
# between the wickets.
#
# Extras (byes, leg byes, wides and no balls) are intentionally handled
# separately.
# ============================================================================

RUN_VALUES: Final[tuple[int, ...]] = (
    0,
    1,
    2,
    3,
    4,
    5,
    6,
)

# ============================================================================
# Named Run Outcomes
# ============================================================================
#
# Human-readable names frequently used throughout scorecards, reports and
# statistics.
# ============================================================================

DOT_BALL: Final[int] = 0

SINGLE: Final[int] = 1

DOUBLE: Final[int] = 2

TRIPLE: Final[int] = 3

FOUR: Final[int] = 4

FIVE: Final[int] = 5

SIX: Final[int] = 6

# ============================================================================
# Boundary Values
# ============================================================================
#
# Numeric values representing recognised boundary scores.
# ============================================================================

BOUNDARY_RUN_VALUES: Final[tuple[int, ...]] = (
    FOUR,
    SIX,
)

# ============================================================================
# Boundary Names
# ============================================================================
#
# Used by validation and reporting modules when identifying boundaries.
# ============================================================================

BOUNDARY_TYPES: Final[tuple[str, ...]] = (
    "Four",
    "Six",
)

# ============================================================================
# Dismissal Types
# ============================================================================
#
# Official dismissal methods recognised by the ICC Laws of Cricket.
#
# These values should be used consistently throughout scoring, statistics and
# validation.
#
# Notes
# -----
# • "Run Out" includes both striker and non-striker dismissals.
# • "Handled the Ball" has been incorporated into
#   "Obstructing the Field" under the modern Laws.
# • "Retired Out" is included because it is now recognised as a dismissal
#   when a batter retires without the umpire's consent.
# ============================================================================

DISMISSAL_TYPES: Final[tuple[str, ...]] = (
    "Bowled",
    "Caught",
    "Caught and Bowled",
    "LBW",
    "Run Out",
    "Stumped",
    "Hit Wicket",
    "Hit the Ball Twice",
    "Obstructing the Field",
    "Timed Out",
    "Retired Out",
)

# ============================================================================
# Frequently Referenced Dismissal Constants
# ============================================================================
#
# Individual constants improve readability throughout the codebase and reduce
# repeated string literals.
# ============================================================================

DISMISSAL_BOWLED: Final[str] = "Bowled"

DISMISSAL_CAUGHT: Final[str] = "Caught"

DISMISSAL_CAUGHT_AND_BOWLED: Final[str] = "Caught and Bowled"

DISMISSAL_LBW: Final[str] = "LBW"

DISMISSAL_RUN_OUT: Final[str] = "Run Out"

DISMISSAL_STUMPED: Final[str] = "Stumped"

DISMISSAL_HIT_WICKET: Final[str] = "Hit Wicket"

DISMISSAL_HIT_BALL_TWICE: Final[str] = "Hit the Ball Twice"

DISMISSAL_OBSTRUCTING_FIELD: Final[str] = "Obstructing the Field"

DISMISSAL_TIMED_OUT: Final[str] = "Timed Out"

DISMISSAL_RETIRED_OUT: Final[str] = "Retired Out"


# ============================================================================
# Extras Configuration
# ============================================================================
#
# This section defines every extra (also known as an "extra run") recognised
# by the Cricket Tournament Workbook.
#
# Extras are runs awarded that are not credited to the striker's batting
# statistics. They contribute to the team's total while being recorded
# separately for scorecards and statistical reporting.
#
# These constants provide a single source of truth for:
#
#   • score_ball.py
#   • validation.py
#   • workbook.py
#   • formulas.py
#   • create_match.py
#
# Using these constants throughout the project avoids duplicated literals,
# simplifies validation dropdowns, and ensures consistent terminology.
#
# Terminology follows the ICC Laws of Cricket where applicable.
# ============================================================================

# ============================================================================
# Individual Extra Types
# ============================================================================
#
# Frequently referenced names for the five recognised categories of extras.
#
# These constants are intended for readability throughout the codebase.
# ============================================================================

EXTRA_WIDE: Final[str] = "Wide"

EXTRA_NO_BALL: Final[str] = "No Ball"

EXTRA_BYE: Final[str] = "Bye"

EXTRA_LEG_BYE: Final[str] = "Leg Bye"

EXTRA_PENALTY: Final[str] = "Penalty"

# ============================================================================
# Extra Classifications
# ============================================================================
#
# Canonical list of all supported extra categories.
#
# Validation dropdowns and scoring modules should reference this tuple instead
# of embedding literal values.
# ============================================================================

EXTRA_CLASSIFICATIONS: Final[tuple[str, ...]] = (
    EXTRA_WIDE,
    EXTRA_NO_BALL,
    EXTRA_BYE,
    EXTRA_LEG_BYE,
    EXTRA_PENALTY,
)

# ============================================================================
# Wide Reasons
# ============================================================================
#
# Describes why a delivery has been adjudged a Wide.
#
# These values provide meaningful classifications for reporting while
# remaining general enough to accommodate different competitions and playing
# conditions.
# ============================================================================

WIDE_REASONS: Final[tuple[str, ...]] = (
    "Outside Off Stump",
    "Outside Leg Stump",
    "Over Batter's Head",
    "Bounce Over Head",
    "Unreachable Delivery",
)

# ============================================================================
# No Ball Reasons
# ============================================================================
#
# Common reasons why a delivery is declared a No Ball.
#
# The list reflects the most frequently encountered ICC Law applications.
# Future tournament-specific rules may extend this tuple if required.
# ============================================================================

NO_BALL_REASONS: Final[tuple[str, ...]] = (
    "Front Foot",
    "Back Foot",
    "Above Waist Full Toss",
    "Dangerous Short-Pitched Ball",
    "Multiple Bouncers",
    "Illegal Fielding Position",
    "Throwing",
    "Breaking the Wicket",
    "Other",
)

# ============================================================================
# Penalty Reasons
# ============================================================================
#
# Reasons for awarding penalty runs.
#
# These are intentionally generic because the Laws of Cricket permit penalty
# runs in a variety of situations involving player conduct, unfair play,
# fielding infringements and equipment.
# ============================================================================

PENALTY_REASONS: Final[tuple[str, ...]] = (
    "Unfair Play",
    "Player Misconduct",
    "Ball Tampering",
    "Illegal Fielding",
    "Damage to Pitch",
    "Time Wasting",
    "Protective Equipment",
    "Other",
)

# ============================================================================
# Extra Summary Groups
# ============================================================================
#
# Useful groupings for reporting and statistical aggregation.
#
# Bowling Extras
#     Extras directly attributed to the bowler.
#
# Fielding Extras
#     Extras not charged to the bowler.
#
# Team Extras
#     All recognised extras.
# ============================================================================

BOWLING_EXTRAS: Final[tuple[str, ...]] = (
    EXTRA_WIDE,
    EXTRA_NO_BALL,
)

FIELDING_EXTRAS: Final[tuple[str, ...]] = (
    EXTRA_BYE,
    EXTRA_LEG_BYE,
    EXTRA_PENALTY,
)

TEAM_EXTRAS: Final[tuple[str, ...]] = EXTRA_CLASSIFICATIONS


# ============================================================================
# Validation Lists
# ============================================================================
#
# This section defines reusable lookup values used by validation.py to build
# Excel dropdown lists throughout the workbook.
#
# Every dropdown presented to the user should originate from these immutable
# constants rather than embedding literal values inside validation.py.
#
# Benefits
# --------
#
# • Consistent terminology across worksheets.
# • Single source of truth.
# • Easier localisation in future.
# • Simpler maintenance.
# • Reduced duplication.
# • Improved testability.
#
# The tuples in this section intentionally contain only user-selectable values.
# Validation logic belongs exclusively in validation.py.
# ============================================================================

# ============================================================================
# Generic Boolean Choices
# ============================================================================
#
# Frequently used throughout workbook configuration and worksheet validation.
# ============================================================================

YES_NO_OPTIONS: Final[tuple[str, ...]] = (
    "Yes",
    "No",
)

TRUE_FALSE_OPTIONS: Final[tuple[str, ...]] = (
    "True",
    "False",
)

ENABLED_DISABLED_OPTIONS: Final[tuple[str, ...]] = (
    "Enabled",
    "Disabled",
)

# ============================================================================
# Gender
# ============================================================================
#
# General-purpose gender options.
#
# These values are intentionally broad enough for most tournament
# administration scenarios and may be customised by future applications.
# ============================================================================

GENDER_OPTIONS: Final[tuple[str, ...]] = (
    "Male",
    "Female",
    "Other",
    "Prefer Not to Say",
)

# ============================================================================
# Countries
# ============================================================================
#
# Placeholder values.
#
# A future version of the project may replace these with a complete ISO-3166
# country list generated automatically.
# ============================================================================

COUNTRY_OPTIONS: Final[tuple[str, ...]] = (
    "Select Country",
)

# ============================================================================
# States / Provinces
# ============================================================================
#
# Placeholder values.
#
# Future implementations may populate these dynamically based on the selected
# country or jurisdiction.
# ============================================================================

STATE_OPTIONS: Final[tuple[str, ...]] = (
    "Select State / Province",
)

# ============================================================================
# Age Groups
# ============================================================================
#
# Common cricket competition age categories.
#
# These values are suitable for junior, youth and senior competitions.
# ============================================================================

AGE_GROUP_OPTIONS: Final[tuple[str, ...]] = (
    "Under 10",
    "Under 12",
    "Under 14",
    "Under 16",
    "Under 19",
    "Open",
    "Senior",
    "Veterans",
)

# ============================================================================
# Competition Formats
# ============================================================================
#
# Common match formats supported by the workbook.
#
# Custom formats can be added by future tournament implementations without
# changing validation.py.
# ============================================================================

COMPETITION_FORMAT_OPTIONS: Final[tuple[str, ...]] = (
    "T5",
    "T10",
    "T20",
    "T50",
    "One Day",
    "Multi-Day",
    "Test Match",
    "Custom",
)

# ============================================================================
# Ball Colours
# ============================================================================
#
# Standard cricket ball colours used internationally.
# ============================================================================

BALL_COLOUR_OPTIONS: Final[tuple[str, ...]] = (
    "Red",
    "White",
    "Pink",
    "Orange",
    "Yellow",
    "Other",
)

# ============================================================================
# Pitch Types
# ============================================================================
#
# Describes the general playing characteristics of the pitch.
# ============================================================================

PITCH_TYPE_OPTIONS: Final[tuple[str, ...]] = (
    "Hard",
    "Dry",
    "Green",
    "Dusty",
    "Grass",
    "Artificial",
    "Matting",
    "Concrete",
)

# ============================================================================
# Weather Conditions
# ============================================================================
#
# General weather descriptors used for match administration and reporting.
# ============================================================================

WEATHER_OPTIONS: Final[tuple[str, ...]] = (
    "Sunny",
    "Partly Cloudy",
    "Cloudy",
    "Overcast",
    "Windy",
    "Humid",
    "Light Rain",
    "Heavy Rain",
    "Storm",
    "Fog",
)

# ============================================================================
# Ground Conditions
# ============================================================================
#
# Overall condition of the playing surface and outfield.
# ============================================================================

GROUND_CONDITION_OPTIONS: Final[tuple[str, ...]] = (
    "Excellent",
    "Good",
    "Normal",
    "Soft",
    "Wet",
    "Damp",
    "Slippery",
    "Poor",
)

# ============================================================================
# Match Result Methods
# ============================================================================
#
# Describes how the official result of a match was determined.
# ============================================================================

RESULT_METHOD_OPTIONS: Final[tuple[str, ...]] = (
    "Normal Result",
    "Super Over",
    "Duckworth-Lewis-Stern (DLS)",
    "Forfeit",
    "Abandoned",
    "No Result",
    "Cancelled",
)

# ============================================================================
# Umpire Roles
# ============================================================================
#
# Match official assignments recognised by the workbook.
# ============================================================================

UMPIRE_ROLE_OPTIONS: Final[tuple[str, ...]] = (
    "On-field Umpire 1",
    "On-field Umpire 2",
    "Third Umpire",
    "Fourth Umpire",
    "Reserve Umpire",
    "Match Referee",
)

# ============================================================================
# Excel Formatting Configuration
# ============================================================================
#
# This section defines workbook-wide formatting defaults used during workbook
# generation.
#
# These constants are intentionally presentation-oriented. They describe how
# data should appear inside Excel and are shared by:
#
#     • workbook.py
#     • create_match.py
#     • score_ball.py
#     • reports (future)
#
# Centralising formatting values provides several benefits:
#
# • Consistent appearance across every worksheet.
# • Easy adjustment of workbook styling.
# • Elimination of "magic numbers".
# • Simpler testing.
# • Improved readability.
#
# No worksheet-specific formatting belongs here. Those settings should be
# defined alongside worksheet creation logic.
# ============================================================================

# ============================================================================
# Row Heights
# ============================================================================
#
# Default row dimensions measured in Excel points.
# ============================================================================

DEFAULT_ROW_HEIGHT: Final[float] = 18.0

HEADER_ROW_HEIGHT: Final[float] = 24.0

TITLE_ROW_HEIGHT: Final[float] = 30.0

# ============================================================================
# Default Column Widths
# ============================================================================
#
# Excel measures column width using the width of the default font.
#
# These values are intended to provide sensible defaults for workbook
# generation. Individual worksheets may override them where necessary.
# ============================================================================

DEFAULT_COLUMN_WIDTH: Final[float] = 15.0

NARROW_COLUMN_WIDTH: Final[float] = 8.0

MEDIUM_COLUMN_WIDTH: Final[float] = 15.0

WIDE_COLUMN_WIDTH: Final[float] = 22.0

EXTRA_WIDE_COLUMN_WIDTH: Final[float] = 35.0

# ============================================================================
# Numeric Precision
# ============================================================================
#
# Standard decimal precision used throughout the workbook.
#
# These values are intended for number formatting rather than calculations.
# ============================================================================

DEFAULT_DECIMAL_PRECISION: Final[int] = 2

PERCENTAGE_PRECISION: Final[int] = 2

CURRENCY_PRECISION: Final[int] = 2

RUN_RATE_PRECISION: Final[int] = 2

AVERAGE_PRECISION: Final[int] = 2

STRIKE_RATE_PRECISION: Final[int] = 2

ECONOMY_RATE_PRECISION: Final[int] = 2

# ============================================================================
# Date and Time Display Formats
# ============================================================================
#
# Excel-compatible number format strings.
#
# NOTE:
# Earlier in this module DATE_FORMAT and TIME_FORMAT define the application's
# logical formatting conventions. The constants below define the actual Excel
# number format strings applied to worksheet cells.
# ============================================================================

EXCEL_DATE_FORMAT: Final[str] = "yyyy-mm-dd"

EXCEL_TIME_FORMAT: Final[str] = "hh:mm:ss"

EXCEL_DATETIME_FORMAT: Final[str] = "yyyy-mm-dd hh:mm:ss"

# ============================================================================
# Standard Excel Number Formats
# ============================================================================
#
# Frequently reused Excel formatting strings.
#
# workbook.py should reference these constants rather than embedding literal
# format strings.
# ============================================================================

NUMBER_FORMAT_GENERAL: Final[str] = "General"

NUMBER_FORMAT_INTEGER: Final[str] = "0"

NUMBER_FORMAT_DECIMAL: Final[str] = "0.00"

NUMBER_FORMAT_THOUSANDS: Final[str] = "#,##0"

NUMBER_FORMAT_THOUSANDS_DECIMAL: Final[str] = "#,##0.00"

NUMBER_FORMAT_PERCENTAGE: Final[str] = "0.00%"

NUMBER_FORMAT_CURRENCY: Final[str] = '₹#,##0.00'

NUMBER_FORMAT_TEXT: Final[str] = "@"

NUMBER_FORMAT_OVERS: Final[str] = "0.0"

# ============================================================================
# Print Margins
# ============================================================================
#
# Measurements are expressed in inches to match Excel's page setup model.
#
# These defaults produce a balanced printable layout suitable for scorecards,
# reports and statistics.
# ============================================================================

PRINT_MARGIN_LEFT: Final[float] = 0.50

PRINT_MARGIN_RIGHT: Final[float] = 0.50

PRINT_MARGIN_TOP: Final[float] = 0.75

PRINT_MARGIN_BOTTOM: Final[float] = 0.75

PRINT_MARGIN_HEADER: Final[float] = 0.30

PRINT_MARGIN_FOOTER: Final[float] = 0.30

# ============================================================================
# Paper Sizes
# ============================================================================
#
# Logical paper size names.
#
# workbook.py may map these values to the corresponding Excel/OpenXML paper
# identifiers.
# ============================================================================

PAPER_SIZE_A4: Final[str] = "A4"

PAPER_SIZE_A3: Final[str] = "A3"

PAPER_SIZE_LETTER: Final[str] = "Letter"

DEFAULT_PAPER_SIZE: Final[str] = PAPER_SIZE_A4

# ============================================================================
# Page Orientation
# ============================================================================
#
# Standard Excel page orientations.
# ============================================================================

PAGE_ORIENTATION_PORTRAIT: Final[str] = "portrait"

PAGE_ORIENTATION_LANDSCAPE: Final[str] = "landscape"

DEFAULT_PAGE_ORIENTATION: Final[str] = PAGE_ORIENTATION_LANDSCAPE


# ============================================================================
# Workbook Colour Palette
# ============================================================================
#
# This section defines the standard colour palette used throughout the Cricket
# Tournament Workbook.
#
# These colours establish a consistent visual identity across every worksheet
# and should be referenced by workbook.py (and future styling modules) rather
# than embedding literal RGB values.
#
# Colour values are expressed as six-character RGB hexadecimal strings
# (without the leading '#') to match the format expected by libraries such as
# openpyxl.
#
# Design Goals
# ------------
#
# • Consistent appearance across all worksheets.
# • High readability.
# • Good contrast for accessibility.
# • Single source of truth.
# • Easy future rebranding.
#
# Future versions may expand this palette with chart colours, heat maps,
# conditional formatting palettes, and theme variants.
# ============================================================================

# ============================================================================
# Header Colours
# ============================================================================
#
# Used for worksheet titles, table headers, report headings and major section
# labels.
# ============================================================================

COLOUR_HEADER_BACKGROUND: Final[str] = "1F4E78"      # Dark Blue

COLOUR_HEADER_TEXT: Final[str] = "FFFFFF"            # White

# ============================================================================
# Table Colours
# ============================================================================
#
# Standard colours used by structured data tables.
# ============================================================================

COLOUR_TABLE_HEADER: Final[str] = "D9EAD3"

COLOUR_TABLE_BORDER: Final[str] = "B7B7B7"

COLOUR_TABLE_BACKGROUND: Final[str] = "FFFFFF"

# ============================================================================
# Alternating Row Colours
# ============================================================================
#
# Zebra-striping improves readability of larger datasets.
# ============================================================================

COLOUR_ROW_ODD: Final[str] = "FFFFFF"

COLOUR_ROW_EVEN: Final[str] = "F5F5F5"

# ============================================================================
# Status Colours
# ============================================================================
#
# Used for conditional formatting and user feedback.
# ============================================================================

COLOUR_SUCCESS: Final[str] = "C6EFCE"

COLOUR_SUCCESS_TEXT: Final[str] = "006100"

COLOUR_WARNING: Final[str] = "FFF2CC"

COLOUR_WARNING_TEXT: Final[str] = "9C6500"

COLOUR_ERROR: Final[str] = "F4CCCC"

COLOUR_ERROR_TEXT: Final[str] = "9C0006"

COLOUR_INFORMATION: Final[str] = "D9EAF7"

COLOUR_INFORMATION_TEXT: Final[str] = "0B5394"

# ============================================================================
# Cell State Colours
# ============================================================================
#
# Indicates how users should interact with worksheet cells.
# ============================================================================

COLOUR_EDITABLE_CELL: Final[str] = "FFF2CC"

COLOUR_LOCKED_CELL: Final[str] = "E6E6E6"

COLOUR_DISABLED_CELL: Final[str] = "D9D9D9"

# ============================================================================
# Navigation Colours
# ============================================================================
#
# Used by the Home and Navigation worksheets for hyperlinks, menus and
# workbook navigation controls.
# ============================================================================

COLOUR_NAVIGATION_BACKGROUND: Final[str] = "D9EAD3"

COLOUR_NAVIGATION_TEXT: Final[str] = "274E13"

COLOUR_NAVIGATION_HIGHLIGHT: Final[str] = "6AA84F"

# ============================================================================
# Tournament Branding
# ============================================================================
#
# Primary branding colours for banners, logos, dashboards and tournament
# identity.
#
# These colours can be changed in one place to re-theme the entire workbook.
# ============================================================================

COLOUR_BRAND_PRIMARY: Final[str] = "0B5394"

COLOUR_BRAND_SECONDARY: Final[str] = "6AA84F"

COLOUR_BRAND_ACCENT: Final[str] = "F1C232"

# ============================================================================
# Neutral Greys
# ============================================================================
#
# General-purpose greys used for borders, backgrounds, separators and subtle
# UI elements.
# ============================================================================

COLOUR_GREY_50: Final[str] = "F9F9F9"

COLOUR_GREY_100: Final[str] = "F2F2F2"

COLOUR_GREY_200: Final[str] = "E6E6E6"

COLOUR_GREY_300: Final[str] = "D9D9D9"

COLOUR_GREY_400: Final[str] = "BFBFBF"

COLOUR_GREY_500: Final[str] = "A6A6A6"

COLOUR_GREY_600: Final[str] = "7F7F7F"

COLOUR_GREY_700: Final[str] = "595959"

COLOUR_GREY_800: Final[str] = "404040"

COLOUR_GREY_900: Final[str] = "262626"

# ============================================================================
# Common Utility Colours
# ============================================================================
#
# Frequently reused colours for text, borders and fills.
# ============================================================================

COLOUR_WHITE: Final[str] = "FFFFFF"

COLOUR_BLACK: Final[str] = "000000"

COLOUR_RED: Final[str] = "FF0000"

COLOUR_GREEN: Final[str] = "008000"

COLOUR_BLUE: Final[str] = "0000FF"

COLOUR_YELLOW: Final[str] = "FFFF00"

COLOUR_ORANGE: Final[str] = "F6B26B"

# ============================================================================
# Excel Cell Formatting Constants
# ============================================================================
#
# This section defines workbook-wide cell formatting constants that are shared
# by workbook.py and any future styling modules.
#
# These constants intentionally mirror common Excel/OpenXML terminology so they
# can be mapped directly to openpyxl formatting objects such as:
#
#     Border
#     Side
#     Alignment
#     Protection
#
# Defining these values here avoids scattered string literals throughout the
# project and makes workbook styling easier to maintain.
#
# No formatting logic belongs in this module—only immutable configuration
# values.
# ============================================================================

# ============================================================================
# Border Styles
# ============================================================================
#
# Canonical border style names recognised by Excel/OpenXML.
#
# workbook.py should translate these values directly into the appropriate
# openpyxl Side(style=...) definitions.
# ============================================================================

BORDER_STYLE_NONE: Final[str] = "none"

BORDER_STYLE_HAIR: Final[str] = "hair"

BORDER_STYLE_DOTTED: Final[str] = "dotted"

BORDER_STYLE_DASHED: Final[str] = "dashed"

BORDER_STYLE_THIN: Final[str] = "thin"

BORDER_STYLE_MEDIUM: Final[str] = "medium"

BORDER_STYLE_THICK: Final[str] = "thick"

BORDER_STYLE_DOUBLE: Final[str] = "double"

BORDER_STYLE_DASH_DOT: Final[str] = "dashDot"

BORDER_STYLE_DASH_DOT_DOT: Final[str] = "dashDotDot"

BORDER_STYLE_MEDIUM_DASHED: Final[str] = "mediumDashed"

BORDER_STYLE_MEDIUM_DASH_DOT: Final[str] = "mediumDashDot"

BORDER_STYLE_MEDIUM_DASH_DOT_DOT: Final[str] = "mediumDashDotDot"

BORDER_STYLE_SLANT_DASH_DOT: Final[str] = "slantDashDot"

# Collection of commonly used border styles.
COMMON_BORDER_STYLES: Final[tuple[str, ...]] = (
    BORDER_STYLE_NONE,
    BORDER_STYLE_THIN,
    BORDER_STYLE_MEDIUM,
    BORDER_STYLE_THICK,
    BORDER_STYLE_DOUBLE,
)

# ============================================================================
# Horizontal Alignment
# ============================================================================
#
# Standard horizontal alignment values recognised by Excel.
# ============================================================================

H_ALIGN_GENERAL: Final[str] = "general"

H_ALIGN_LEFT: Final[str] = "left"

H_ALIGN_CENTER: Final[str] = "center"

H_ALIGN_RIGHT: Final[str] = "right"

H_ALIGN_FILL: Final[str] = "fill"

H_ALIGN_JUSTIFY: Final[str] = "justify"

H_ALIGN_CENTER_CONTINUOUS: Final[str] = "centerContinuous"

H_ALIGN_DISTRIBUTED: Final[str] = "distributed"

DEFAULT_HORIZONTAL_ALIGNMENT: Final[str] = H_ALIGN_LEFT

# ============================================================================
# Vertical Alignment
# ============================================================================
#
# Standard vertical alignment values recognised by Excel.
# ============================================================================

V_ALIGN_TOP: Final[str] = "top"

V_ALIGN_CENTER: Final[str] = "center"

V_ALIGN_BOTTOM: Final[str] = "bottom"

V_ALIGN_JUSTIFY: Final[str] = "justify"

V_ALIGN_DISTRIBUTED: Final[str] = "distributed"

DEFAULT_VERTICAL_ALIGNMENT: Final[str] = V_ALIGN_CENTER

# ============================================================================
# Text Wrapping
# ============================================================================
#
# Controls whether long text is wrapped within the cell boundaries.
# ============================================================================

TEXT_WRAP_ENABLED: Final[bool] = True

TEXT_WRAP_DISABLED: Final[bool] = False

DEFAULT_TEXT_WRAP: Final[bool] = TEXT_WRAP_ENABLED

# ============================================================================
# Text Rotation
# ============================================================================
#
# Excel supports text rotation from 0–180 degrees (with some implementation
# restrictions depending on the library).
#
# These values provide sensible defaults used throughout the workbook.
# ============================================================================

TEXT_ROTATION_NONE: Final[int] = 0

TEXT_ROTATION_VERTICAL: Final[int] = 90

DEFAULT_TEXT_ROTATION: Final[int] = TEXT_ROTATION_NONE

# ============================================================================
# Cell Indentation
# ============================================================================
#
# Indentation levels measured in Excel indentation units.
#
# Useful for hierarchical reports and grouped data.
# ============================================================================

INDENT_NONE: Final[int] = 0

INDENT_SMALL: Final[int] = 1

INDENT_MEDIUM: Final[int] = 2

INDENT_LARGE: Final[int] = 3

DEFAULT_INDENT: Final[int] = INDENT_NONE

# ============================================================================
# Reading Order
# ============================================================================
#
# Controls text direction for multilingual workbooks.
#
# Context
#     Excel automatically determines the reading order.
#
# Left-to-Right
#     Suitable for English and most European languages.
#
# Right-to-Left
#     Suitable for Arabic and Hebrew.
# ============================================================================

READING_ORDER_CONTEXT: Final[int] = 0

READING_ORDER_LEFT_TO_RIGHT: Final[int] = 1

READING_ORDER_RIGHT_TO_LEFT: Final[int] = 2

DEFAULT_READING_ORDER: Final[int] = READING_ORDER_LEFT_TO_RIGHT

# ============================================================================
# Cell Protection Defaults
# ============================================================================
#
# Default protection behaviour for workbook cells.
#
# NOTE:
# Cell protection only becomes effective after worksheet protection has been
# enabled.
#
# These defaults provide a consistent starting point while allowing
# workbook.py to override settings for specific worksheets or cells.
# ============================================================================

DEFAULT_CELL_LOCKED: Final[bool] = False

DEFAULT_CELL_HIDDEN: Final[bool] = False

PROTECTED_CELL_LOCKED: Final[bool] = True

FORMULA_CELL_HIDDEN: Final[bool] = True

# ============================================================================
# Excel Named Range Constants
# ============================================================================
#
# This section defines the canonical names of workbook-level Named Ranges.
#
# Named ranges provide a stable mechanism for referring to cells, tables and
# validation lists without relying on absolute worksheet references.
#
# Advantages
# ----------
#
# • Easier worksheet refactoring.
# • Cleaner formulas.
# • Centralised naming convention.
# • Improved readability.
# • Simplified data validation.
# • Reduced hard-coded worksheet references.
#
# Naming Convention
# -----------------
#
# All names:
#
# • Begin with a letter.
# • Contain only letters, numbers and underscores.
# • Contain no spaces.
# • Avoid Excel reserved names.
# • Remain stable even if worksheet names change.
#
# workbook.py should create these named ranges once during workbook
# construction. validation.py and formulas.py should reference these constants
# rather than embedding literal names.
# ============================================================================

# ============================================================================
# Master Data Named Ranges
# ============================================================================
#
# Dynamic lists used throughout the workbook.
# ============================================================================

NR_TEAMS: Final[str] = "Teams"

NR_PLAYERS: Final[str] = "Players"

NR_VENUES: Final[str] = "Venues"

NR_OFFICIALS: Final[str] = "Officials"

NR_TOURNAMENT_FORMATS: Final[str] = "TournamentFormats"

# ============================================================================
# Cricket Lookup Lists
# ============================================================================
#
# Named ranges used by validation.py to populate cricket-specific dropdowns.
# ============================================================================

NR_BATTING_STYLES: Final[str] = "BattingStyles"

NR_BOWLING_STYLES: Final[str] = "BowlingStyles"

NR_PLAYER_ROLES: Final[str] = "PlayerRoles"

NR_DISMISSALS: Final[str] = "Dismissals"

NR_EXTRAS: Final[str] = "Extras"

NR_POWERPLAYS: Final[str] = "Powerplays"

# ============================================================================
# Validation Lists
# ============================================================================
#
# Frequently reused validation sources.
# ============================================================================

NR_YES_NO: Final[str] = "YesNo"

NR_TRUE_FALSE: Final[str] = "TrueFalse"

NR_ENABLED_DISABLED: Final[str] = "EnabledDisabled"

NR_GENDERS: Final[str] = "Genders"

NR_COUNTRIES: Final[str] = "Countries"

NR_STATES: Final[str] = "States"

NR_AGE_GROUPS: Final[str] = "AgeGroups"

NR_COMPETITION_FORMATS: Final[str] = "CompetitionFormats"

NR_BALL_COLOURS: Final[str] = "BallColours"

NR_PITCH_TYPES: Final[str] = "PitchTypes"

NR_WEATHER: Final[str] = "Weather"

NR_GROUND_CONDITIONS: Final[str] = "GroundConditions"

NR_RESULT_METHODS: Final[str] = "ResultMethods"

NR_UMPIRE_ROLES: Final[str] = "UmpireRoles"

# ============================================================================
# Match Configuration Lists
# ============================================================================
#
# Validation ranges related to match administration.
# ============================================================================

NR_TOSS_CHOICES: Final[str] = "TossChoices"

NR_TOSS_DECISIONS: Final[str] = "TossDecisions"

NR_MATCH_STATUS: Final[str] = "MatchStatus"

NR_MATCH_RESULTS: Final[str] = "MatchResults"

NR_POWERPLAY_NAMES: Final[str] = "PowerplayNames"

NR_FIELD_RESTRICTIONS: Final[str] = "FieldRestrictions"

# ============================================================================
# Ball Outcome Lists
# ============================================================================
#
# Validation ranges used while recording deliveries.
# ============================================================================

NR_RUN_VALUES: Final[str] = "RunValues"

NR_BOUNDARY_TYPES: Final[str] = "BoundaryTypes"

NR_EXTRA_TYPES: Final[str] = "ExtraTypes"

NR_NO_BALL_REASONS: Final[str] = "NoBallReasons"

NR_WIDE_REASONS: Final[str] = "WideReasons"

NR_PENALTY_REASONS: Final[str] = "PenaltyReasons"

# ============================================================================
# Convenience Collection
# ============================================================================
#
# Complete list of named ranges expected to exist in the workbook.
#
# workbook.py may iterate over this tuple when validating workbook integrity,
# creating named ranges, or performing diagnostics.
# ============================================================================

ALL_NAMED_RANGES: Final[tuple[str, ...]] = (
    NR_TEAMS,
    NR_PLAYERS,
    NR_VENUES,
    NR_OFFICIALS,
    NR_TOURNAMENT_FORMATS,
    NR_BATTING_STYLES,
    NR_BOWLING_STYLES,
    NR_PLAYER_ROLES,
    NR_DISMISSALS,
    NR_EXTRAS,
    NR_POWERPLAYS,
    NR_YES_NO,
    NR_TRUE_FALSE,
    NR_ENABLED_DISABLED,
    NR_GENDERS,
    NR_COUNTRIES,
    NR_STATES,
    NR_AGE_GROUPS,
    NR_COMPETITION_FORMATS,
    NR_BALL_COLOURS,
    NR_PITCH_TYPES,
    NR_WEATHER,
    NR_GROUND_CONDITIONS,
    NR_RESULT_METHODS,
    NR_UMPIRE_ROLES,
    NR_TOSS_CHOICES,
    NR_TOSS_DECISIONS,
    NR_MATCH_STATUS,
    NR_MATCH_RESULTS,
    NR_POWERPLAY_NAMES,
    NR_FIELD_RESTRICTIONS,
    NR_RUN_VALUES,
    NR_BOUNDARY_TYPES,
    NR_EXTRA_TYPES,
    NR_NO_BALL_REASONS,
    NR_WIDE_REASONS,
    NR_PENALTY_REASONS,
)

# ============================================================================
# Formula Engine Constants
# ============================================================================
#
# This section contains constants used exclusively by formulas.py.
#
# The formulas module performs mathematical calculations for cricket scoring
# and statistics. Centralising these values ensures that every calculation
# throughout the application follows identical assumptions and rounding rules.
#
# Design Principles
# -----------------
#
# • Deterministic calculations.
# • No duplicated numeric literals.
# • Consistent rounding.
# • Easy future customisation.
# • Shared by workbook.py, score_ball.py and reporting modules.
#
# IMPORTANT
# ---------
#
# Several of these values intentionally reference previously defined constants
# (such as BALLS_PER_OVER and DEFAULT_MAX_OVERS) rather than redefining
# identical numeric values. This preserves a true single source of truth.
# ============================================================================

# ============================================================================
# Core Cricket Calculation Constants
# ============================================================================
#
# Re-export frequently used cricket values for clarity within formulas.py.
# These aliases make the intent of calculations more obvious while avoiding
# duplicate numeric literals.
# ============================================================================

FORMULA_BALLS_PER_OVER: Final[int] = BALLS_PER_OVER

FORMULA_MAX_OVERS: Final[int] = DEFAULT_MAX_OVERS

# ============================================================================
# Minimum Overs
# ============================================================================
#
# Smallest valid innings length supported by generic calculations.
#
# Individual tournaments may impose larger minimums.
# ============================================================================

FORMULA_MIN_OVERS: Final[int] = 1

# ============================================================================
# Precision Settings
# ============================================================================
#
# Decimal precision used when returning calculated values.
#
# These settings affect presentation and rounding only—they do not alter the
# underlying mathematical calculations.
# ============================================================================

FORMULA_RUN_RATE_PRECISION: Final[int] = RUN_RATE_PRECISION

FORMULA_STRIKE_RATE_PRECISION: Final[int] = STRIKE_RATE_PRECISION

FORMULA_ECONOMY_PRECISION: Final[int] = ECONOMY_RATE_PRECISION

FORMULA_NET_RUN_RATE_PRECISION: Final[int] = 3

FORMULA_AVERAGE_PRECISION: Final[int] = AVERAGE_PRECISION

# ============================================================================
# Floating-Point Calculation Tolerances
# ============================================================================
#
# Floating-point arithmetic can introduce tiny rounding differences.
#
# These tolerances allow formulas.py to safely compare values without relying
# upon exact binary equality.
#
# SMALL
#     Suitable for most cricket calculations.
#
# MEDIUM
#     Suitable for accumulated statistics.
#
# LARGE
#     Suitable for broad comparison where exact precision is unnecessary.
# ============================================================================

FLOAT_TOLERANCE_SMALL: Final[float] = 1e-9

FLOAT_TOLERANCE_MEDIUM: Final[float] = 1e-6

FLOAT_TOLERANCE_LARGE: Final[float] = 1e-3

DEFAULT_FLOAT_TOLERANCE: Final[float] = FLOAT_TOLERANCE_SMALL

# ============================================================================
# Mathematical Defaults
# ============================================================================
#
# Frequently reused numeric constants improve readability throughout the
# formulas module and eliminate repeated literals.
# ============================================================================

ZERO_RUNS: Final[int] = 0

ZERO_BALLS: Final[int] = 0

ZERO_WICKETS: Final[int] = 0

ONE_OVER_IN_BALLS: Final[int] = BALLS_PER_OVER

# ============================================================================
# Formula Error Values
# ============================================================================
#
# Canonical return values representing common calculation outcomes.
#
# These symbolic constants allow formulas.py to communicate error conditions
# consistently without embedding string literals throughout the codebase.
#
# The formulas module may alternatively choose to raise exceptions; these
# constants remain useful for workbook generation, logging and diagnostics.
# ============================================================================

FORMULA_ERROR_DIVISION_BY_ZERO: Final[str] = "DIVISION_BY_ZERO"

FORMULA_ERROR_INVALID_OVERS: Final[str] = "INVALID_OVERS"

FORMULA_ERROR_INVALID_BALLS: Final[str] = "INVALID_BALLS"

FORMULA_ERROR_INVALID_RUNS: Final[str] = "INVALID_RUNS"

FORMULA_ERROR_NEGATIVE_VALUE: Final[str] = "NEGATIVE_VALUE"

FORMULA_ERROR_INVALID_INPUT: Final[str] = "INVALID_INPUT"

FORMULA_ERROR_INCOMPLETE_INNINGS: Final[str] = "INCOMPLETE_INNINGS"

# ============================================================================
# Formula Sentinel Values
# ============================================================================
#
# Standard default values returned when a meaningful cricket statistic cannot
# yet be calculated.
# ============================================================================

NOT_AVAILABLE_FLOAT: Final[float] = 0.0

NOT_AVAILABLE_INTEGER: Final[int] = 0

NOT_AVAILABLE_TEXT: Final[str] = "N/A"

# ============================================================================
# Worksheet Protection Constants
# ============================================================================
#
# This section defines the default worksheet and cell protection behaviour for
# the Cricket Tournament Workbook.
#
# Purpose
# -------
#
# Excel worksheet protection can selectively prevent users from modifying
# workbook content while still allowing specific interactions such as sorting,
# filtering or selecting unlocked cells.
#
# By centralising these settings:
#
# • workbook.py can apply consistent protection across every worksheet.
# • create_match.py can inherit sensible defaults.
# • score_ball.py can temporarily unlock only required scoring cells.
# • Future worksheet-specific overrides remain simple.
#
# IMPORTANT
# ---------
#
# These are DEFAULTS.
#
# Individual worksheets may override these values where appropriate.
#
# Password management should be handled externally. This module intentionally
# stores only a placeholder string.
#
# No worksheet protection logic belongs in this module.
# ============================================================================

# ============================================================================
# Cell Lock States
# ============================================================================
#
# Convenience constants for readability when applying cell protection.
# ============================================================================

CELL_LOCKED: Final[bool] = True

CELL_UNLOCKED: Final[bool] = False

# ============================================================================
# Formula Visibility
# ============================================================================
#
# Controls whether Excel displays formulas when worksheet protection is
# enabled.
#
# Hidden formulas remain functional but are not visible in the formula bar.
# ============================================================================

FORMULAS_HIDDEN: Final[bool] = True

FORMULAS_VISIBLE: Final[bool] = False

# ============================================================================
# Worksheet Password
# ============================================================================
#
# Placeholder password.
#
# Production applications should inject passwords from configuration rather
# than hard-coding them into source code.
# ============================================================================

DEFAULT_WORKSHEET_PASSWORD: Final[str] = ""

# ============================================================================
# Worksheet Protection Defaults
# ============================================================================
#
# Master switches controlling worksheet protection behaviour.
# ============================================================================

DEFAULT_WORKSHEET_PROTECTION_ENABLED: Final[bool] = True

DEFAULT_CONTENT_PROTECTION: Final[bool] = True

DEFAULT_OBJECT_PROTECTION: Final[bool] = True

DEFAULT_SCENARIO_PROTECTION: Final[bool] = True

# ============================================================================
# Selection Permissions
# ============================================================================
#
# Determines which cells users may select after worksheet protection has been
# enabled.
# ============================================================================

ALLOW_SELECT_LOCKED_CELLS: Final[bool] = False

ALLOW_SELECT_UNLOCKED_CELLS: Final[bool] = True

# ============================================================================
# Editing Permissions
# ============================================================================
#
# Controls whether users may edit various worksheet objects.
#
# The defaults are intentionally conservative for protecting scorecards and
# tournament data.
# ============================================================================

ALLOW_FORMAT_CELLS: Final[bool] = False

ALLOW_FORMAT_COLUMNS: Final[bool] = False

ALLOW_FORMAT_ROWS: Final[bool] = False

ALLOW_INSERT_COLUMNS: Final[bool] = False

ALLOW_INSERT_ROWS: Final[bool] = False

ALLOW_INSERT_HYPERLINKS: Final[bool] = False

ALLOW_DELETE_COLUMNS: Final[bool] = False

ALLOW_DELETE_ROWS: Final[bool] = False

# ============================================================================
# Data Operations
# ============================================================================
#
# These options determine whether users may reorganise worksheet data without
# changing protected content.
# ============================================================================

ALLOW_SORT: Final[bool] = True

ALLOW_AUTO_FILTER: Final[bool] = True

ALLOW_PIVOT_TABLES: Final[bool] = False

# ============================================================================
# Object Interaction
# ============================================================================
#
# Controls interaction with graphical workbook objects.
# ============================================================================

ALLOW_EDIT_OBJECTS: Final[bool] = False

ALLOW_EDIT_SCENARIOS: Final[bool] = False

# ============================================================================
# Allowed Actions
# ============================================================================
#
# Human-readable names describing supported worksheet actions.
#
# These values are useful for logging, diagnostics, configuration screens and
# future user interfaces.
# ============================================================================

WORKSHEET_ALLOWED_ACTIONS: Final[tuple[str, ...]] = (
    "Select Unlocked Cells",
    "Select Locked Cells",
    "Format Cells",
    "Format Columns",
    "Format Rows",
    "Insert Columns",
    "Insert Rows",
    "Insert Hyperlinks",
    "Delete Columns",
    "Delete Rows",
    "Sort",
    "AutoFilter",
    "Pivot Tables",
    "Edit Objects",
    "Edit Scenarios",
)

# ============================================================================
# Recommended Protection Profiles
# ============================================================================
#
# Named protection profiles that may be referenced by workbook.py when
# applying worksheet-specific security policies.
#
# These are descriptive identifiers only. The implementation logic belongs in
# workbook.py.
# ============================================================================

PROTECTION_PROFILE_READ_ONLY: Final[str] = "ReadOnly"

PROTECTION_PROFILE_DATA_ENTRY: Final[str] = "DataEntry"

PROTECTION_PROFILE_SCORING: Final[str] = "Scoring"

PROTECTION_PROFILE_ADMIN: Final[str] = "Administrator"

PROTECTION_PROFILES: Final[tuple[str, ...]] = (
    PROTECTION_PROFILE_READ_ONLY,
    PROTECTION_PROFILE_DATA_ENTRY,
    PROTECTION_PROFILE_SCORING,
    PROTECTION_PROFILE_ADMIN,
)
# ============================================================================
# Application Metadata
# ============================================================================
#
# This section contains descriptive metadata about the Cricket Tournament
# Workbook application itself.
#
# Unlike workbook metadata (title, subject, author, etc.), these constants
# describe the software project that generates and manages the workbook.
#
# These values may be used by:
#
#     • workbook.py
#     • create_match.py
#     • logging
#     • diagnostic reports
#     • About dialogs
#     • generated documentation
#     • future REST APIs
#
# Keeping all metadata in one location provides a single source of truth for
# application identification and version tracking.
#
# IMPORTANT
# ---------
#
# Some of these values intentionally reference constants defined earlier in
# this module (APPLICATION_NAME and __version__) to avoid duplication.
#
# Placeholder values should be replaced during project setup or automated
# build processes.
# ============================================================================

# ============================================================================
# Core Application Metadata
# ============================================================================
#
# Canonical application identity.
#
# These aliases improve readability for modules that specifically require
# metadata without directly referencing the earlier foundational constants.
# ============================================================================

APPLICATION_DISPLAY_NAME: Final[str] = APPLICATION_NAME

APPLICATION_VERSION: Final[str] = __version__

APPLICATION_IDENTIFIER: Final[str] = "cricket-tournament-workbook"

APPLICATION_DESCRIPTION: Final[str] = (
    "A comprehensive Excel-based cricket tournament management, scoring, "
    "statistics and reporting application."
)

# ============================================================================
# Repository Information
# ============================================================================
#
# Source control metadata.
#
# Replace these placeholder values with the actual repository details once the
# project is published.
# ============================================================================

PROJECT_REPOSITORY: Final[str] = "https://github.com/your-organization/your-repository"

PROJECT_REPOSITORY_NAME: Final[str] = "your-repository"

PROJECT_REPOSITORY_BRANCH: Final[str] = "main"

# ============================================================================
# Licensing
# ============================================================================
#
# Placeholder licensing information.
#
# Update these values to match the project's chosen license before public
# release.
# ============================================================================

PROJECT_LICENSE: Final[str] = "License Placeholder"

LICENSE_URL: Final[str] = "https://example.com/license"

# ============================================================================
# Support Information
# ============================================================================
#
# Contact information presented to users in generated workbooks, logs or
# diagnostic reports.
#
# Replace these placeholders with real support channels when deploying the
# application.
# ============================================================================

SUPPORT_EMAIL: Final[str] = "support@example.com"

SUPPORT_WEBSITE: Final[str] = "https://example.com"

ISSUE_TRACKER_URL: Final[str] = "https://example.com/issues"

# ============================================================================
# Copyright
# ============================================================================
#
# Copyright information embedded into generated workbook metadata and future
# application documentation.
# ============================================================================

COPYRIGHT_NOTICE: Final[str] = (
    "Copyright (c) YYYY Your Organization. All rights reserved."
)

COPYRIGHT_OWNER: Final[str] = "Your Organization"

# ============================================================================
# Build Information
# ============================================================================
#
# Build metadata.
#
# Continuous Integration (CI) or release pipelines may replace these values
# automatically during packaging.
# ============================================================================

BUILD_NUMBER: Final[str] = "dev"

BUILD_DATE: Final[str] = "YYYY-MM-DD"

BUILD_COMMIT: Final[str] = "0000000"

BUILD_CONFIGURATION: Final[str] = "Development"

BUILD_PLATFORM: Final[str] = "Python 3.12"

# ============================================================================
# Release Status
# ============================================================================
#
# Indicates the maturity level of the current build.
#
# These values are descriptive only and may be used by future About dialogs,
# logging systems or diagnostics.
# ============================================================================

RELEASE_STATUS: Final[Literal[
    "Development",
    "Alpha",
    "Beta",
    "Release Candidate",
    "Production",
]] = "Development"


# ============================================================================
# Application Limits
# ============================================================================
#
# This section defines the practical upper limits used throughout the Cricket
# Tournament Workbook.
#
# Purpose
# -------
#
# These limits are not intended to represent the absolute capabilities of
# Microsoft Excel. Instead, they define the maximum values that the application
# is designed to support efficiently.
#
# Benefits of Centralising Limits
# -------------------------------
#
# • Prevents "magic numbers" throughout the codebase.
# • Simplifies workbook validation.
# • Enables consistent input validation.
# • Provides predictable workbook performance.
# • Documents application assumptions.
# • Makes future scaling straightforward.
#
# Individual tournaments may use significantly smaller values. These constants
# simply represent the application's supported design limits.
#
# Excel itself supports far larger datasets than those defined here, but very
# large workbooks become slower to generate, validate and calculate.
# ============================================================================

# ============================================================================
# Tournament Limits
# ============================================================================
#
# Maximum number of tournaments expected within a single workbook.
#
# Although the current workbook is primarily intended for one tournament,
# future versions may support multiple tournaments sharing lookup data.
# ============================================================================

MAX_TOURNAMENTS: Final[int] = 10

# ============================================================================
# Team Limits
# ============================================================================
#
# Maximum number of teams.
#
# This comfortably accommodates leagues, pools and knockout tournaments while
# maintaining good workbook performance.
# ============================================================================

MAX_TEAMS: Final[int] = 128

# ============================================================================
# Player Limits
# ============================================================================
#
# Maximum registered players.
#
# Supports large club competitions while remaining well within Excel's
# practical performance limits.
# ============================================================================

MAX_PLAYERS: Final[int] = 5000

# ============================================================================
# Match Officials
# ============================================================================
#
# Includes umpires, referees, scorers and other recognised officials.
# ============================================================================

MAX_OFFICIALS: Final[int] = 500

# ============================================================================
# Venue Limits
# ============================================================================
#
# Maximum number of cricket grounds or venues maintained in the workbook.
# ============================================================================

MAX_VENUES: Final[int] = 250

# ============================================================================
# Match Limits
# ============================================================================
#
# Maximum fixtures recorded in one workbook.
#
# This supports multiple seasons, league stages and knockout competitions.
# ============================================================================

MAX_MATCHES: Final[int] = 5000

# ============================================================================
# Innings Limits
# ============================================================================
#
# Default maximum innings supported for a single match.
#
# This aligns with the application's limited-overs focus while allowing future
# extensions for multi-innings formats.
# ============================================================================

APPLICATION_MAX_INNINGS: Final[int] = MAX_INNINGS

# ============================================================================
# Delivery Limits
# ============================================================================
#
# Maximum legal deliveries stored for a single innings.
#
# Calculated using the application's maximum overs and standard six-ball over.
# This intentionally excludes additional wides and no-balls.
# ============================================================================

MAX_LEGAL_BALLS_PER_INNINGS: Final[int] = (
    DEFAULT_MAX_OVERS * BALLS_PER_OVER
)

# ============================================================================
# Score Limits
# ============================================================================
#
# Practical upper scoring limit for validation purposes.
#
# This is intentionally generous and is not intended to represent a realistic
# cricket score.
# ============================================================================

MAX_TEAM_SCORE: Final[int] = 9999

# ============================================================================
# Partnership Limits
# ============================================================================
#
# Maximum partnerships in a completed innings.
#
# An innings with ten wickets contains eleven batting partnerships,
# including the final unbroken partnership when applicable.
# ============================================================================

MAX_PARTNERSHIPS: Final[int] = 11

# ============================================================================
# Over Limits
# ============================================================================
#
# Practical maximum overs supported by the workbook.
#
# Future tournaments may override DEFAULT_MAX_OVERS while remaining below this
# application-wide ceiling.
# ============================================================================

APPLICATION_MAX_OVERS: Final[int] = 100

APPLICATION_MIN_OVERS: Final[int] = 1

# ============================================================================
# Worksheet Limits
# ============================================================================
#
# Maximum number of worksheets expected within the workbook.
#
# Excel supports significantly more worksheets; this limit exists to maintain
# a manageable workbook structure.
# ============================================================================

MAX_WORKSHEETS: Final[int] = 100

# ============================================================================
# Workbook Size Assumptions
# ============================================================================
#
# These values document the workbook size assumptions used during design.
#
# They are not enforced by Excel itself but provide useful guidance for
# validation, diagnostics and future optimisation.
# ============================================================================

ASSUMED_MAX_WORKBOOK_SIZE_MB: Final[int] = 50

ASSUMED_MAX_ROWS_PER_WORKSHEET: Final[int] = 50000

ASSUMED_MAX_COLUMNS_PER_WORKSHEET: Final[int] = 100

ASSUMED_MAX_NAMED_RANGES: Final[int] = 500

ASSUMED_MAX_DATA_VALIDATIONS: Final[int] = 10000

# ============================================================================
# Performance Thresholds
# ============================================================================
#
# Soft thresholds that may be used by future diagnostics to warn users when a
# workbook is approaching sizes where performance could begin to degrade.
# ============================================================================

PERFORMANCE_WARNING_MATCHES: Final[int] = 2500

PERFORMANCE_WARNING_PLAYERS: Final[int] = 2500

PERFORMANCE_WARNING_FILE_SIZE_MB: Final[int] = 25

# ============================================================================
# Public Module Interface
# ============================================================================
#
# This final section defines the public API exported by constants.py.
#
# Rather than manually maintaining a potentially very large tuple of exported
# names, the module automatically exports every documented constant while
# excluding imported symbols and private implementation names.
#
# Export Rules
# ------------
#
# A name is exported if:
#
# • It does not begin with "_".
# • It is not an imported typing helper.
# • It is not a module object.
# • It is not one of the standard module metadata variables.
#
# This approach ensures:
#
# • Newly added constants automatically become part of the public API.
# • No maintenance is required when constants are added or removed.
# • All project modules receive a consistent interface.
#
# Future developers should continue documenting every constant where it is
# declared. This section intentionally performs no validation or runtime logic
# beyond constructing the public export list.
# ============================================================================

# Module metadata and imported helper names that should not be exported.
_EXCLUDED_EXPORTS: Final[frozenset[str]] = frozenset(
    {
        "__all__",
        "__author__",
        "__builtins__",
        "__cached__",
        "__doc__",
        "__file__",
        "__loader__",
        "__name__",
        "__package__",
        "__spec__",
        "__version__",
        "annotations",
        "Final",
        "Literal",
    }
)

# Public API
#
# Export every documented constant defined in this module while excluding
# implementation details and imported helper symbols.
__all__: tuple[str, ...] = tuple(
    sorted(
        name
        for name in globals()
        if not name.startswith("_")
        and name not in _EXCLUDED_EXPORTS
    )
)
