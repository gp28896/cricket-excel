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

WORKBOOK_TITLE: Final[str] = "Cricket Tournament Manager"

WORKBOOK_DESCRIPTION: Final[str] = (
    "An Excel-based cricket tournament management system supporting "
    "tournaments, teams, players, fixtures, live scoring, scorecards, "
    "statistics, and reporting."
)

WORKBOOK_AUTHOR: Final[str] = "Your Name"

WORKBOOK_COMPANY: Final[str] = "Your Organization"

WORKBOOK_SUBJECT: Final[str] = "Cricket Tournament Management"

WORKBOOK_CATEGORY: Final[str] = "Sports Management"

WORKBOOK_KEYWORDS: Final[str] = (
    "cricket, tournament, scoring, scorebook, excel, sports, statistics"
)

# ============================================================================
# Default Workbook File
# ============================================================================

DEFAULT_WORKBOOK_FILENAME: Final[str] = "Cricket_Tournament_Manager.xlsx"

# ============================================================================
# Workbook Appearance
# ============================================================================

WORKBOOK_THEME_NAME: Final[str] = "Default"

# ============================================================================
# Workbook Protection Defaults
# ============================================================================

DEFAULT_WORKBOOK_STRUCTURE_PROTECTION: Final[bool] = False

DEFAULT_WORKBOOK_WINDOWS_PROTECTION: Final[bool] = False

DEFAULT_WORKBOOK_PASSWORD: Final[str] = ""

# ============================================================================
# Freeze Pane Defaults
# ============================================================================

DEFAULT_FREEZE_PANE: Final[str] = "A1"

# ============================================================================
# Auto-Fit Defaults
# ============================================================================

DEFAULT_AUTO_FIT_COLUMNS: Final[bool] = True

DEFAULT_AUTO_FIT_ROWS: Final[bool] = False

# ============================================================================
# Excel Limits
# ============================================================================

MAX_SHEET_NAME_LENGTH: Final[int] = 31

# ============================================================================
# Worksheet Names
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

SHEET_LEADERBOARD: Final[str] = SHEET_STANDINGS

DEFAULT_LEADERBOARD_SHEET_NAME: Final[str] = SHEET_STANDINGS

SHEET_SETTINGS: Final[str] = "Settings"

SHEET_VALIDATION: Final[str] = "Validation"

SHEET_HIDDEN_LISTS: Final[str] = "Hidden Lists"

SHEET_NAVIGATION: Final[str] = "Navigation"

SHEET_REPORTS: Final[str] = "Reports"

# ============================================================================
# Workbook Worksheet Order
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

DEFAULT_WORKSHEET_ORDER: Final[tuple[str, ...]] = WORKSHEET_ORDER

# ============================================================================
# Visible Worksheets
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

HIDDEN_SHEETS: Final[tuple[str, ...]] = (
    SHEET_VALIDATION,
    SHEET_HIDDEN_LISTS,
)

# ============================================================================
# Cricket Match Configuration
# ============================================================================

# ============================================================================
# Innings Configuration
# ============================================================================

MAX_INNINGS: Final[int] = 2

TEST_MAX_INNINGS: Final[int] = 4

# ============================================================================
# Over Configuration
# ============================================================================

DEFAULT_MAX_OVERS: Final[int] = 20

ODI_MAX_OVERS: Final[int] = 50

TEST_MAX_OVERS: Final[int] = 450

# ============================================================================
# Ball Configuration
# ============================================================================

BALLS_PER_OVER: Final[int] = 6

# ============================================================================
# Wicket Configuration
# ============================================================================

MAX_WICKETS: Final[int] = 10

# ============================================================================
# Toss Options
# ============================================================================

TOSS_CHOICES: Final[tuple[str, ...]] = (
    "Heads",
    "Tails",
)

# ============================================================================
# Toss Decision Options
# ============================================================================

TOSS_DECISION_BAT: Final[str] = "Bat"

TOSS_DECISION_BOWL: Final[str] = "Bowl"

TOSS_DECISION_CHOICES: Final[tuple[str, ...]] = (
    TOSS_DECISION_BAT,
    TOSS_DECISION_BOWL,
)

VALID_TOSS_DECISIONS: Final[tuple[str, ...]] = (
    TOSS_DECISION_BAT,
    TOSS_DECISION_BOWL,
    "bat",
    "bowl",
)

# ============================================================================
# Match Result Types
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

SUPER_OVER_ENABLED: Final[bool] = True

ENABLE_SUPER_OVER: Final[bool] = SUPER_OVER_ENABLED

SUPER_OVER_MAX_OVERS: Final[int] = 1

SUPER_OVER_MAX_WICKETS: Final[int] = 2

SUPER_OVER_BALLS_PER_OVER: Final[int] = 6

# ============================================================================
# Rain Interruption Configuration
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

FIELD_RESTRICTION_NAMES: Final[tuple[str, ...]] = (
    "Powerplay",
    "Normal Field",
    "Death Overs",
)

# ============================================================================
# Match Status Values
# ============================================================================

MATCH_STATUS_SCHEDULED: Final[str] = "Scheduled"
MATCH_STATUS_TOSS_PENDING: Final[str] = "Toss Pending"
MATCH_STATUS_LIVE: Final[str] = "Live"
MATCH_STATUS_FIRST_INNINGS: Final[str] = "First Innings"
MATCH_STATUS_INNINGS_COMPLETE: Final[str] = "Innings Complete"
MATCH_STATUS_INNINGS_BREAK: Final[str] = "Innings Break"
MATCH_STATUS_SECOND_INNINGS: Final[str] = "Second Innings"
MATCH_STATUS_SUPER_OVER: Final[str] = "Super Over"
MATCH_STATUS_COMPLETED: Final[str] = "Completed"
MATCH_STATUS_ABANDONED: Final[str] = "Abandoned"
MATCH_STATUS_NO_RESULT: Final[str] = "No Result"
MATCH_STATUS_CANCELLED: Final[str] = "Cancelled"

MATCH_STATUS_VALUES: Final[tuple[str, ...]] = (
    MATCH_STATUS_SCHEDULED,
    MATCH_STATUS_TOSS_PENDING,
    MATCH_STATUS_LIVE,
    MATCH_STATUS_FIRST_INNINGS,
    MATCH_STATUS_INNINGS_COMPLETE,
    MATCH_STATUS_INNINGS_BREAK,
    MATCH_STATUS_SECOND_INNINGS,
    MATCH_STATUS_SUPER_OVER,
    MATCH_STATUS_COMPLETED,
    MATCH_STATUS_ABANDONED,
    MATCH_STATUS_NO_RESULT,
    MATCH_STATUS_CANCELLED,
)

# ============================================================================
# Player Configuration
# ============================================================================

# ============================================================================
# Player Roles
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

TEAM_LEADERSHIP_ROLES: Final[tuple[str, ...]] = (
    "Captain",
    "Vice Captain",
)

# ============================================================================
# Captaincy Labels
# ============================================================================

CAPTAIN_LABEL: Final[str] = "Captain"

VICE_CAPTAIN_LABEL: Final[str] = "Vice Captain"

# ============================================================================
# Wicket Keeper Labels
# ============================================================================

WICKET_KEEPER_LABEL: Final[str] = "Wicket Keeper"

# ============================================================================
# Batting Styles
# ============================================================================

BATTING_STYLES: Final[tuple[str, ...]] = (
    "Right-hand Bat",
    "Left-hand Bat",
)

# ============================================================================
# Bowling Styles
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

BOWLING_ARMS: Final[tuple[str, ...]] = (
    "Right-arm",
    "Left-arm",
    "None",
)

# ============================================================================
# Player Status
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

# ============================================================================
# Delivery Types
# ============================================================================
#
# Reconciles delivery names across:
#   • formulas.py   ("Normal", "Bye", "Leg Bye", "Wide", "No Ball")
#   • score_ball.py ("LEGAL", "NO_BALL", "WIDE", "DEAD_BALL")
#   • validation.py ("Legal Delivery", "No Ball", "Wide")
# ============================================================================

DELIVERY_LEGAL: Final[str] = "Legal Delivery"
DELIVERY_NORMAL: Final[str] = "Normal"
DELIVERY_BYE: Final[str] = "Bye"
DELIVERY_LEG_BYE: Final[str] = "Leg Bye"
DELIVERY_WIDE: Final[str] = "Wide"
DELIVERY_NO_BALL: Final[str] = "No Ball"
DELIVERY_DEAD_BALL: Final[str] = "Dead Ball"

LEGAL_DELIVERY_DISPLAY_TYPES: Final[tuple[str, ...]] = (
    DELIVERY_LEGAL,
    DELIVERY_NORMAL,
    DELIVERY_BYE,
    DELIVERY_LEG_BYE,
)

ILLEGAL_DELIVERY_DISPLAY_TYPES: Final[tuple[str, ...]] = (
    DELIVERY_NO_BALL,
    DELIVERY_WIDE,
)

LEGAL_DELIVERIES: Final[tuple[str, ...]] = (
    "LEGAL",
    DELIVERY_LEGAL,
    DELIVERY_NORMAL,
    DELIVERY_BYE,
    DELIVERY_LEG_BYE,
)

ILLEGAL_DELIVERIES: Final[tuple[str, ...]] = (
    "NO_BALL",
    "WIDE",
    DELIVERY_NO_BALL,
    DELIVERY_WIDE,
)

DEAD_BALL_DELIVERIES: Final[tuple[str, ...]] = (
    "DEAD_BALL",
    DELIVERY_DEAD_BALL,
)

LEGAL_DELIVERY_TYPES: Final[frozenset[str]] = frozenset(LEGAL_DELIVERIES)

ILLEGAL_DELIVERY_TYPES: Final[frozenset[str]] = frozenset(ILLEGAL_DELIVERIES)

ALL_DELIVERY_TYPES: Final[frozenset[str]] = (
    LEGAL_DELIVERY_TYPES | ILLEGAL_DELIVERY_TYPES
)

# ============================================================================
# Ball Run Values & Named Run Outcomes
# ============================================================================

DOT_BALL: Final[int] = 0

SINGLE: Final[int] = 1

DOUBLE: Final[int] = 2

TRIPLE: Final[int] = 3

FOUR: Final[int] = 4

FIVE: Final[int] = 5

SIX: Final[int] = 6

MAX_RUNS_OFF_BAT: Final[int] = SIX

RUN_VALUES: Final[tuple[int, ...]] = (
    DOT_BALL,
    SINGLE,
    DOUBLE,
    TRIPLE,
    FOUR,
    FIVE,
    SIX,
)

# ============================================================================
# Boundary Values & Names
# ============================================================================

BOUNDARY_RUN_VALUES: Final[tuple[int, ...]] = (
    FOUR,
    SIX,
)

BOUNDARY_TYPES: Final[tuple[str, ...]] = (
    "Four",
    "Six",
)

# ============================================================================
# Frequently Referenced Dismissal Constants
# ============================================================================

DISMISSAL_NONE: Final[str] = "None"

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

DISMISSAL_RETIRED_HURT: Final[str] = "Retired Hurt"

DISMISSAL_HANDLED_BALL: Final[str] = "Handled the Ball"

# ============================================================================
# Dismissal Types
# ============================================================================
#
# DISMISSAL_DISPLAY_TYPES provides the canonical Title Case labels for Excel
# dropdowns and scorecards.
#
# DISMISSAL_ENUM_NAMES provides the UPPER_SNAKE_CASE enum identifiers used by
# WicketType / DismissalKind / DismissalType in score_ball.py.
#
# DISMISSAL_TYPES includes both sets so validation in score_ball.py and
# worksheet dropdown helpers both succeed.
# ============================================================================

DISMISSAL_DISPLAY_TYPES: Final[tuple[str, ...]] = (
    DISMISSAL_BOWLED,
    DISMISSAL_CAUGHT,
    DISMISSAL_CAUGHT_AND_BOWLED,
    DISMISSAL_LBW,
    DISMISSAL_RUN_OUT,
    DISMISSAL_STUMPED,
    DISMISSAL_HIT_WICKET,
    DISMISSAL_HIT_BALL_TWICE,
    DISMISSAL_OBSTRUCTING_FIELD,
    DISMISSAL_TIMED_OUT,
    DISMISSAL_RETIRED_OUT,
    DISMISSAL_RETIRED_HURT,
    DISMISSAL_HANDLED_BALL,
)

DISMISSAL_ENUM_NAMES: Final[tuple[str, ...]] = (
    "NONE",
    "BOWLED",
    "CAUGHT",
    "CAUGHT_AND_BOWLED",
    "LBW",
    "RUN_OUT",
    "STUMPED",
    "HIT_WICKET",
    "HIT_BALL_TWICE",
    "HIT_THE_BALL_TWICE",
    "OBSTRUCTING_FIELD",
    "OBSTRUCTING_THE_FIELD",
    "TIMED_OUT",
    "RETIRED_OUT",
    "RETIRED_HURT",
    "HANDLED_BALL",
    "HANDLED_THE_BALL",
)

DISMISSAL_TYPES: Final[tuple[str, ...]] = (
    DISMISSAL_NONE,
    *DISMISSAL_DISPLAY_TYPES,
    *DISMISSAL_ENUM_NAMES,
)

# ============================================================================
# Extras Configuration
# ============================================================================

EXTRA_NONE: Final[str] = "None"

EXTRA_WIDE: Final[str] = "Wide"

EXTRA_NO_BALL: Final[str] = "No Ball"

EXTRA_BYE: Final[str] = "Bye"

EXTRA_LEG_BYE: Final[str] = "Leg Bye"

EXTRA_PENALTY: Final[str] = "Penalty"

# Canonical list of all user-selectable extra categories for dropdowns.
EXTRA_CLASSIFICATIONS: Final[tuple[str, ...]] = (
    EXTRA_WIDE,
    EXTRA_NO_BALL,
    EXTRA_BYE,
    EXTRA_LEG_BYE,
    EXTRA_PENALTY,
)

EXTRA_ENUM_NAMES: Final[tuple[str, ...]] = (
    "NONE",
    "WIDE",
    "NO_BALL",
    "BYE",
    "LEG_BYE",
    "PENALTY",
)

# Complete extra type lookup supporting both display labels and enum names
# checked by score_ball.validate_extra_type().
EXTRA_TYPES: Final[tuple[str, ...]] = (
    EXTRA_NONE,
    *EXTRA_CLASSIFICATIONS,
    *EXTRA_ENUM_NAMES,
)

# ============================================================================
# Wide Reasons
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

# ============================================================================
# Generic Boolean Choices
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

GENDER_OPTIONS: Final[tuple[str, ...]] = (
    "Male",
    "Female",
    "Other",
    "Prefer Not to Say",
)

# ============================================================================
# Countries & States / Provinces
# ============================================================================

COUNTRY_OPTIONS: Final[tuple[str, ...]] = (
    "Select Country",
)

STATE_OPTIONS: Final[tuple[str, ...]] = (
    "Select State / Province",
)

# ============================================================================
# Age Groups
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

COMPETITION_FORMAT_OPTIONS: Final[tuple[str, ...]] = (
    "T5",
    "T10",
    "T20",
    "T50",
    "ODI",
    "One Day",
    "Multi-Day",
    "Test",
    "Test Match",
    "Super Over",
    "Custom",
)

# ============================================================================
# Ball Colours
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

# ============================================================================
# Row Heights
# ============================================================================

DEFAULT_ROW_HEIGHT: Final[float] = 18.0

HEADER_ROW_HEIGHT: Final[float] = 24.0

TITLE_ROW_HEIGHT: Final[float] = 30.0

# ============================================================================
# Default Column Widths
# ============================================================================

DEFAULT_COLUMN_WIDTH: Final[float] = 15.0

NARROW_COLUMN_WIDTH: Final[float] = 8.0

MEDIUM_COLUMN_WIDTH: Final[float] = 15.0

WIDE_COLUMN_WIDTH: Final[float] = 22.0

EXTRA_WIDE_COLUMN_WIDTH: Final[float] = 35.0

# ============================================================================
# Numeric Precision
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

EXCEL_DATE_FORMAT: Final[str] = "yyyy-mm-dd"

EXCEL_TIME_FORMAT: Final[str] = "hh:mm:ss"

EXCEL_DATETIME_FORMAT: Final[str] = "yyyy-mm-dd hh:mm:ss"

# ============================================================================
# Standard Excel Number Formats
# ============================================================================

NUMBER_FORMAT_GENERAL: Final[str] = "General"

NUMBER_FORMAT_INTEGER: Final[str] = "0"

NUMBER_FORMAT_DECIMAL: Final[str] = "0.00"

NUMBER_FORMAT_THOUSANDS: Final[str] = "#,##0"

NUMBER_FORMAT_THOUSANDS_DECIMAL: Final[str] = "#,##0.00"

NUMBER_FORMAT_PERCENTAGE: Final[str] = "0.00%"

NUMBER_FORMAT_CURRENCY: Final[str] = "₹#,##0.00"

NUMBER_FORMAT_TEXT: Final[str] = "@"

NUMBER_FORMAT_OVERS: Final[str] = "0.0"

# ============================================================================
# Print Margins
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

PAPER_SIZE_A4: Final[str] = "A4"

PAPER_SIZE_A3: Final[str] = "A3"

PAPER_SIZE_LETTER: Final[str] = "Letter"

DEFAULT_PAPER_SIZE: Final[str] = PAPER_SIZE_A4

# ============================================================================
# Page Orientation
# ============================================================================

PAGE_ORIENTATION_PORTRAIT: Final[str] = "portrait"

PAGE_ORIENTATION_LANDSCAPE: Final[str] = "landscape"

DEFAULT_PAGE_ORIENTATION: Final[str] = PAGE_ORIENTATION_LANDSCAPE

# ============================================================================
# Workbook Colour Palette
# ============================================================================

# ============================================================================
# Header Colours
# ============================================================================

COLOUR_HEADER_BACKGROUND: Final[str] = "1F4E78"      # Dark Blue

COLOUR_HEADER_TEXT: Final[str] = "FFFFFF"            # White

# ============================================================================
# Table Colours
# ============================================================================

COLOUR_TABLE_HEADER: Final[str] = "D9EAD3"

COLOUR_TABLE_BORDER: Final[str] = "B7B7B7"

COLOUR_TABLE_BACKGROUND: Final[str] = "FFFFFF"

# ============================================================================
# Alternating Row Colours
# ============================================================================

COLOUR_ROW_ODD: Final[str] = "FFFFFF"

COLOUR_ROW_EVEN: Final[str] = "F5F5F5"

# ============================================================================
# Status Colours
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

COLOUR_EDITABLE_CELL: Final[str] = "FFF2CC"

COLOUR_LOCKED_CELL: Final[str] = "E6E6E6"

COLOUR_DISABLED_CELL: Final[str] = "D9D9D9"

# ============================================================================
# Navigation Colours
# ============================================================================

COLOUR_NAVIGATION_BACKGROUND: Final[str] = "D9EAD3"

COLOUR_NAVIGATION_TEXT: Final[str] = "274E13"

COLOUR_NAVIGATION_HIGHLIGHT: Final[str] = "6AA84F"

# ============================================================================
# Tournament Branding
# ============================================================================

COLOUR_BRAND_PRIMARY: Final[str] = "0B5394"

COLOUR_BRAND_SECONDARY: Final[str] = "6AA84F"

COLOUR_BRAND_ACCENT: Final[str] = "F1C232"

# ============================================================================
# Neutral Greys
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

# ============================================================================
# Border Styles
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

V_ALIGN_TOP: Final[str] = "top"

V_ALIGN_CENTER: Final[str] = "center"

V_ALIGN_BOTTOM: Final[str] = "bottom"

V_ALIGN_JUSTIFY: Final[str] = "justify"

V_ALIGN_DISTRIBUTED: Final[str] = "distributed"

DEFAULT_VERTICAL_ALIGNMENT: Final[str] = V_ALIGN_CENTER

# ============================================================================
# Text Wrapping
# ============================================================================

TEXT_WRAP_ENABLED: Final[bool] = True

TEXT_WRAP_DISABLED: Final[bool] = False

DEFAULT_TEXT_WRAP: Final[bool] = TEXT_WRAP_ENABLED

# ============================================================================
# Text Rotation
# ============================================================================

TEXT_ROTATION_NONE: Final[int] = 0

TEXT_ROTATION_VERTICAL: Final[int] = 90

DEFAULT_TEXT_ROTATION: Final[int] = TEXT_ROTATION_NONE

# ============================================================================
# Cell Indentation
# ============================================================================

INDENT_NONE: Final[int] = 0

INDENT_SMALL: Final[int] = 1

INDENT_MEDIUM: Final[int] = 2

INDENT_LARGE: Final[int] = 3

DEFAULT_INDENT: Final[int] = INDENT_NONE

# ============================================================================
# Reading Order
# ============================================================================

READING_ORDER_CONTEXT: Final[int] = 0

READING_ORDER_LEFT_TO_RIGHT: Final[int] = 1

READING_ORDER_RIGHT_TO_LEFT: Final[int] = 2

DEFAULT_READING_ORDER: Final[int] = READING_ORDER_LEFT_TO_RIGHT

# ============================================================================
# Cell Protection Defaults
# ============================================================================

DEFAULT_CELL_LOCKED: Final[bool] = False

DEFAULT_CELL_HIDDEN: Final[bool] = False

PROTECTED_CELL_LOCKED: Final[bool] = True

FORMULA_CELL_HIDDEN: Final[bool] = True

# ============================================================================
# Excel Named Range Constants
# ============================================================================

# ============================================================================
# Master Data Named Ranges
# ============================================================================

NR_TEAMS: Final[str] = "Teams"

NR_PLAYERS: Final[str] = "Players"

NR_VENUES: Final[str] = "Venues"

NR_OFFICIALS: Final[str] = "Officials"

NR_TOURNAMENT_FORMATS: Final[str] = "TournamentFormats"

# ============================================================================
# Cricket Lookup Lists
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

NR_TOSS_CHOICES: Final[str] = "TossChoices"

NR_TOSS_DECISIONS: Final[str] = "TossDecisions"

NR_MATCH_STATUS: Final[str] = "MatchStatus"

NR_MATCH_RESULTS: Final[str] = "MatchResults"

NR_POWERPLAY_NAMES: Final[str] = "PowerplayNames"

NR_FIELD_RESTRICTIONS: Final[str] = "FieldRestrictions"

# ============================================================================
# Ball Outcome Lists
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

FORMULA_BALLS_PER_OVER: Final[int] = BALLS_PER_OVER

FORMULA_MAX_OVERS: Final[int] = DEFAULT_MAX_OVERS

FORMULA_MIN_OVERS: Final[int] = 1

FORMULA_RUN_RATE_PRECISION: Final[int] = RUN_RATE_PRECISION

FORMULA_STRIKE_RATE_PRECISION: Final[int] = STRIKE_RATE_PRECISION

FORMULA_ECONOMY_PRECISION: Final[int] = ECONOMY_RATE_PRECISION

FORMULA_NET_RUN_RATE_PRECISION: Final[int] = 3

FORMULA_AVERAGE_PRECISION: Final[int] = AVERAGE_PRECISION

FLOAT_TOLERANCE_SMALL: Final[float] = 1e-9

FLOAT_TOLERANCE_MEDIUM: Final[float] = 1e-6

FLOAT_TOLERANCE_LARGE: Final[float] = 1e-3

DEFAULT_FLOAT_TOLERANCE: Final[float] = FLOAT_TOLERANCE_SMALL

ZERO_RUNS: Final[int] = 0

ZERO_BALLS: Final[int] = 0

ZERO_WICKETS: Final[int] = 0

ONE_OVER_IN_BALLS: Final[int] = BALLS_PER_OVER

FORMULA_ERROR_DIVISION_BY_ZERO: Final[str] = "DIVISION_BY_ZERO"

FORMULA_ERROR_INVALID_OVERS: Final[str] = "INVALID_OVERS"

FORMULA_ERROR_INVALID_BALLS: Final[str] = "INVALID_BALLS"

FORMULA_ERROR_INVALID_RUNS: Final[str] = "INVALID_RUNS"

FORMULA_ERROR_NEGATIVE_VALUE: Final[str] = "NEGATIVE_VALUE"

FORMULA_ERROR_INVALID_INPUT: Final[str] = "INVALID_INPUT"

FORMULA_ERROR_INCOMPLETE_INNINGS: Final[str] = "INCOMPLETE_INNINGS"

NOT_AVAILABLE_FLOAT: Final[float] = 0.0

NOT_AVAILABLE_INTEGER: Final[int] = 0

NOT_AVAILABLE_TEXT: Final[str] = "N/A"

# ============================================================================
# Leaderboard / Points Table Defaults
# ============================================================================

DEFAULT_POINTS_FOR_WIN: Final[int] = 2

DEFAULT_POINTS_FOR_TIE: Final[int] = 1

DEFAULT_POINTS_FOR_NO_RESULT: Final[int] = 1

DEFAULT_POINTS_FOR_LOSS: Final[int] = 0

DEFAULT_NRR_DECIMAL_PLACES: Final[int] = FORMULA_NET_RUN_RATE_PRECISION

DEFAULT_RANK_START: Final[int] = 1

DEFAULT_MAX_DISPLAY_TEAMS: Final[int] = 1000

# ============================================================================
# Worksheet Protection Constants
# ============================================================================

CELL_LOCKED: Final[bool] = True

CELL_UNLOCKED: Final[bool] = False

FORMULAS_HIDDEN: Final[bool] = True

FORMULAS_VISIBLE: Final[bool] = False

DEFAULT_WORKSHEET_PASSWORD: Final[str] = ""

DEFAULT_WORKSHEET_PROTECTION_ENABLED: Final[bool] = True

DEFAULT_CONTENT_PROTECTION: Final[bool] = True

DEFAULT_OBJECT_PROTECTION: Final[bool] = True

DEFAULT_SCENARIO_PROTECTION: Final[bool] = True

ALLOW_SELECT_LOCKED_CELLS: Final[bool] = False

ALLOW_SELECT_UNLOCKED_CELLS: Final[bool] = True

ALLOW_FORMAT_CELLS: Final[bool] = False

ALLOW_FORMAT_COLUMNS: Final[bool] = False

ALLOW_FORMAT_ROWS: Final[bool] = False

ALLOW_INSERT_COLUMNS: Final[bool] = False

ALLOW_INSERT_ROWS: Final[bool] = False

ALLOW_INSERT_HYPERLINKS: Final[bool] = False

ALLOW_DELETE_COLUMNS: Final[bool] = False

ALLOW_DELETE_ROWS: Final[bool] = False

ALLOW_SORT: Final[bool] = True

ALLOW_AUTO_FILTER: Final[bool] = True

ALLOW_PIVOT_TABLES: Final[bool] = False

ALLOW_EDIT_OBJECTS: Final[bool] = False

ALLOW_EDIT_SCENARIOS: Final[bool] = False

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
# Core Application Identity & Build Metadata
# ============================================================================

APPLICATION_DISPLAY_NAME: Final[str] = APPLICATION_NAME

APPLICATION_VERSION: Final[str] = __version__

APPLICATION_IDENTIFIER: Final[str] = "cricket-tournament-workbook"

APPLICATION_DESCRIPTION: Final[str] = (
    "A comprehensive Excel-based cricket tournament management, scoring, "
    "statistics and reporting application."
)

PROJECT_REPOSITORY: Final[str] = (
    "https://github.com/your-organization/your-repository"
)

PROJECT_REPOSITORY_NAME: Final[str] = "your-repository"

PROJECT_REPOSITORY_BRANCH: Final[str] = "main"

PROJECT_LICENSE: Final[str] = "License Placeholder"

LICENSE_URL: Final[str] = "https://example.com/license"

SUPPORT_EMAIL: Final[str] = "support@example.com"

SUPPORT_WEBSITE: Final[str] = "https://example.com"

ISSUE_TRACKER_URL: Final[str] = "https://example.com/issues"

COPYRIGHT_NOTICE: Final[str] = (
    "Copyright (c) YYYY Your Organization. All rights reserved."
)

COPYRIGHT_OWNER: Final[str] = "Your Organization"

BUILD_NUMBER: Final[str] = "dev"

BUILD_DATE: Final[str] = "YYYY-MM-DD"

BUILD_COMMIT: Final[str] = "0000000"

BUILD_CONFIGURATION: Final[str] = "Development"

BUILD_PLATFORM: Final[str] = "Python 3.12"

RELEASE_STATUS: Final[
    Literal[
        "Development",
        "Alpha",
        "Beta",
        "Release Candidate",
        "Production",
    ]
] = "Development"

# ============================================================================
# Application Limits
# ============================================================================

MAX_TOURNAMENTS: Final[int] = 10

MAX_TEAMS: Final[int] = 128

MAX_PLAYERS: Final[int] = 5000

MAX_OFFICIALS: Final[int] = 500

MAX_VENUES: Final[int] = 250

MAX_MATCHES: Final[int] = 5000

APPLICATION_MAX_INNINGS: Final[int] = MAX_INNINGS

MAX_LEGAL_BALLS_PER_INNINGS: Final[int] = (
    DEFAULT_MAX_OVERS * BALLS_PER_OVER
)

MAX_TEAM_SCORE: Final[int] = 9999

MAX_PARTNERSHIPS: Final[int] = 11

APPLICATION_MAX_OVERS: Final[int] = 100

APPLICATION_MIN_OVERS: Final[int] = 1

MAX_WORKSHEETS: Final[int] = 100

ASSUMED_MAX_WORKBOOK_SIZE_MB: Final[int] = 50

ASSUMED_MAX_ROWS_PER_WORKSHEET: Final[int] = 50000

ASSUMED_MAX_COLUMNS_PER_WORKSHEET: Final[int] = 100

ASSUMED_MAX_NAMED_RANGES: Final[int] = 500

ASSUMED_MAX_DATA_VALIDATIONS: Final[int] = 10000

PERFORMANCE_WARNING_MATCHES: Final[int] = 2500

PERFORMANCE_WARNING_PLAYERS: Final[int] = 2500

PERFORMANCE_WARNING_FILE_SIZE_MB: Final[int] = 25

# ============================================================================
# Public Module Interface
# ============================================================================

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

__all__: tuple[str, ...] = tuple(
    sorted(
        name
        for name in globals()
        if not name.startswith("_")
        and name not in _EXCLUDED_EXPORTS
    )
)
