"""
create_match.py
================

Module responsible for generating and initializing every worksheet required
for a single cricket match within the tournament workbook.

Overview
--------
This module acts as the match-generation layer of the application.

Unlike ``workbook.py``, which provides reusable workbook construction
utilities, this module is responsible only for assembling the worksheets
needed to represent one cricket match.

Its responsibilities include:

* Creating all match-related worksheets.
* Naming worksheets consistently.
* Initializing sheet metadata.
* Connecting worksheets through workbook helper APIs.
* Populating default structures.
* Delegating styling, validation, formulas, and formatting to the
  project's shared infrastructure.

This module intentionally does **not** implement:

* Cell styling
* Named styles
* Conditional formatting
* Data validation rules
* Formula generation
* Workbook-wide configuration
* Print settings
* Freeze panes
* Column sizing
* Page layouts

Those responsibilities belong exclusively to:

* ``workbook.py``
* ``styles.py``
* ``validation.py``
* ``formulas.py``

Design Philosophy
-----------------
The overall architecture follows a layered approach.

::

    create_match.py
            │
            ▼
      WorkbookBuilder
            │
            ▼
    workbook.py public API
            │
     ┌──────┼────────┐
     ▼      ▼        ▼
 styles  validation formulas

This separation keeps the project maintainable and avoids duplicated logic.

Typical Workflow
----------------

1. Receive a WorkbookBuilder instance.
2. Receive match metadata.
3. Create required worksheets.
4. Initialize worksheet structures.
5. Delegate workbook configuration to WorkbookBuilder.
6. Return references to the generated worksheets.

Notes
-----
This module is intentionally "thin".

Business rules for workbook construction should remain inside
``workbook.py`` so that all workbook generation throughout the project
shares identical behavior.

Future Parts
------------
Subsequent sections of this module will implement:

* Match configuration models
* Worksheet naming helpers
* Match sheet creation
* Scorebook initialization
* Playing XI sheets
* Innings sheets
* Bowling cards
* Batting cards
* Extras sheets
* Fall-of-wickets sheets
* Partnership sheets
* Wagon wheel placeholders
* Statistics sheets
* Match summary sheets
* Result sheets
* Worksheet registration
* Public orchestration APIs

This file intentionally contains no worksheet creation logic in Part 1.
"""

from __future__ import annotations

###############################################################################
# Standard Library Imports
###############################################################################

import logging
from collections.abc import Mapping, MutableMapping, Sequence
from pathlib import Path
from typing import (
    Any,
    Final,
    Literal,
    NewType,
    TypeAlias,
)

###############################################################################
# Third-Party Imports
###############################################################################

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

###############################################################################
# Internal Project Imports
###############################################################################

from constants import *
from formulas import *
from styles import *
from validation import *
from workbook import WorkbookBuilder

###############################################################################
# Logging Configuration
###############################################################################

logger = logging.getLogger(__name__)

###############################################################################
# Type Aliases
###############################################################################

#: Dictionary of worksheet objects keyed by worksheet name.
WorksheetMap: TypeAlias = MutableMapping[str, Worksheet]

#: Immutable mapping of worksheet names to worksheet objects.
WorksheetMapping: TypeAlias = Mapping[str, Worksheet]

#: Ordered collection of worksheets.
WorksheetSequence: TypeAlias = Sequence[Worksheet]

#: Match identifier.
MatchId = NewType("MatchId", str)

#: Tournament identifier.
TournamentId = NewType("TournamentId", str)

#: Team identifier.
TeamId = NewType("TeamId", str)

#: Worksheet role used internally while generating sheets.
WorksheetRole: TypeAlias = Literal[
    "summary",
    "scorecard",
    "innings",
    "batting",
    "bowling",
    "extras",
    "fow",
    "partnerships",
    "statistics",
    "result",
    "metadata",
]

###############################################################################
# Internal Constants
###############################################################################

#: Name used by the module logger.
_LOGGER_NAME: Final[str] = __name__

#: Default worksheet title prefix.
_DEFAULT_MATCH_PREFIX: Final[str] = "Match"

#: Maximum worksheet title length supported by Excel.
_EXCEL_MAX_SHEET_NAME_LENGTH: Final[int] = 31

#: Default encoding used for exported text (future use).
_DEFAULT_ENCODING: Final[str] = "utf-8"

#: Current module version.
_MODULE_VERSION: Final[str] = "1.0.0"

###############################################################################
# Public Exports
###############################################################################

__all__: list[str] = []

###############################################################################
# Match Configuration Model
###############################################################################

from dataclasses import dataclass, field
from datetime import date, time


@dataclass(slots=True, kw_only=True)
class MatchConfig:
    """Immutable configuration describing a single cricket match.

    This dataclass contains all metadata required to create and initialize
    the worksheets representing one cricket match. It intentionally stores
    configuration only and performs no workbook or worksheet manipulation.

    Validation is performed during initialization to ensure invalid match
    configurations are detected before workbook generation begins.

    Attributes:
        match_type:
            Human-readable match type (for example ``"T20"``,
            ``"ODI"``, ``"Test"``, or ``"Custom"``).

        overs:
            Scheduled overs per innings. Must be greater than zero.

        innings_count:
            Number of innings in the match.

            Typical values:

            * 2 — Limited-overs cricket
            * 4 — Test cricket

        venue:
            Venue or ground where the match will be played.

        match_date:
            Scheduled match date.

        start_time:
            Scheduled local start time.

        tournament:
            Tournament or competition name.

        home_team:
            Name of the home team.

        away_team:
            Name of the away team.

        toss_winner:
            Team that won the toss.

            May be ``None`` if the toss has not yet taken place.

        toss_decision:
            Toss decision.

            Supported values are:

            * ``"bat"``
            * ``"bowl"``

            May be ``None`` before the toss.

        powerplay_enabled:
            Indicates whether powerplays are used.

        powerplay_overs:
            Ordered sequence defining the number of overs in each powerplay
            phase. Empty if powerplays are disabled.

        dls_enabled:
            Whether Duckworth-Lewis-Stern calculations are enabled.

        super_over_enabled:
            Whether a Super Over may be played.

    Raises:
        ValueError:
            If any supplied configuration is invalid.
    """

    match_type: str

    overs: int

    innings_count: int

    venue: str

    match_date: date

    start_time: time

    tournament: str

    home_team: str

    away_team: str

    toss_winner: str | None = None

    toss_decision: Literal["bat", "bowl"] | None = None

    powerplay_enabled: bool = True

    powerplay_overs: tuple[int, ...] = field(default_factory=tuple)

    dls_enabled: bool = False

    super_over_enabled: bool = False

    def __post_init__(self) -> None:
        """Validate the supplied configuration."""

        self._validate_match_type()
        self._validate_overs()
        self._validate_innings()
        self._validate_teams()
        self._validate_strings()
        self._validate_toss()
        self._validate_powerplays()

    ###########################################################################
    # Validation Helpers
    ###########################################################################

    def _validate_match_type(self) -> None:
        """Validate the configured match type."""

        if not self.match_type.strip():
            raise ValueError("match_type cannot be empty.")

    def _validate_overs(self) -> None:
        """Validate scheduled overs."""

        if self.overs <= 0:
            raise ValueError("overs must be greater than zero.")

    def _validate_innings(self) -> None:
        """Validate innings count."""

        if self.innings_count not in (2, 4):
            raise ValueError(
                "innings_count must be either 2 (limited overs) "
                "or 4 (Test cricket)."
            )

    def _validate_strings(self) -> None:
        """Validate required text fields."""

        required = {
            "venue": self.venue,
            "tournament": self.tournament,
            "home_team": self.home_team,
            "away_team": self.away_team,
        }

        for field_name, value in required.items():
            if not value.strip():
                raise ValueError(f"{field_name} cannot be empty.")

    def _validate_teams(self) -> None:
        """Validate participating teams."""

        if self.home_team.strip() == self.away_team.strip():
            raise ValueError(
                "home_team and away_team must be different."
            )

    def _validate_toss(self) -> None:
        """Validate toss configuration."""

        if self.toss_winner is None and self.toss_decision is not None:
            raise ValueError(
                "toss_decision cannot be specified without toss_winner."
            )

        if self.toss_winner is not None:
            if self.toss_winner not in {
                self.home_team,
                self.away_team,
            }:
                raise ValueError(
                    "toss_winner must be either home_team or away_team."
                )

    def _validate_powerplays(self) -> None:
        """Validate powerplay configuration."""

        if not self.powerplay_enabled:
            if self.powerplay_overs:
                raise ValueError(
                    "powerplay_overs must be empty when "
                    "powerplay_enabled is False."
                )
            return

        if any(over <= 0 for over in self.powerplay_overs):
            raise ValueError(
                "All powerplay overs must be greater than zero."
            )

        if sum(self.powerplay_overs) > self.overs:
            raise ValueError(
                "Total powerplay overs cannot exceed scheduled overs."
            )


###############################################################################
# Public Exports
###############################################################################

__all__.append("MatchConfig")

###############################################################################
# Match Worksheet Builder
###############################################################################


class MatchSheetBuilder:
    """Builder responsible for generating worksheets for a single cricket match.

    This class represents the orchestration layer for match-specific workbook
    generation. It owns the configuration for one match and coordinates calls
    to :class:`WorkbookBuilder` without duplicating any workbook logic.

    Responsibilities:
        * Store validated match configuration.
        * Maintain references to workbook infrastructure.
        * Track worksheets created for the match.
        * Provide internal helper state for subsequent worksheet generation.
        * Delegate workbook operations to :class:`WorkbookBuilder`.

    Non-Responsibilities:
        * Workbook styling.
        * Cell formatting.
        * Data validation.
        * Formula creation.
        * Workbook-wide configuration.
        * Worksheet implementation details owned by ``workbook.py``.

    Notes:
        Actual worksheet creation is intentionally implemented in later parts
        of this module. This class currently initializes only the internal
        state required by the builder.
    """

    ###########################################################################
    # Construction
    ###########################################################################

    def __init__(
        self,
        workbook_builder: WorkbookBuilder,
        config: MatchConfig,
    ) -> None:
        """Initialize a new match worksheet builder.

        Args:
            workbook_builder:
                Initialized :class:`WorkbookBuilder` responsible for all
                workbook-level operations.

            config:
                Validated configuration describing the cricket match.

        Raises:
            TypeError:
                If ``workbook_builder`` or ``config`` are of unexpected types.
        """
        if not isinstance(workbook_builder, WorkbookBuilder):
            raise TypeError(
                "workbook_builder must be an instance of WorkbookBuilder."
            )

        if not isinstance(config, MatchConfig):
            raise TypeError(
                "config must be an instance of MatchConfig."
            )

        self._builder: WorkbookBuilder = workbook_builder
        self._config: MatchConfig = config

        self._logger: logging.Logger = logger.getChild(
            self.__class__.__name__
        )

        self._worksheets: WorksheetMap = {}

        self._worksheet_order: list[str] = []

        self._metadata: dict[str, Any] = {}

        self._initialize_helpers()

        self._logger.debug(
            "Initialized MatchSheetBuilder for %s vs %s.",
            self._config.home_team,
            self._config.away_team,
        )

    ###########################################################################
    # Internal Initialization
    ###########################################################################

    def _initialize_helpers(self) -> None:
        """Initialize internal helper state.

        This method prepares reusable containers and caches that will be
        populated during worksheet generation in subsequent parts of the
        module.

        No worksheets are created here.
        """
        self._sheet_roles: dict[str, WorksheetRole] = {}

        self._named_ranges: dict[str, str] = {}

        self._creation_sequence: list[str] = []

        self._initialized: bool = True

    ###########################################################################
    # Read-Only Properties
    ###########################################################################

    @property
    def workbook_builder(self) -> WorkbookBuilder:
        """Return the associated workbook builder."""
        return self._builder

    @property
    def config(self) -> MatchConfig:
        """Return the match configuration."""
        return self._config

    @property
    def logger(self) -> logging.Logger:
        """Return the builder-specific logger."""
        return self._logger

    @property
    def worksheet_registry(self) -> WorksheetMapping:
        """Return a read-only mapping of registered worksheets.

        The registry is initially empty and will be populated as worksheets
        are created in later parts of this module.
        """
        return self._worksheets


    ###########################################################################
    # Worksheet Planning
    ###########################################################################

    def plan_match_sheets(self) -> tuple[str, ...]:
        """Plan the complete worksheet structure for this match.

        This method determines every worksheet that will eventually be created
        for the match. It performs planning only—no worksheets are created and
        the workbook remains unchanged.

        Planning consists of four stages:

        1. Generate worksheet names.
        2. Validate worksheet names.
        3. Register worksheet creation order.
        4. Return the finalized plan.

        Returns:
            Ordered tuple of worksheet names representing the worksheet
            creation plan.
        """
        sheet_names = self.generate_sheet_names()

        self.validate_sheet_names(sheet_names)

        self.register_sheet_order(sheet_names)

        return tuple(self._worksheet_order)

    def generate_sheet_names(self) -> list[str]:
        """Generate the logical worksheet names for this match.

        The generated names are deterministic so that every match workbook
        follows an identical worksheet layout.

        Supported worksheets include:

        * Match Summary
        * Playing XI
        * Scorecard
        * Ball Log
        * Batting
        * Bowling
        * Extras
        * Partnerships
        * Fall of Wickets
        * DLS
        * Statistics
        * Hidden Data

        Returns:
            Ordered list of worksheet names.
        """
        prefix = f"{_DEFAULT_MATCH_PREFIX} {self._config.home_team} vs {self._config.away_team}"

        return [
            f"{prefix} - Summary",
            f"{prefix} - Playing XI",
            f"{prefix} - Scorecard",
            f"{prefix} - Ball Log",
            f"{prefix} - Batting",
            f"{prefix} - Bowling",
            f"{prefix} - Extras",
            f"{prefix} - Partnerships",
            f"{prefix} - Fall of Wickets",
            f"{prefix} - DLS",
            f"{prefix} - Statistics",
            f"{prefix} - Hidden Data",
        ]

    def validate_sheet_names(
        self,
        sheet_names: Sequence[str],
    ) -> None:
        """Validate the planned worksheet names.

        Validation performed:

        * Worksheet name is not empty.
        * Worksheet name length does not exceed Excel limits.
        * Worksheet names are unique.

        Args:
            sheet_names:
                Planned worksheet names.

        Raises:
            ValueError:
                If any worksheet name is invalid.
        """
        seen: set[str] = set()

        for sheet_name in sheet_names:
            if not sheet_name.strip():
                raise ValueError("Worksheet names cannot be empty.")

            if len(sheet_name) > _EXCEL_MAX_SHEET_NAME_LENGTH:
                raise ValueError(
                    f"Worksheet name exceeds Excel's "
                    f"{_EXCEL_MAX_SHEET_NAME_LENGTH}-character limit: "
                    f"{sheet_name!r}"
                )

            if sheet_name in seen:
                raise ValueError(
                    f"Duplicate worksheet name detected: {sheet_name!r}"
                )

            seen.add(sheet_name)

    def register_sheet_order(
        self,
        sheet_names: Sequence[str],
    ) -> None:
        """Register the planned worksheet creation order.

        This method records the logical worksheet sequence for later worksheet
        creation. No workbook modifications are performed.

        Args:
            sheet_names:
                Ordered worksheet names.
        """
        self._worksheet_order.clear()
        self._creation_sequence.clear()
        self._sheet_roles.clear()

        role_map: tuple[WorksheetRole, ...] = (
            "summary",
            "metadata",
            "scorecard",
            "metadata",
            "batting",
            "bowling",
            "extras",
            "partnerships",
            "fow",
            "metadata",
            "statistics",
            "metadata",
        )

        for sheet_name, role in zip(sheet_names, role_map, strict=True):
            self._worksheet_order.append(sheet_name)
            self._creation_sequence.append(sheet_name)
            self._sheet_roles[sheet_name] = role

        self._logger.debug(
            "Planned %d worksheets for match '%s vs %s'.",
            len(self._worksheet_order),
            self._config.home_team,
            self._config.away_team,
        )

    ###########################################################################
    # Worksheet Creation
    ###########################################################################

    def create_planned_worksheets(self) -> WorksheetMapping:
        """Create every worksheet defined by the worksheet plan.

        This method is responsible only for orchestrating worksheet creation.
        All workbook manipulation is delegated to :class:`WorkbookBuilder`
        through its public API.

        Workflow:

        1. Ensure a worksheet plan exists.
        2. Create each worksheet using WorkbookBuilder.
        3. Register each worksheet.
        4. Apply default workbook settings.
        5. Return the worksheet registry.

        Returns:
            Read-only mapping of worksheet names to worksheet objects.

        Raises:
            RuntimeError:
                If worksheet creation fails.
        """
        if not self._worksheet_order:
            self.plan_match_sheets()

        self._logger.info(
            "Creating %d worksheets.",
            len(self._worksheet_order),
        )

        for sheet_name in self._worksheet_order:
            worksheet = self._create_single_worksheet(sheet_name)
            self._register_worksheet(sheet_name, worksheet)

        self._apply_default_workbook_settings()

        self._logger.info(
            "Successfully created %d worksheets.",
            len(self._worksheets),
        )

        return self.worksheet_registry

    ###########################################################################
    # Internal Worksheet Helpers
    ###########################################################################

    def _create_single_worksheet(
        self,
        sheet_name: str,
    ) -> Worksheet:
        """Create a single worksheet using ``WorkbookBuilder``.

        This helper intentionally delegates worksheet creation entirely to the
        public API exposed by ``WorkbookBuilder``.

        Args:
            sheet_name:
                Name of the worksheet to create.

        Returns:
            Newly created worksheet.

        Raises:
            RuntimeError:
                If the workbook builder cannot create the worksheet.
        """
        worksheet = self._builder.create_worksheet(title=sheet_name)

        if worksheet is None:
            raise RuntimeError(
                f"WorkbookBuilder failed to create worksheet "
                f"{sheet_name!r}."
            )

        return worksheet

    def _register_worksheet(
        self,
        sheet_name: str,
        worksheet: Worksheet,
    ) -> None:
        """Register a created worksheet.

        Registration stores worksheet references for later initialization.

        Args:
            sheet_name:
                Worksheet title.

            worksheet:
                Newly created worksheet.
        """
        self._worksheets[sheet_name] = worksheet

        self._logger.debug(
            "Registered worksheet '%s'.",
            sheet_name,
        )

    ###########################################################################
    # Workbook Initialization
    ###########################################################################

    def _apply_default_workbook_settings(self) -> None:
        """Apply default workbook configuration.

        This method intentionally performs no styling, formatting,
        validation, or formula generation.

        All configuration is delegated to the public APIs exposed by
        :class:`WorkbookBuilder`.
        """
        self._builder.apply_default_workbook_settings()

        self._logger.debug(
            "Applied default workbook settings."
        )

    ###########################################################################
    # Match Summary Worksheet
    ###########################################################################

    def create_match_summary_sheet(self) -> Worksheet:
        """Create and initialize the Match Summary worksheet.

        This method creates the worksheet (if necessary) and populates the
        standard match metadata template.

        Workbook formatting, styling, validation, formulas, page setup and
        sizing are intentionally delegated to ``WorkbookBuilder``.

        Returns:
            The initialized Match Summary worksheet.
        """
        sheet_name = next(
            (
                name
                for name in self._worksheet_order
                if name.endswith("Summary")
            ),
            None,
        )

        if sheet_name is None:
            raise RuntimeError(
                "Match Summary worksheet has not been planned."
            )

        worksheet = self._worksheets.get(sheet_name)

        if worksheet is None:
            worksheet = self._create_single_worksheet(sheet_name)
            self._register_worksheet(sheet_name, worksheet)

        self._populate_match_summary(worksheet)

        self._builder.initialize_worksheet(worksheet)

        return worksheet

    ###########################################################################
    # Internal Helpers
    ###########################################################################

    def _populate_match_summary(
        self,
        worksheet: Worksheet,
    ) -> None:
        """Populate the Match Summary worksheet.

        The layout contains only workbook data. Presentation concerns remain
        the responsibility of ``WorkbookBuilder``.

        Args:
            worksheet:
                Match Summary worksheet.
        """
        rows: tuple[tuple[str, Any], ...] = (
            ("Tournament", self._config.tournament),
            ("Match Number", ""),
            ("Venue", self._config.venue),
            ("Date", self._config.match_date),
            ("Start Time", self._config.start_time),
            ("Home Team", self._config.home_team),
            ("Away Team", self._config.away_team),
            (
                "Toss Winner",
                self._config.toss_winner or "",
            ),
            (
                "Toss Decision",
                self._config.toss_decision or "",
            ),
            ("Result", ""),
            ("Winning Team", ""),
            ("Margin", ""),
            ("Match Status", ""),
            ("Umpire 1", ""),
            ("Umpire 2", ""),
            ("Third Umpire", ""),
            ("Match Referee", ""),
            ("Scorer 1", ""),
            ("Scorer 2", ""),
            ("Weather", ""),
            ("Temperature", ""),
            ("Conditions", ""),
            ("Pitch", ""),
            ("Outfield", ""),
            ("Player of the Match", ""),
        )

        start_row = 1

        for index, (label, value) in enumerate(rows):
            row = start_row + index

            self._builder.write_cell(
                worksheet=worksheet,
                row=row,
                column=1,
                value=label,
            )

            self._builder.write_cell(
                worksheet=worksheet,
                row=row,
                column=2,
                value=value,
            )

        self._logger.debug(
            "Initialized Match Summary worksheet '%s'.",
            worksheet.title,
        )


    ###########################################################################
    # Playing XI Worksheet
    ###########################################################################

    def create_playing_xi_sheet(self) -> Worksheet:
        """Create and initialize the Playing XI worksheet.

        The worksheet captures the team sheets for both participating teams.
        Only worksheet structure and default values are populated here.

        Styling, formatting, validation, named ranges and formulas are
        delegated entirely to :class:`WorkbookBuilder`.

        Returns:
            The initialized Playing XI worksheet.
        """
        sheet_name = next(
            (
                name
                for name in self._worksheet_order
                if name.endswith("Playing XI")
            ),
            None,
        )

        if sheet_name is None:
            raise RuntimeError(
                "Playing XI worksheet has not been planned."
            )

        worksheet = self._worksheets.get(sheet_name)

        if worksheet is None:
            worksheet = self._create_single_worksheet(sheet_name)
            self._register_worksheet(sheet_name, worksheet)

        self._populate_playing_xi(worksheet)

        self._builder.initialize_worksheet(worksheet)

        return worksheet

    ###########################################################################
    # Internal Helpers
    ###########################################################################

    def _populate_playing_xi(
        self,
        worksheet: Worksheet,
    ) -> None:
        """Populate the Playing XI worksheet.

        The worksheet is divided into two mirrored sections:

        * Team A
        * Team B

        Each section contains placeholders for:

        * Playing XI
        * Captain
        * Vice Captain
        * Wicketkeeper
        * Substitutes
        * Impact Player
        * Reserve Players
        * Coach
        * Manager

        Args:
            worksheet:
                Playing XI worksheet.
        """
        self._populate_team_roster(
            worksheet=worksheet,
            start_column=1,
            team_name=self._config.home_team,
        )

        self._populate_team_roster(
            worksheet=worksheet,
            start_column=5,
            team_name=self._config.away_team,
        )

        self._logger.debug(
            "Initialized Playing XI worksheet '%s'.",
            worksheet.title,
        )

    def _populate_team_roster(
        self,
        *,
        worksheet: Worksheet,
        start_column: int,
        team_name: str,
    ) -> None:
        """Populate one team's roster section.

        Args:
            worksheet:
                Target worksheet.

            start_column:
                Starting column for the team block.

            team_name:
                Team name.
        """
        self._builder.write_cell(
            worksheet=worksheet,
            row=1,
            column=start_column,
            value=team_name,
        )

        self._builder.write_cell(
            worksheet=worksheet,
            row=2,
            column=start_column,
            value="Playing XI",
        )

        for index in range(11):
            self._builder.write_cell(
                worksheet=worksheet,
                row=3 + index,
                column=start_column,
                value=index + 1,
            )

            self._builder.write_cell(
                worksheet=worksheet,
                row=3 + index,
                column=start_column + 1,
                value="",
            )

        metadata = (
            ("Captain", 15),
            ("Vice Captain", 16),
            ("Wicketkeeper", 17),
            ("Impact Player", 18),
            ("Coach", 19),
            ("Manager", 20),
        )

        for label, row in metadata:
            self._builder.write_cell(
                worksheet=worksheet,
                row=row,
                column=start_column,
                value=label,
            )

            self._builder.write_cell(
                worksheet=worksheet,
                row=row,
                column=start_column + 1,
                value="",
            )

        self._builder.write_cell(
            worksheet=worksheet,
            row=22,
            column=start_column,
            value="Substitutes",
        )

        for row in range(23, 28):
            self._builder.write_cell(
                worksheet=worksheet,
                row=row,
                column=start_column + 1,
                value="",
            )

        self._builder.write_cell(
            worksheet=worksheet,
            row=30,
            column=start_column,
            value="Reserve Players",
        )

        for row in range(31, 36):
            self._builder.write_cell(
                worksheet=worksheet,
                row=row,
                column=start_column + 1,
                value="",
            )

    ###########################################################################
    # Innings Scorecard Worksheet
    ###########################################################################

    def create_scorecard_sheet(self) -> Worksheet:
        """Create and initialize the innings scorecard worksheet.

        The worksheet provides the primary scoring interface for a single
        innings. It creates only the worksheet structure and delegates all
        workbook operations to :class:`WorkbookBuilder`.

        Formula generation is delegated entirely to the public API exposed by
        ``formulas.py`` through ``WorkbookBuilder``. This module never embeds
        Excel formulas directly.

        Returns:
            Initialized scorecard worksheet.
        """
        sheet_name = next(
            (
                name
                for name in self._worksheet_order
                if name.endswith("Scorecard")
            ),
            None,
        )

        if sheet_name is None:
            raise RuntimeError(
                "Scorecard worksheet has not been planned."
            )

        worksheet = self._worksheets.get(sheet_name)

        if worksheet is None:
            worksheet = self._create_single_worksheet(sheet_name)
            self._register_worksheet(sheet_name, worksheet)

        self._populate_scorecard(worksheet)

        self._builder.initialize_worksheet(worksheet)

        return worksheet

    ###########################################################################
    # Internal Helpers
    ###########################################################################

    def _populate_scorecard(
        self,
        worksheet: Worksheet,
    ) -> None:
        """Populate the scorecard worksheet.

        Sections created:

        * Batting
        * Bowling
        * Extras
        * Totals
        * Run Rate
        * Required Run Rate
        * Overs
        * Target

        Formula placement is delegated to ``WorkbookBuilder`` using reusable
        helpers backed by ``formulas.py``.

        Args:
            worksheet:
                Scorecard worksheet.
        """
        self._builder.write_cell(
            worksheet=worksheet,
            row=1,
            column=1,
            value="Batting",
        )

        self._builder.write_cell(
            worksheet=worksheet,
            row=1,
            column=10,
            value="Bowling",
        )

        batting_headers = (
            "Batter",
            "Dismissal",
            "R",
            "B",
            "4s",
            "6s",
            "SR",
        )

        bowling_headers = (
            "Bowler",
            "O",
            "M",
            "R",
            "W",
            "Econ",
        )

        for column, header in enumerate(batting_headers, start=1):
            self._builder.write_cell(
                worksheet=worksheet,
                row=2,
                column=column,
                value=header,
            )

        for column, header in enumerate(bowling_headers, start=10):
            self._builder.write_cell(
                worksheet=worksheet,
                row=2,
                column=column,
                value=header,
            )

        summary_labels = (
            "Extras",
            "Total",
            "Overs",
            "Run Rate",
            "Required Run Rate",
            "Target",
        )

        start_row = 18

        for offset, label in enumerate(summary_labels):
            row = start_row + offset

            self._builder.write_cell(
                worksheet=worksheet,
                row=row,
                column=1,
                value=label,
            )

            self._builder.bind_formula(
                worksheet=worksheet,
                row=row,
                column=2,
                formula_key=label.lower().replace(" ", "_"),
            )

        self._logger.debug(
            "Initialized Scorecard worksheet '%s'.",
            worksheet.title,
        )

    ###########################################################################
    # Ball Log Worksheet
    ###########################################################################

    def create_ball_log_sheet(self) -> Worksheet:
        """Create and initialize the Ball Log worksheet.

        The Ball Log worksheet records every delivery bowled during the match.
        Each row represents a single delivery and contains sufficient
        information to reconstruct the innings and derive scorecards,
        partnerships, statistics and other reports.

        Workbook formatting, styling, data validation and formulas are
        delegated entirely to :class:`WorkbookBuilder`.

        Returns:
            Initialized Ball Log worksheet.
        """
        sheet_name = next(
            (
                name
                for name in self._worksheet_order
                if name.endswith("Ball Log")
            ),
            None,
        )

        if sheet_name is None:
            raise RuntimeError(
                "Ball Log worksheet has not been planned."
            )

        worksheet = self._worksheets.get(sheet_name)

        if worksheet is None:
            worksheet = self._create_single_worksheet(sheet_name)
            self._register_worksheet(sheet_name, worksheet)

        self._populate_ball_log(worksheet)

        self._builder.initialize_worksheet(worksheet)

        return worksheet

    ###########################################################################
    # Internal Helpers
    ###########################################################################

    def _populate_ball_log(
        self,
        worksheet: Worksheet,
    ) -> None:
        """Populate the Ball Log worksheet.

        The worksheet consists of a header row followed by placeholder rows
        for delivery-by-delivery scoring.

        Args:
            worksheet:
                Ball Log worksheet.
        """
        headers = (
            "Over",
            "Ball",
            "Batter",
            "Bowler",
            "Runs",
            "Extras",
            "Dismissal",
            "Comment",
            "Legal Delivery",
            "Match State",
        )

        for column, header in enumerate(headers, start=1):
            self._builder.write_cell(
                worksheet=worksheet,
                row=1,
                column=column,
                value=header,
            )

        # Reserve rows for ball-by-ball scoring.
        #
        # The default allocation is intentionally generous and can be expanded
        # later if required. No workbook formatting or validation is applied
        # here.
        first_data_row = 2
        last_data_row = 1000

        for row in range(first_data_row, last_data_row + 1):
            for column in range(1, len(headers) + 1):
                self._builder.write_cell(
                    worksheet=worksheet,
                    row=row,
                    column=column,
                    value="",
                )

        self._logger.debug(
            "Initialized Ball Log worksheet '%s' with %d delivery rows.",
            worksheet.title,
            last_data_row - first_data_row + 1,
        )

