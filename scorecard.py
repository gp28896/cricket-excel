"""
scorecard.py
=============

Utilities for populating cricket scorecards inside an existing tournament
workbook.

Overview
--------
This module is responsible for writing all batting, bowling, innings and
match-summary scorecards into worksheets that have already been created by
``create_match.py``.

Unlike ``create_match.py``, this module **never creates worksheets**.
Its sole responsibility is to populate cells, formulas, summaries and
statistics within an existing workbook structure.

The module is intentionally designed to be independent of match simulation
or scoring logic. It receives already-created worksheet objects together with
structured score data and renders those values into the workbook using the
shared helper APIs provided by ``workbook.py``.

Responsibilities
----------------
The responsibilities of this module include:

* Populating batting scorecards.
* Populating bowling scorecards.
* Writing innings summaries.
* Writing extras.
* Writing fall-of-wickets.
* Writing partnership summaries.
* Writing batting statistics.
* Writing bowling statistics.
* Writing run-rate summaries.
* Writing innings metadata.
* Writing match summary sections.
* Applying workbook helper utilities where appropriate.
* Preserving workbook formatting created during worksheet generation.

Non-Responsibilities
--------------------
This module deliberately does **not** perform any of the following tasks:

* Creating worksheets.
* Creating workbook styles.
* Creating named ranges.
* Applying workbook validation.
* Performing ball-by-ball scoring.
* Simulating cricket matches.
* Tournament leaderboard calculations.
* Workbook generation.

Those responsibilities belong to:

* ``create_match.py``
* ``styles.py``
* ``validation.py``
* ``score_ball.py``
* ``leaderboard.py``
* ``workbook.py``

Architecture
------------
The module is intended to remain highly modular.

::

    score_ball.py
            │
            ▼
      structured score data
            │
            ▼
        scorecard.py
            │
            ▼
      workbook.py helper APIs
            │
            ▼
      Existing worksheet objects

The implementation philosophy is:

1. Never directly manipulate workbook structure.
2. Prefer reusable helper functions.
3. Keep formatting decisions inside workbook helpers.
4. Keep score rendering deterministic.
5. Separate batting, bowling and innings rendering into independent sections.
6. Minimize duplicated Excel-writing logic.
7. Keep public APIs stable for future GUI/web implementations.

Thread Safety
-------------
The module is not intended to be thread-safe.

Compatibility
-------------
Python:
    3.12+

Dependencies:
    * openpyxl
    * workbook.py
    * constants.py
    * formulas.py

Style Guide
-----------
* Google-style docstrings.
* Extensive type hints.
* Pure rendering logic.
* Minimal side effects.
* Deterministic output.

Typical Usage
-------------
Example::

    from scorecard import populate_batting_scorecard

    populate_batting_scorecard(
        worksheet=batting_sheet,
        innings=innings_data,
    )

Public API
----------
The complete public API will be exported after implementation of all module
sections.
"""

from __future__ import annotations

###############################################################################
# Standard Library Imports
###############################################################################

import logging
from collections.abc import Mapping, MutableMapping, Sequence
from pathlib import Path
from typing import Any, Final, Literal, TypeAlias

###############################################################################
# Third-Party Imports
###############################################################################

from openpyxl.cell import Cell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook.workbook import Workbook

###############################################################################
# Internal Project Imports
###############################################################################

from constants import *
from formulas import *
from workbook import *

###############################################################################
# Logging Configuration
###############################################################################

logger = logging.getLogger(__name__)

###############################################################################
# Type Aliases
###############################################################################

WorksheetMap: TypeAlias = Mapping[str, Worksheet]
MutableWorksheetMap: TypeAlias = MutableMapping[str, Worksheet]

CellValue: TypeAlias = (
    str
    | int
    | float
    | bool
    | None
)

RowValues: TypeAlias = Sequence[CellValue]

PlayerRecord: TypeAlias = Mapping[str, Any]
BowlerRecord: TypeAlias = Mapping[str, Any]
InningsRecord: TypeAlias = Mapping[str, Any]
MatchRecord: TypeAlias = Mapping[str, Any]

###############################################################################
# Internal Constants
###############################################################################

MODULE_NAME: Final[str] = "scorecard"

LOGGER_NAME: Final[str] = __name__

SUPPORTED_INNINGS: Final[tuple[int, ...]] = (1, 2)

DEFAULT_ALIGNMENT = Alignment(horizontal="center", vertical="center")

DEFAULT_THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

###############################################################################
# Public Exports
###############################################################################

__all__: list[str] = []

###############################################################################
# Scorecard Configuration
###############################################################################

from dataclasses import dataclass


@dataclass(slots=True, kw_only=True)
class ScorecardConfig:
    """Immutable configuration describing an innings scorecard.

    This dataclass encapsulates all match- and innings-level metadata required
    by future scorecard rendering functions. It intentionally contains **only**
    configuration values and performs no worksheet manipulation.

    The configuration object is expected to be created by higher-level modules
    after match creation and before scorecard population.

    Attributes:
        innings_number:
            Innings being rendered. Supported values are 1 and 2.

        batting_team:
            Name of the batting team.

        bowling_team:
            Name of the bowling team.

        target:
            Target score for the innings.

            For first innings this is typically ``None``.
            For chase innings this is the number of runs required to win.

        max_overs:
            Maximum scheduled overs for the innings.

            Examples:
                * 20
                * 50
                * 90

            Fractional overs are not permitted.

        wickets:
            Maximum wickets available.

            Usually 10.

        dls_enabled:
            Indicates whether the Duckworth-Lewis-Stern method applies.

        super_over:
            True if this innings belongs to a Super Over.

        follow_on:
            True when the innings is played under follow-on conditions.

        declaration:
            True if the batting side declared its innings.

        match_type:
            Human-readable match format.

            Typical values include:

            * "T20"
            * "ODI"
            * "TEST"
            * "FIRST_CLASS"
            * "LIST_A"
            * "T10"
    """

    innings_number: int
    batting_team: str
    bowling_team: str

    target: int | None = None
    max_overs: int = 20
    wickets: int = 10

    dls_enabled: bool = False
    super_over: bool = False
    follow_on: bool = False
    declaration: bool = False

    match_type: str = "T20"

    def __post_init__(self) -> None:
        """Validate configuration values.

        Raises:
            TypeError:
                If any field has an unexpected type.

            ValueError:
                If any value is outside the supported range.
        """

        if not isinstance(self.innings_number, int):
            raise TypeError("innings_number must be an integer.")

        if self.innings_number not in SUPPORTED_INNINGS:
            raise ValueError(
                f"innings_number must be one of {SUPPORTED_INNINGS}."
            )

        if not isinstance(self.batting_team, str):
            raise TypeError("batting_team must be a string.")

        if not self.batting_team.strip():
            raise ValueError("batting_team cannot be empty.")

        if not isinstance(self.bowling_team, str):
            raise TypeError("bowling_team must be a string.")

        if not self.bowling_team.strip():
            raise ValueError("bowling_team cannot be empty.")

        if self.batting_team.strip() == self.bowling_team.strip():
            raise ValueError(
                "batting_team and bowling_team must be different."
            )

        if self.target is not None:
            if not isinstance(self.target, int):
                raise TypeError("target must be an integer or None.")

            if self.target < 0:
                raise ValueError("target cannot be negative.")

        if not isinstance(self.max_overs, int):
            raise TypeError("max_overs must be an integer.")

        if self.max_overs <= 0:
            raise ValueError("max_overs must be greater than zero.")

        if not isinstance(self.wickets, int):
            raise TypeError("wickets must be an integer.")

        if self.wickets <= 0:
            raise ValueError("wickets must be greater than zero.")

        if not isinstance(self.dls_enabled, bool):
            raise TypeError("dls_enabled must be a boolean.")

        if not isinstance(self.super_over, bool):
            raise TypeError("super_over must be a boolean.")

        if not isinstance(self.follow_on, bool):
            raise TypeError("follow_on must be a boolean.")

        if not isinstance(self.declaration, bool):
            raise TypeError("declaration must be a boolean.")

        if not isinstance(self.match_type, str):
            raise TypeError("match_type must be a string.")

        self.match_type = self.match_type.strip().upper()

        if not self.match_type:
            raise ValueError("match_type cannot be empty.")

        logger.debug(
            "Validated ScorecardConfig("
            "innings=%s, batting=%s, bowling=%s, type=%s)",
            self.innings_number,
            self.batting_team,
            self.bowling_team,
            self.match_type,
        )

###############################################################################
# Scorecard Builder
###############################################################################


class ScorecardBuilder:
    """Coordinates scorecard population for an existing match workbook.

    The builder serves as the central orchestration object for all scorecard
    rendering operations. It owns references to the workbook, relevant
    worksheets, reusable helper APIs and internal caches that will be used by
    later rendering methods.

    This class **does not** create worksheets and **does not** populate any
    scorecard data during construction. Initialization is intentionally kept
    lightweight so that the builder may be safely instantiated before any
    innings are processed.

    Design Goals:
        * Centralize workbook references.
        * Centralize worksheet lookups.
        * Reuse helper APIs.
        * Avoid repeated worksheet searches.
        * Provide a shared logger.
        * Maintain reusable caches.
        * Keep rendering logic independent from workbook creation.

    Attributes:
        workbook:
            Existing workbook generated by ``create_match.py``.

        worksheets:
            Mapping of worksheet names to worksheet objects.

        config:
            Immutable scorecard configuration.

        logger:
            Module-specific logger.

        workbook_helper:
            Reference to workbook helper APIs.

        worksheet_cache:
            Cache of worksheet lookups.

        cell_cache:
            Cache of frequently accessed cells.

        style_cache:
            Cache reserved for reusable style objects.

        metadata:
            Arbitrary builder metadata reserved for future use.
    """

    def __init__(
        self,
        workbook: Workbook,
        worksheets: WorksheetMap,
        config: ScorecardConfig,
    ) -> None:
        """Initialize a scorecard builder.

        Args:
            workbook:
                Existing tournament workbook.

            worksheets:
                Mapping of worksheet names to worksheet objects created by
                ``create_match.py``.

            config:
                Immutable scorecard configuration describing the innings.

        Raises:
            TypeError:
                If supplied arguments are of unexpected types.

            ValueError:
                If the worksheet mapping is empty.
        """
        if not isinstance(workbook, Workbook):
            raise TypeError("workbook must be an openpyxl Workbook.")

        if not isinstance(worksheets, Mapping):
            raise TypeError("worksheets must implement Mapping.")

        if not worksheets:
            raise ValueError("worksheets mapping cannot be empty.")

        if not isinstance(config, ScorecardConfig):
            raise TypeError("config must be a ScorecardConfig instance.")

        self.workbook: Workbook = workbook
        self.worksheets: WorksheetMap = worksheets
        self.config: ScorecardConfig = config

        # Shared logger.
        self.logger: logging.Logger = logger.getChild(
            self.__class__.__name__
        )

        # ------------------------------------------------------------------
        # Internal caches
        # ------------------------------------------------------------------

        self.worksheet_cache: dict[str, Worksheet] = dict(worksheets)

        self.cell_cache: dict[str, Cell] = {}

        self.style_cache: dict[str, Any] = {}

        self.metadata: dict[str, Any] = {}

        # ------------------------------------------------------------------
        # Workbook helper initialization
        # ------------------------------------------------------------------

        # Placeholder for helper object(s) exposed by workbook.py.
        # Future parts of this module may replace this with a concrete helper
        # class or façade without affecting downstream rendering logic.
        self.workbook_helper: Any | None = None

        self.logger.debug(
            "Initialized %s for innings %d (%s vs %s).",
            self.__class__.__name__,
            self.config.innings_number,
            self.config.batting_team,
            self.config.bowling_team,
        )

    ###########################################################################
    # Worksheet Discovery
    ###########################################################################

    def locate_scorecard_sheet(self) -> Worksheet:
        """Locate the primary scorecard worksheet.

        This method delegates worksheet discovery to the helper APIs provided
        by ``workbook.py``. It does not perform any worksheet creation or
        scorecard population.

        Returns:
            The scorecard worksheet.

        Raises:
            RuntimeError:
                If the workbook helper has not been initialized.

            KeyError:
                If the worksheet cannot be located.
        """
        if self.workbook_helper is None:
            raise RuntimeError("Workbook helper has not been initialized.")

        worksheet = self.workbook_helper.locate_scorecard_sheet(
            workbook=self.workbook,
            innings=self.config.innings_number,
        )

        self.logger.debug(
            "Located scorecard worksheet '%s'.",
            worksheet.title,
        )

        return worksheet

    def locate_ball_log_sheet(self) -> Worksheet:
        """Locate the ball-by-ball worksheet.

        The worksheet is expected to have already been created by
        ``create_match.py``. Discovery is delegated entirely to the
        workbook helper layer.

        Returns:
            The ball log worksheet.

        Raises:
            RuntimeError:
                If workbook helpers are unavailable.
        """
        if self.workbook_helper is None:
            raise RuntimeError("Workbook helper has not been initialized.")

        worksheet = self.workbook_helper.locate_ball_log_sheet(
            workbook=self.workbook,
            innings=self.config.innings_number,
        )

        self.logger.debug(
            "Located ball log worksheet '%s'.",
            worksheet.title,
        )

        return worksheet

    def locate_summary_sheet(self) -> Worksheet:
        """Locate the match summary worksheet.

        Returns:
            The summary worksheet.

        Raises:
            RuntimeError:
                If workbook helpers are unavailable.
        """
        if self.workbook_helper is None:
            raise RuntimeError("Workbook helper has not been initialized.")

        worksheet = self.workbook_helper.locate_summary_sheet(
            workbook=self.workbook,
        )

        self.logger.debug(
            "Located summary worksheet '%s'.",
            worksheet.title,
        )

        return worksheet

    def validate_required_sheets(self) -> None:
        """Validate that all required worksheets exist.

        Worksheet validation is delegated to the helper APIs provided by
        ``workbook.py``.

        Raises:
            RuntimeError:
                If workbook helpers are unavailable.

            ValueError:
                If any required worksheet is missing.
        """
        if self.workbook_helper is None:
            raise RuntimeError("Workbook helper has not been initialized.")

        self.workbook_helper.validate_required_sheets(
            workbook=self.workbook,
            innings=self.config.innings_number,
        )

        self.logger.debug("Required worksheets successfully validated.")

    def register_sheet_references(self) -> None:
        """Populate the internal worksheet cache.

        This method discovers all worksheets required by the scorecard module
        and stores their references for future rendering operations.

        No worksheet content is modified.
        """
        self.validate_required_sheets()

        self.worksheet_cache["scorecard"] = self.locate_scorecard_sheet()
        self.worksheet_cache["ball_log"] = self.locate_ball_log_sheet()
        self.worksheet_cache["summary"] = self.locate_summary_sheet()

        self.logger.debug(
            "Registered %d worksheet reference(s).",
            len(self.worksheet_cache),
        )

    ###########################################################################
    # Batting Scorecard Population
    ###########################################################################

    def populate_batting_scorecard(
        self,
        worksheet: Worksheet,
        batters: Sequence[PlayerRecord],
    ) -> None:
        """Populate the batting scorecard.

        This method renders batting information into an existing worksheet
        created by ``create_match.py``. Cell addressing, formatting and formula
        insertion are delegated to helper APIs provided by ``workbook.py``.
        Statistical calculations are delegated to ``formulas.py``.

        No worksheet creation is performed.

        Args:
            worksheet:
                Destination batting scorecard worksheet.

            batters:
                Ordered batting records.

        Raises:
            TypeError:
                If invalid arguments are supplied.
        """
        if not isinstance(worksheet, Worksheet):
            raise TypeError("worksheet must be an openpyxl Worksheet.")

        if not isinstance(batters, Sequence):
            raise TypeError("batters must be a sequence.")

        self.logger.info(
            "Populating batting scorecard with %d batter(s).",
            len(batters),
        )

        for batting_position, batter in enumerate(batters, start=1):
            self._populate_batter_row(
                worksheet=worksheet,
                batting_position=batting_position,
                batter=batter,
            )

    def _populate_batter_row(
        self,
        worksheet: Worksheet,
        batting_position: int,
        batter: PlayerRecord,
    ) -> None:
        """Populate a single batting row.

        Args:
            worksheet:
                Destination worksheet.

            batting_position:
                Batting order.

            batter:
                Batter statistics.
        """
        helper = self.workbook_helper

        if helper is None:
            raise RuntimeError("Workbook helper has not been initialized.")

        row = helper.get_batting_row(
            worksheet=worksheet,
            batting_position=batting_position,
        )

        helper.write_value(
            worksheet=worksheet,
            row=row,
            field="BATTER_NAME",
            value=batter.get("name", ""),
        )

        helper.write_value(
            worksheet=worksheet,
            row=row,
            field="RUNS",
            value=batter.get("runs", 0),
        )

        helper.write_value(
            worksheet=worksheet,
            row=row,
            field="BALLS",
            value=batter.get("balls", 0),
        )

        helper.write_value(
            worksheet=worksheet,
            row=row,
            field="FOURS",
            value=batter.get("fours", 0),
        )

        helper.write_value(
            worksheet=worksheet,
            row=row,
            field="SIXES",
            value=batter.get("sixes", 0),
        )

        # Strike-rate calculation is delegated entirely to formulas.py.
        helper.write_formula(
            worksheet=worksheet,
            row=row,
            field="STRIKE_RATE",
            formula=calculate_run_rate(
                runs=batter.get("runs", 0),
                legal_balls=batter.get("balls", 0),
            ),
        )

        helper.write_value(
            worksheet=worksheet,
            row=row,
            field="DISMISSAL",
            value=batter.get("dismissal", ""),
        )

        helper.write_value(
            worksheet=worksheet,
            row=row,
            field="FIELDER",
            value=batter.get("fielder", ""),
        )

        helper.write_value(
            worksheet=worksheet,
            row=row,
            field="BOWLER",
            value=batter.get("bowler", ""),
        )

        self.logger.debug(
            "Populated batting row %d for '%s'.",
            batting_position,
            batter.get("name", ""),
        )

    ###########################################################################
    # Bowling Scorecard Population
    ###########################################################################

    def populate_bowling_scorecard(
        self,
        worksheet: Worksheet,
        bowlers: Sequence[BowlerRecord],
    ) -> None:
        """Populate the bowling scorecard.

        This method writes bowling statistics into an existing bowling
        scorecard worksheet created by ``create_match.py``.

        Workbook formatting, worksheet addressing and formula placement are
        delegated to helper APIs provided by ``workbook.py``. Mathematical
        calculations are delegated exclusively to ``formulas.py`` to ensure
        there is no duplicated calculation logic within this module.

        Args:
            worksheet:
                Destination bowling scorecard worksheet.

            bowlers:
                Ordered sequence of bowling records.

        Raises:
            TypeError:
                If invalid arguments are supplied.
        """
        if not isinstance(worksheet, Worksheet):
            raise TypeError("worksheet must be an openpyxl Worksheet.")

        if not isinstance(bowlers, Sequence):
            raise TypeError("bowlers must be a sequence.")

        self.logger.info(
            "Populating bowling scorecard with %d bowler(s).",
            len(bowlers),
        )

        for bowling_position, bowler in enumerate(bowlers, start=1):
            self._populate_bowler_row(
                worksheet=worksheet,
                bowling_position=bowling_position,
                bowler=bowler,
            )

    def _populate_bowler_row(
        self,
        worksheet: Worksheet,
        bowling_position: int,
        bowler: BowlerRecord,
    ) -> None:
        """Populate a single bowling scorecard row.

        Args:
            worksheet:
                Destination worksheet.

            bowling_position:
                Bowling order.

            bowler:
                Bowler statistics.
        """
        helper = self.workbook_helper

        if helper is None:
            raise RuntimeError("Workbook helper has not been initialized.")

        row = helper.get_bowling_row(
            worksheet=worksheet,
            bowling_position=bowling_position,
        )

        helper.write_value(
            worksheet=worksheet,
            row=row,
            field="BOWLER",
            value=bowler.get("name", ""),
        )

        helper.write_value(
            worksheet=worksheet,
            row=row,
            field="OVERS",
            value=balls_to_overs(
                bowler.get("legal_balls", 0),
            ),
        )

        helper.write_value(
            worksheet=worksheet,
            row=row,
            field="MAIDENS",
            value=bowler.get("maidens", 0),
        )

        helper.write_value(
            worksheet=worksheet,
            row=row,
            field="RUNS",
            value=bowler.get("runs", 0),
        )

        helper.write_value(
            worksheet=worksheet,
            row=row,
            field="WICKETS",
            value=bowler.get("wickets", 0),
        )

        # Economy-rate calculation is delegated entirely to formulas.py.
        helper.write_formula(
            worksheet=worksheet,
            row=row,
            field="ECONOMY",
            formula=calculate_run_rate(
                runs=bowler.get("runs", 0),
                legal_balls=bowler.get("legal_balls", 0),
            ),
        )

        helper.write_value(
            worksheet=worksheet,
            row=row,
            field="NO_BALLS",
            value=bowler.get("no_balls", 0),
        )

        helper.write_value(
            worksheet=worksheet,
            row=row,
            field="WIDES",
            value=bowler.get("wides", 0),
        )

        self.logger.debug(
            "Populated bowling row %d for '%s'.",
            bowling_position,
            bowler.get("name", ""),
        )


    ###########################################################################
    # Extras Section Population
    ###########################################################################

    def populate_extras(
        self,
        worksheet: Worksheet,
        extras: Mapping[str, int],
    ) -> None:
        """Populate the innings extras section.

        This method writes the individual extras breakdown into the existing
        scorecard worksheet. Cell discovery, addressing and writing are
        delegated entirely to helper APIs provided by ``workbook.py``.

        Calculation of total extras is delegated to ``formulas.py`` to avoid
        duplicating arithmetic within this module.

        Args:
            worksheet:
                Existing scorecard worksheet.

            extras:
                Mapping containing extras information.

                Supported keys:

                * byes
                * leg_byes
                * wides
                * no_balls
                * penalty

        Raises:
            TypeError:
                If supplied arguments are invalid.
        """
        if not isinstance(worksheet, Worksheet):
            raise TypeError("worksheet must be an openpyxl Worksheet.")

        if not isinstance(extras, Mapping):
            raise TypeError("extras must implement Mapping.")

        helper = self.workbook_helper

        if helper is None:
            raise RuntimeError("Workbook helper has not been initialized.")

        byes = int(extras.get("byes", 0))
        leg_byes = int(extras.get("leg_byes", 0))
        wides = int(extras.get("wides", 0))
        no_balls = int(extras.get("no_balls", 0))
        penalty = int(extras.get("penalty", 0))

        helper.write_named_value(
            worksheet=worksheet,
            field="BYES",
            value=byes,
        )

        helper.write_named_value(
            worksheet=worksheet,
            field="LEG_BYES",
            value=leg_byes,
        )

        helper.write_named_value(
            worksheet=worksheet,
            field="WIDES",
            value=wides,
        )

        helper.write_named_value(
            worksheet=worksheet,
            field="NO_BALLS",
            value=no_balls,
        )

        helper.write_named_value(
            worksheet=worksheet,
            field="PENALTY",
            value=penalty,
        )

        # Total extras calculation is delegated exclusively to formulas.py.
        helper.write_formula(
            worksheet=worksheet,
            field="TOTAL_EXTRAS",
            formula=calculate_total_extras(
                byes=byes,
                leg_byes=leg_byes,
                wides=wides,
                no_balls=no_balls,
                penalty=penalty,
            ),
        )

        self.logger.debug(
            "Extras populated (B=%d, LB=%d, W=%d, NB=%d, P=%d).",
            byes,
            leg_byes,
            wides,
            no_balls,
            penalty,
        )

    ###########################################################################
    # Innings Summary Population
    ###########################################################################

    def populate_innings_summary(
        self,
        worksheet: Worksheet,
        innings: InningsRecord,
    ) -> None:
        """Populate the innings summary section.

        This method writes the high-level innings summary into an existing
        worksheet. Cell addressing, formatting and writing are delegated to
        helper APIs provided by ``workbook.py``. All statistical calculations
        are delegated to ``formulas.py`` to ensure there is no duplicated
        calculation logic within this module.

        The summary includes:

        * Total
        * Wickets
        * Overs
        * Run Rate
        * Required Run Rate
        * Target
        * Remaining Balls
        * Remaining Runs

        Args:
            worksheet:
                Existing scorecard worksheet.

            innings:
                Mapping containing innings statistics.

        Raises:
            TypeError:
                If invalid arguments are supplied.
        """
        if not isinstance(worksheet, Worksheet):
            raise TypeError("worksheet must be an openpyxl Worksheet.")

        if not isinstance(innings, Mapping):
            raise TypeError("innings must implement Mapping.")

        helper = self.workbook_helper

        if helper is None:
            raise RuntimeError("Workbook helper has not been initialized.")

        total_runs = int(innings.get("total_runs", 0))
        wickets = int(innings.get("wickets", 0))
        legal_balls = int(innings.get("legal_balls", 0))

        target = (
            int(self.config.target)
            if self.config.target is not None
            else None
        )

        helper.write_named_value(
            worksheet=worksheet,
            field="TOTAL",
            value=total_runs,
        )

        helper.write_named_value(
            worksheet=worksheet,
            field="WICKETS",
            value=wickets,
        )

        helper.write_named_value(
            worksheet=worksheet,
            field="OVERS",
            value=balls_to_overs(legal_balls),
        )

        # Run rate calculation is delegated to formulas.py.
        helper.write_formula(
            worksheet=worksheet,
            field="RUN_RATE",
            formula=calculate_run_rate(
                runs=total_runs,
                legal_balls=legal_balls,
            ),
        )

        if target is not None:
            remaining_runs = max(target - total_runs, 0)

            remaining_balls_value = remaining_balls(
                max_overs=self.config.max_overs,
                legal_balls=legal_balls,
            )

            helper.write_formula(
                worksheet=worksheet,
                field="REQUIRED_RUN_RATE",
                formula=calculate_required_run_rate(
                    runs_required=remaining_runs,
                    legal_balls_remaining=remaining_balls_value,
                ),
            )

            helper.write_named_value(
                worksheet=worksheet,
                field="TARGET",
                value=target,
            )

            helper.write_formula(
                worksheet=worksheet,
                field="REMAINING_BALLS",
                formula=remaining_balls_value,
            )

            helper.write_named_value(
                worksheet=worksheet,
                field="REMAINING_RUNS",
                value=remaining_runs,
            )
        else:
            helper.write_named_value(
                worksheet=worksheet,
                field="TARGET",
                value="",
            )

            helper.write_named_value(
                worksheet=worksheet,
                field="REQUIRED_RUN_RATE",
                value="",
            )

            helper.write_named_value(
                worksheet=worksheet,
                field="REMAINING_BALLS",
                value="",
            )

            helper.write_named_value(
                worksheet=worksheet,
                field="REMAINING_RUNS",
                value="",
            )

        self.logger.debug(
            "Innings summary populated: %d/%d in %s overs.",
            total_runs,
            wickets,
            balls_to_overs(legal_balls),
        )


    ###########################################################################
    # Fall of Wickets
    ###########################################################################

    def populate_fall_of_wickets(
        self,
        worksheet: Worksheet,
        fall_of_wickets: Sequence[Mapping[str, Any]],
    ) -> None:
        """Populate the Fall of Wickets section.

        This method writes the wicket progression for an innings into an
        existing worksheet. Worksheet discovery, row addressing and cell
        writing are delegated entirely to helper APIs provided by
        ``workbook.py``.

        No calculations are performed by this method.

        Args:
            worksheet:
                Existing scorecard worksheet.

            fall_of_wickets:
                Ordered wicket records. Each record may contain:

                * wicket_number
                * batter
                * score
                * over
                * bowler

        Raises:
            TypeError:
                If supplied arguments are invalid.
        """
        if not isinstance(worksheet, Worksheet):
            raise TypeError("worksheet must be an openpyxl Worksheet.")

        if not isinstance(fall_of_wickets, Sequence):
            raise TypeError(
                "fall_of_wickets must be a sequence."
            )

        helper = self.workbook_helper

        if helper is None:
            raise RuntimeError(
                "Workbook helper has not been initialized."
            )

        self.logger.info(
            "Populating Fall of Wickets with %d record(s).",
            len(fall_of_wickets),
        )

        for index, wicket in enumerate(fall_of_wickets, start=1):
            self._populate_fall_of_wicket_row(
                worksheet=worksheet,
                row_number=index,
                wicket=wicket,
            )

    def _populate_fall_of_wicket_row(
        self,
        worksheet: Worksheet,
        row_number: int,
        wicket: Mapping[str, Any],
    ) -> None:
        """Populate a single Fall of Wickets row.

        Args:
            worksheet:
                Destination worksheet.

            row_number:
                Sequential wicket entry number.

            wicket:
                Wicket information mapping.
        """
        helper = self.workbook_helper

        if helper is None:
            raise RuntimeError(
                "Workbook helper has not been initialized."
            )

        row = helper.get_fall_of_wicket_row(
            worksheet=worksheet,
            wicket_number=row_number,
        )

        helper.write_value(
            worksheet=worksheet,
            row=row,
            field="WICKET_NUMBER",
            value=wicket.get("wicket_number", row_number),
        )

        helper.write_value(
            worksheet=worksheet,
            row=row,
            field="BATTER",
            value=wicket.get("batter", ""),
        )

        helper.write_value(
            worksheet=worksheet,
            row=row,
            field="SCORE",
            value=wicket.get("score", ""),
        )

        helper.write_value(
            worksheet=worksheet,
            row=row,
            field="OVER",
            value=wicket.get("over", ""),
        )

        helper.write_value(
            worksheet=worksheet,
            row=row,
            field="BOWLER",
            value=wicket.get("bowler", ""),
        )

        self.logger.debug(
            "Recorded Fall of Wicket #%s (%s).",
            wicket.get("wicket_number", row_number),
            wicket.get("batter", ""),
        )


    ###########################################################################
    # Partnership Table
    ###########################################################################

    def populate_partnership_table(
        self,
        worksheet: Worksheet,
        partnerships: Sequence[Mapping[str, Any]],
    ) -> None:
        """Populate the batting partnership table.

        This method renders partnership information into an existing
        scorecard worksheet. Worksheet layout, row discovery and cell
        addressing are delegated entirely to helper APIs provided by
        ``workbook.py``.

        No statistical calculations are performed by this method.

        Args:
            worksheet:
                Existing scorecard worksheet.

            partnerships:
                Ordered partnership records.

                Each record may contain:

                * partners
                * runs
                * balls
                * start_score
                * end_score

        Raises:
            TypeError:
                If invalid arguments are supplied.
        """
        if not isinstance(worksheet, Worksheet):
            raise TypeError("worksheet must be an openpyxl Worksheet.")

        if not isinstance(partnerships, Sequence):
            raise TypeError(
                "partnerships must be a sequence."
            )

        helper = self.workbook_helper

        if helper is None:
            raise RuntimeError(
                "Workbook helper has not been initialized."
            )

        self.logger.info(
            "Populating partnership table with %d partnership(s).",
            len(partnerships),
        )

        for partnership_number, partnership in enumerate(
            partnerships,
            start=1,
        ):
            self._populate_partnership_row(
                worksheet=worksheet,
                partnership_number=partnership_number,
                partnership=partnership,
            )

    def _populate_partnership_row(
        self,
        worksheet: Worksheet,
        partnership_number: int,
        partnership: Mapping[str, Any],
    ) -> None:
        """Populate a single partnership row.

        Args:
            worksheet:
                Destination worksheet.

            partnership_number:
                Sequential partnership number.

            partnership:
                Partnership information mapping.
        """
        helper = self.workbook_helper

        if helper is None:
            raise RuntimeError(
                "Workbook helper has not been initialized."
            )

        row = helper.get_partnership_row(
            worksheet=worksheet,
            partnership_number=partnership_number,
        )

        helper.write_value(
            worksheet=worksheet,
            row=row,
            field="PARTNERS",
            value=partnership.get("partners", ""),
        )

        helper.write_value(
            worksheet=worksheet,
            row=row,
            field="RUNS",
            value=partnership.get("runs", 0),
        )

        helper.write_value(
            worksheet=worksheet,
            row=row,
            field="BALLS",
            value=partnership.get("balls", 0),
        )

        helper.write_value(
            worksheet=worksheet,
            row=row,
            field="START_SCORE",
            value=partnership.get("start_score", ""),
        )

        helper.write_value(
            worksheet=worksheet,
            row=row,
            field="END_SCORE",
            value=partnership.get("end_score", ""),
        )

        self.logger.debug(
            "Recorded partnership %d (%s).",
            partnership_number,
            partnership.get("partners", ""),
        )


    ###########################################################################
    # Validation Integration
    ###########################################################################

    def apply_scorecard_validations(self) -> None:
        """Apply all scorecard data validations.

        This method applies reusable validation rules to the scorecard
        worksheets. Validation construction and application are delegated
        entirely to helper APIs exposed by ``validation.py`` via the
        ``workbook.py`` helper layer.

        The following reusable dropdowns are applied:

        * Batter
        * Bowler
        * Dismissal
        * Extras
        * Fielder

        No validation rules are implemented directly in this module.

        Raises:
            RuntimeError:
                If workbook helpers have not been initialized.
        """
        helper = self.workbook_helper

        if helper is None:
            raise RuntimeError(
                "Workbook helper has not been initialized."
            )

        self.logger.info("Applying scorecard data validations.")

        self._apply_batter_validation()
        self._apply_bowler_validation()
        self._apply_dismissal_validation()
        self._apply_extras_validation()
        self._apply_fielder_validation()

        self.logger.debug("Scorecard validations successfully applied.")

    def _apply_batter_validation(self) -> None:
        """Apply reusable batter dropdown validation."""
        helper = self.workbook_helper

        helper.apply_batter_validation(
            worksheet=self.worksheet_cache["scorecard"],
        )

        self.logger.debug("Applied batter validation.")

    def _apply_bowler_validation(self) -> None:
        """Apply reusable bowler dropdown validation."""
        helper = self.workbook_helper

        helper.apply_bowler_validation(
            worksheet=self.worksheet_cache["scorecard"],
        )

        self.logger.debug("Applied bowler validation.")

    def _apply_dismissal_validation(self) -> None:
        """Apply reusable dismissal dropdown validation."""
        helper = self.workbook_helper

        helper.apply_dismissal_validation(
            worksheet=self.worksheet_cache["scorecard"],
        )

        self.logger.debug("Applied dismissal validation.")

    def _apply_extras_validation(self) -> None:
        """Apply reusable extras dropdown validation."""
        helper = self.workbook_helper

        helper.apply_extras_validation(
            worksheet=self.worksheet_cache["scorecard"],
        )

        self.logger.debug("Applied extras validation.")

    def _apply_fielder_validation(self) -> None:
        """Apply reusable fielder dropdown validation."""
        helper = self.workbook_helper

        helper.apply_fielder_validation(
            worksheet=self.worksheet_cache["scorecard"],
        )

        self.logger.debug("Applied fielder validation.")



    ###########################################################################
    # Formula Integration
    ###########################################################################

    def populate_formula_fields(
        self,
        worksheet: Worksheet,
        innings: InningsRecord,
    ) -> None:
        """Populate all calculated scorecard fields.

        This method centralizes every formula-driven value required by the
        scorecard. All mathematical calculations are delegated exclusively to
        ``formulas.py`` while worksheet interaction is delegated to
        ``workbook.py``.

        The following calculated fields are populated:

        * Batting strike rates
        * Bowling economy rates
        * Innings totals
        * Total extras
        * Current run rate
        * Remaining balls
        * Required run rate

        No calculation logic is implemented directly within this module.

        Args:
            worksheet:
                Destination scorecard worksheet.

            innings:
                Mapping containing innings statistics.

        Raises:
            TypeError:
                If invalid arguments are supplied.

            RuntimeError:
                If workbook helpers have not been initialized.
        """
        if not isinstance(worksheet, Worksheet):
            raise TypeError("worksheet must be an openpyxl Worksheet.")

        if not isinstance(innings, Mapping):
            raise TypeError("innings must implement Mapping.")

        helper = self.workbook_helper

        if helper is None:
            raise RuntimeError(
                "Workbook helper has not been initialized."
            )

        self.logger.info("Populating calculated scorecard fields.")

        self._populate_batting_formula_fields(
            worksheet=worksheet,
            batters=innings.get("batters", ()),
        )

        self._populate_bowling_formula_fields(
            worksheet=worksheet,
            bowlers=innings.get("bowlers", ()),
        )

        self._populate_summary_formula_fields(
            worksheet=worksheet,
            innings=innings,
        )

        self.logger.debug(
            "Completed formula integration for innings %d.",
            self.config.innings_number,
        )

    def _populate_batting_formula_fields(
        self,
        worksheet: Worksheet,
        batters: Sequence[PlayerRecord],
    ) -> None:
        """Populate batting formula fields."""
        helper = self.workbook_helper

        for batting_position, batter in enumerate(batters, start=1):
            row = helper.get_batting_row(
                worksheet=worksheet,
                batting_position=batting_position,
            )

            helper.write_formula(
                worksheet=worksheet,
                row=row,
                field="STRIKE_RATE",
                formula=calculate_run_rate(
                    runs=int(batter.get("runs", 0)),
                    legal_balls=int(batter.get("balls", 0)),
                ),
            )

    def _populate_bowling_formula_fields(
        self,
        worksheet: Worksheet,
        bowlers: Sequence[BowlerRecord],
    ) -> None:
        """Populate bowling formula fields."""
        helper = self.workbook_helper

        for bowling_position, bowler in enumerate(bowlers, start=1):
            row = helper.get_bowling_row(
                worksheet=worksheet,
                bowling_position=bowling_position,
            )

            helper.write_formula(
                worksheet=worksheet,
                row=row,
                field="ECONOMY",
                formula=calculate_run_rate(
                    runs=int(bowler.get("runs", 0)),
                    legal_balls=int(bowler.get("legal_balls", 0)),
                ),
            )

    def _populate_summary_formula_fields(
        self,
        worksheet: Worksheet,
        innings: InningsRecord,
    ) -> None:
        """Populate innings summary formula fields."""
        helper = self.workbook_helper

        total_runs = int(innings.get("total_runs", 0))
        legal_balls = int(innings.get("legal_balls", 0))

        byes = int(innings.get("byes", 0))
        leg_byes = int(innings.get("leg_byes", 0))
        wides = int(innings.get("wides", 0))
        no_balls = int(innings.get("no_balls", 0))
        penalty = int(innings.get("penalty", 0))

        helper.write_formula(
            worksheet=worksheet,
            field="TOTAL_EXTRAS",
            formula=calculate_total_extras(
                byes=byes,
                leg_byes=leg_byes,
                wides=wides,
                no_balls=no_balls,
                penalty=penalty,
            ),
        )

        helper.write_formula(
            worksheet=worksheet,
            field="RUN_RATE",
            formula=calculate_run_rate(
                runs=total_runs,
                legal_balls=legal_balls,
            ),
        )

        if self.config.target is not None:
            runs_required = max(
                self.config.target - total_runs,
                0,
            )

            balls_remaining = remaining_balls(
                max_overs=self.config.max_overs,
                legal_balls=legal_balls,
            )

            helper.write_formula(
                worksheet=worksheet,
                field="REMAINING_BALLS",
                formula=balls_remaining,
            )

            helper.write_formula(
                worksheet=worksheet,
                field="REQUIRED_RUN_RATE",
                formula=calculate_required_run_rate(
                    runs_required=runs_required,
                    legal_balls_remaining=balls_remaining,
                ),
            )

        helper.write_formula(
            worksheet=worksheet,
            field="TOTAL",
            formula=calculate_innings_total(
                batters=innings.get("batters", ()),
                extras=calculate_total_extras(
                    byes=byes,
                    leg_byes=leg_byes,
                    wides=wides,
                    no_balls=no_balls,
                    penalty=penalty,
                ),
            ),
        )


    ###########################################################################
    # Statistics Generation
    ###########################################################################

    def populate_statistics(
        self,
        worksheet: Worksheet,
        innings: InningsRecord,
    ) -> None:
        """Populate innings statistical summaries.

        This method generates and writes commonly used innings statistics.
        All statistical calculations are delegated exclusively to
        ``formulas.py`` while worksheet interaction is delegated to the
        helper APIs provided by ``workbook.py``.

        Statistics generated include:

        * Highest scorer
        * Best bowling figures
        * Total boundaries
        * Total dot balls
        * Maiden overs
        * Highest partnership

        Args:
            worksheet:
                Destination scorecard worksheet.

            innings:
                Structured innings record.

        Raises:
            TypeError:
                If supplied arguments are invalid.

            RuntimeError:
                If workbook helpers have not been initialized.
        """
        if not isinstance(worksheet, Worksheet):
            raise TypeError("worksheet must be an openpyxl Worksheet.")

        if not isinstance(innings, Mapping):
            raise TypeError("innings must implement Mapping.")

        helper = self.workbook_helper

        if helper is None:
            raise RuntimeError(
                "Workbook helper has not been initialized."
            )

        self.logger.info("Generating innings statistics.")

        batters = innings.get("batters", ())
        bowlers = innings.get("bowlers", ())
        partnerships = innings.get("partnerships", ())

        helper.write_named_value(
            worksheet=worksheet,
            field="HIGHEST_SCORER",
            value=highest_scorer(batters),
        )

        helper.write_named_value(
            worksheet=worksheet,
            field="BEST_BOWLING",
            value=best_bowling_figures(bowlers),
        )

        helper.write_formula(
            worksheet=worksheet,
            field="BOUNDARIES",
            formula=calculate_total_boundaries(batters),
        )

        helper.write_formula(
            worksheet=worksheet,
            field="DOT_BALLS",
            formula=calculate_dot_balls(bowlers),
        )

        helper.write_formula(
            worksheet=worksheet,
            field="MAIDEN_OVERS",
            formula=calculate_maiden_overs(bowlers),
        )

        helper.write_named_value(
            worksheet=worksheet,
            field="PARTNERSHIP_RECORD",
            value=highest_partnership(partnerships),
        )

        self.logger.debug(
            "Statistics populated for innings %d.",
            self.config.innings_number,
        )


    ###########################################################################
    # Navigation Links
    ###########################################################################

    def register_navigation_links(self) -> None:
        """Register all scorecard navigation hyperlinks.

        This method creates the internal workbook hyperlinks used to navigate
        between related worksheets. Hyperlink creation is delegated entirely to
        reusable helper APIs provided by ``workbook.py``.

        Navigation links include:

        * Scorecard
        * Ball Log
        * Summary
        * Statistics

        No worksheet creation or score population occurs.

        Raises:
            RuntimeError:
                If workbook helpers have not been initialized.
        """
        helper = self.workbook_helper

        if helper is None:
            raise RuntimeError(
                "Workbook helper has not been initialized."
            )

        self.logger.info("Registering scorecard navigation links.")

        self._register_scorecard_link()
        self._register_ball_log_link()
        self._register_summary_link()
        self._register_statistics_link()

        self.logger.debug("Navigation links successfully registered.")

    def _register_scorecard_link(self) -> None:
        """Create hyperlink to the scorecard worksheet."""
        helper = self.workbook_helper

        helper.create_navigation_link(
            source=self.worksheet_cache["scorecard"],
            destination=self.worksheet_cache["scorecard"],
            link_name="SCORECARD",
        )

        self.logger.debug("Registered Scorecard hyperlink.")

    def _register_ball_log_link(self) -> None:
        """Create hyperlink to the ball log worksheet."""
        helper = self.workbook_helper

        helper.create_navigation_link(
            source=self.worksheet_cache["scorecard"],
            destination=self.worksheet_cache["ball_log"],
            link_name="BALL_LOG",
        )

        self.logger.debug("Registered Ball Log hyperlink.")

    def _register_summary_link(self) -> None:
        """Create hyperlink to the summary worksheet."""
        helper = self.workbook_helper

        helper.create_navigation_link(
            source=self.worksheet_cache["scorecard"],
            destination=self.worksheet_cache["summary"],
            link_name="SUMMARY",
        )

        self.logger.debug("Registered Summary hyperlink.")

    def _register_statistics_link(self) -> None:
        """Create hyperlink to the statistics worksheet."""
        helper = self.workbook_helper

        statistics_sheet = self.worksheet_cache.get(
            "statistics",
            self.worksheet_cache["summary"],
        )

        helper.create_navigation_link(
            source=self.worksheet_cache["scorecard"],
            destination=statistics_sheet,
            link_name="STATISTICS",
        )

        self.logger.debug("Registered Statistics hyperlink.")


    ###########################################################################
    # Worksheet Protection & Print Configuration
    ###########################################################################

    def configure_worksheet(self) -> None:
        """Configure worksheet protection and print settings.

        This method applies all worksheet-level configuration required after
        scorecard population has completed.

        Configuration includes:

        * Worksheet protection
        * Unlocking scoring/input cells
        * Freeze panes
        * Print settings

        All operations are delegated exclusively to helper APIs provided by
        ``workbook.py``. No worksheet properties are manipulated directly by
        this module.

        Raises:
            RuntimeError:
                If workbook helpers have not been initialized.
        """
        helper = self.workbook_helper

        if helper is None:
            raise RuntimeError(
                "Workbook helper has not been initialized."
            )

        self.logger.info(
            "Applying worksheet protection and print configuration."
        )

        for worksheet in self.worksheet_cache.values():
            self._configure_single_worksheet(worksheet)

        self.logger.debug(
            "Worksheet configuration completed for %d worksheet(s).",
            len(self.worksheet_cache),
        )

    def _configure_single_worksheet(
        self,
        worksheet: Worksheet,
    ) -> None:
        """Configure a single worksheet.

        Args:
            worksheet:
                Existing worksheet created by ``create_match.py``.
        """
        helper = self.workbook_helper

        helper.unlock_scoring_cells(
            worksheet=worksheet,
        )

        helper.configure_worksheet_protection(
            worksheet=worksheet,
        )

        helper.configure_freeze_panes(
            worksheet=worksheet,
        )

        helper.configure_print_settings(
            worksheet=worksheet,
        )

        self.logger.debug(
            "Configured worksheet '%s'.",
            worksheet.title,
        )


    ###########################################################################
    # Scorecard Orchestration
    ###########################################################################

    def build_scorecard(
        self,
        innings: InningsRecord,
    ) -> None:
        """Build a complete innings scorecard.

        This is the primary orchestration entry point for scorecard generation.
        It coordinates the existing builder functionality without introducing
        duplicate rendering or calculation logic.

        The workflow is intentionally linear:

        1. Discover and validate worksheets.
        2. Populate batting scorecard.
        3. Populate bowling scorecard.
        4. Populate innings summary.
        5. Populate supporting sections.
        6. Apply validations.
        7. Populate calculated fields.
        8. Generate statistics.
        9. Register navigation links.
        10. Configure worksheet protection.
        11. Finalize the scorecard.

        Args:
            innings:
                Structured innings data.

        Raises:
            RuntimeError:
                If required worksheets cannot be located.
        """
        self.logger.info(
            "Building scorecard for innings %d.",
            self.config.innings_number,
        )

        self.register_sheet_references()

        self.populate_batting(innings)
        self.populate_bowling(innings)
        self.populate_summary(innings)

        scorecard_sheet = self.worksheet_cache["scorecard"]

        self.populate_extras(
            scorecard_sheet,
            innings.get("extras", {}),
        )

        self.populate_fall_of_wickets(
            scorecard_sheet,
            innings.get("fall_of_wickets", ()),
        )

        self.populate_partnership_table(
            scorecard_sheet,
            innings.get("partnerships", ()),
        )

        self.apply_scorecard_validations()

        self.populate_formula_fields(
            scorecard_sheet,
            innings,
        )

        self.populate_statistics(
            scorecard_sheet,
            innings,
        )

        self.register_navigation_links()

        self.configure_worksheet()

        self.finalize_scorecard()

    def populate_batting(
        self,
        innings: InningsRecord,
    ) -> None:
        """Populate the batting scorecard.

        Args:
            innings:
                Structured innings record.
        """
        self.populate_batting_scorecard(
            worksheet=self.worksheet_cache["scorecard"],
            batters=innings.get("batters", ()),
        )

    def populate_bowling(
        self,
        innings: InningsRecord,
    ) -> None:
        """Populate the bowling scorecard.

        Args:
            innings:
                Structured innings record.
        """
        self.populate_bowling_scorecard(
            worksheet=self.worksheet_cache["scorecard"],
            bowlers=innings.get("bowlers", ()),
        )

    def populate_summary(
        self,
        innings: InningsRecord,
    ) -> None:
        """Populate the innings summary.

        Args:
            innings:
                Structured innings record.
        """
        self.populate_innings_summary(
            worksheet=self.worksheet_cache["scorecard"],
            innings=innings,
        )

    def finalize_scorecard(self) -> None:
        """Finalize scorecard generation.

        Finalization activities are delegated to the workbook helper layer and
        may include workbook recalculation flags, hyperlink refresh,
        print-area updates or other post-processing operations.

        No rendering logic is duplicated here.
        """
        helper = self.workbook_helper

        if helper is None:
            raise RuntimeError(
                "Workbook helper has not been initialized."
            )

        helper.finalize_scorecard(
            workbook=self.workbook,
            worksheets=self.worksheet_cache,
        )

        self.logger.info(
            "Completed scorecard generation for innings %d.",
            self.config.innings_number,
        )


###############################################################################
# Convenience Functions
###############################################################################

def build_scorecard(
    workbook: Workbook,
    worksheets: WorksheetMap,
    config: ScorecardConfig,
    innings: InningsRecord,
) -> ScorecardBuilder:
    """Build a scorecard for any supported innings.

    This is the primary convenience entry point for callers that do not need
    to interact with :class:`ScorecardBuilder` directly.

    Args:
        workbook:
            Existing workbook created by ``create_match.py``.

        worksheets:
            Mapping of worksheet names to worksheet objects.

        config:
            Scorecard configuration.

        innings:
            Structured innings data.

    Returns:
        The initialized :class:`ScorecardBuilder` after scorecard generation.
    """
    builder = ScorecardBuilder(
        workbook=workbook,
        worksheets=worksheets,
        config=config,
    )

    builder.build_scorecard(innings)

    return builder


def build_first_innings_scorecard(
    workbook: Workbook,
    worksheets: WorksheetMap,
    config: ScorecardConfig,
    innings: InningsRecord,
) -> ScorecardBuilder:
    """Build the first-innings scorecard."""
    if config.innings_number != 1:
        raise ValueError(
            "build_first_innings_scorecard() requires innings_number == 1."
        )

    return build_scorecard(
        workbook=workbook,
        worksheets=worksheets,
        config=config,
        innings=innings,
    )


def build_second_innings_scorecard(
    workbook: Workbook,
    worksheets: WorksheetMap,
    config: ScorecardConfig,
    innings: InningsRecord,
) -> ScorecardBuilder:
    """Build the second-innings scorecard."""
    if config.innings_number != 2:
        raise ValueError(
            "build_second_innings_scorecard() requires innings_number == 2."
        )

    return build_scorecard(
        workbook=workbook,
        worksheets=worksheets,
        config=config,
        innings=innings,
    )


def build_test_innings_scorecard(
    workbook: Workbook,
    worksheets: WorksheetMap,
    config: ScorecardConfig,
    innings: InningsRecord,
) -> ScorecardBuilder:
    """Build a Test-match innings scorecard.

    This is a thin wrapper around :func:`build_scorecard`.
    """
    if config.match_type != "TEST":
        raise ValueError(
            "build_test_innings_scorecard() requires match_type='TEST'."
        )

    return build_scorecard(
        workbook=workbook,
        worksheets=worksheets,
        config=config,
        innings=innings,
    )


def update_scorecard(
    builder: ScorecardBuilder,
    innings: InningsRecord,
) -> ScorecardBuilder:
    """Update an existing scorecard.

    This convenience wrapper reuses an existing
    :class:`ScorecardBuilder` instance and invokes its orchestration logic.

    Args:
        builder:
            Existing scorecard builder.

        innings:
            Updated innings data.

    Returns:
        The supplied builder instance.
    """
    if not isinstance(builder, ScorecardBuilder):
        raise TypeError(
            "builder must be an instance of ScorecardBuilder."
        )

    builder.build_scorecard(innings)

    return builder

###############################################################################
# Module Finalization
###############################################################################

"""
Finalization Notes
------------------
This completes the public API for ``scorecard.py``.

Compatibility
~~~~~~~~~~~~~
The module is designed to integrate with the following project modules:

* constants.py
* formulas.py
* validation.py
* workbook.py
* create_match.py

Design Guarantees
~~~~~~~~~~~~~~~~~
* Scorecard rendering only.
* No worksheet creation.
* No duplicated statistical calculations.
* Formula generation delegated to ``formulas.py``.
* Workbook interaction delegated to ``workbook.py``.
* Validation delegated to ``validation.py``.
* Thin convenience wrappers for external callers.
* Stable public API for future desktop, web and CLI front-ends.
"""

###############################################################################
# Public API
###############################################################################

__all__.extend(
    [
        # Configuration
        "ScorecardConfig",
        # Primary builder
        "ScorecardBuilder",
        # Convenience functions
        "build_scorecard",
        "build_first_innings_scorecard",
        "build_second_innings_scorecard",
        "build_test_innings_scorecard",
        "update_scorecard",
    ]
)

# Keep exports deterministic and remove accidental duplicates introduced
# during incremental generation.
__all__ = sorted(set(__all__))
