"""
workbook.py
===========

Central workbook orchestration module for the Cricket Tournament Excel system.

Overview
--------
This module acts as the highest-level coordinator responsible for constructing,
configuring, and exporting complete Excel workbooks that represent an entire
cricket tournament.

Unlike lower-level modules that focus on a single concern (such as formulas,
validation, constants, styling, or worksheet generation), this module provides
the orchestration layer that combines those components into a fully functional
spreadsheet application.

The responsibilities of this module include (implemented incrementally across
future parts of this file):

* Workbook creation
* Workbook configuration
* Metadata initialization
* Worksheet orchestration
* Named range registration
* Formula integration
* Data validation registration
* Conditional formatting integration
* Freeze panes
* Print settings
* Workbook properties
* Protection configuration
* Export helpers
* Version compatibility checks
* Logging
* High-level error handling

Design Philosophy
-----------------
The architecture intentionally separates orchestration from implementation.

Each worksheet should be created by a dedicated builder or helper rather than
embedding worksheet logic directly inside this module. Likewise, formatting,
validation, formulas, and constants are delegated to their respective modules.

This design provides:

* excellent maintainability
* easier testing
* modular development
* predictable extension points
* reduced coupling
* simpler future refactoring

Thread Safety
-------------
Workbook objects provided by openpyxl are mutable and should not be considered
thread-safe. A workbook should be constructed and modified from a single thread.

Python Version
--------------
Python 3.12+

Third-party Dependencies
------------------------
* openpyxl

Internal Dependencies
---------------------
* constants
* formulas
* validation

Nothing in this initial module foundation performs workbook creation. This file
currently establishes the shared imports, typing infrastructure, logging, and
module-level constants used throughout later sections.
"""

from __future__ import annotations

###############################################################################
# Standard library imports
###############################################################################

import logging
from pathlib import Path
from typing import Final
from typing import TypeAlias

###############################################################################
# Third-party imports
###############################################################################

from openpyxl import Workbook
from openpyxl.cell.cell import Cell
from openpyxl.styles import Alignment
from openpyxl.styles import Border
from openpyxl.styles import Font
from openpyxl.styles import PatternFill
from openpyxl.styles import Protection
from openpyxl.styles import Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

###############################################################################
# Internal project imports
###############################################################################

import constants
import formulas
import validation

###############################################################################
# Type aliases
###############################################################################

WorkbookType: TypeAlias = Workbook
"""Alias representing an openpyxl workbook."""

WorksheetType: TypeAlias = Worksheet
"""Alias representing an openpyxl worksheet."""

CellType: TypeAlias = Cell
"""Alias representing an openpyxl cell."""

PathLike: TypeAlias = str | Path
"""
Represents an output path accepted by workbook export helpers.

The alias allows callers to supply either a string path or a pathlib.Path
instance.
"""

###############################################################################
# Logging configuration
###############################################################################

LOGGER_NAME: Final[str] = "cricket_excel.workbook"
"""
Canonical logger name used throughout this module.

Applications embedding this project may configure this logger independently
from the root logger to obtain detailed workbook construction diagnostics.
"""

logger = logging.getLogger(LOGGER_NAME)

###############################################################################
# Module constants
###############################################################################

MODULE_NAME: Final[str] = "workbook"
"""Human-readable module identifier."""

MODULE_VERSION: Final[str] = "1.0.0"
"""
Internal module implementation version.

This value identifies the implementation of this orchestration module rather
than the version of the overall application.
"""

DEFAULT_SHEET_INDEX: Final[int] = 0
"""
Default insertion index for worksheets.

Future worksheet builders may override this value when constructing workbooks
with custom ordering.
"""

EXCEL_MAX_ROWS: Final[int] = 1_048_576
"""Maximum number of rows supported by modern Excel workbooks."""

EXCEL_MAX_COLUMNS: Final[int] = 16_384
"""Maximum number of columns supported by modern Excel workbooks."""

EXCEL_FIRST_ROW: Final[int] = 1
"""Index of the first worksheet row."""

EXCEL_FIRST_COLUMN: Final[int] = 1
"""Index of the first worksheet column."""

###############################################################################
# Internal helper constants
###############################################################################

_INTERNAL_PREFIX: Final[str] = "_"
"""
Naming prefix reserved for workbook-internal identifiers.

Future named ranges, hidden worksheets, and implementation-specific resources
may use this prefix to distinguish them from user-visible objects.
"""

_LOG_MESSAGE_PREFIX: Final[str] = "[Workbook]"
"""
Standard prefix applied to diagnostic log messages generated by this module.
"""

_COLUMN_CACHE_LIMIT: Final[int] = EXCEL_MAX_COLUMNS
"""
Upper bound used by future column-letter helper caches.

The value matches Excel's maximum supported column count.
"""

###############################################################################
# Public exports
###############################################################################

__all__: list[str] = []

###############################################################################
# Workbook configuration
###############################################################################

from dataclasses import dataclass, field


@dataclass(slots=True, kw_only=True)
class WorkbookConfig:
    """Configuration used by :class:`WorkbookBuilder`.

    This dataclass centralizes every configurable option required to construct
    a workbook. Future parts of ``workbook.py`` will consume this object during
    workbook creation, worksheet initialization, styling, printing, protection,
    and export.

    Keeping all configuration in a single immutable-style object makes the
    workbook builder easier to test, extend, and validate.

    Attributes:
        workbook_title:
            Human-readable workbook title.

        author:
            Workbook author written to document properties.

        company:
            Company or organization associated with the workbook.

        tournament_name:
            Name of the cricket tournament represented by the workbook.

        default_font:
            Default font family applied throughout the workbook.

        default_font_size:
            Default font size in points.

        default_row_height:
            Default worksheet row height.

        default_column_width:
            Default worksheet column width.

        theme_name:
            Logical theme identifier. Future style modules may map this value
            to fonts, colours, borders, and fills.

        date_format:
            Excel number format for dates.

        time_format:
            Excel number format for times.

        datetime_format:
            Excel number format for combined date/time values.

        workbook_protection_enabled:
            Enables workbook-level protection by default.

        structure_protection:
            Protect workbook structure.

        windows_protection:
            Protect workbook window layout.

        default_sheet_protection:
            Enable worksheet protection by default.

        allow_select_locked_cells:
            Allow selecting locked cells.

        allow_select_unlocked_cells:
            Allow selecting unlocked cells.

        fit_to_page:
            Enable worksheet fit-to-page printing.

        fit_to_width:
            Printed page width.

        fit_to_height:
            Printed page height.

        landscape:
            Print using landscape orientation.

        center_horizontally:
            Horizontally center printed pages.

        center_vertically:
            Vertically center printed pages.
    """

    workbook_title: str = "Cricket Tournament Workbook"
    author: str = "Cricket Tournament Manager"
    company: str = ""
    tournament_name: str = ""

    default_font: str = "Calibri"
    default_font_size: float = 11.0

    default_row_height: float = 15.0
    default_column_width: float = 12.0

    theme_name: str = "default"

    date_format: str = "yyyy-mm-dd"
    time_format: str = "hh:mm:ss"
    datetime_format: str = "yyyy-mm-dd hh:mm:ss"

    workbook_protection_enabled: bool = False
    structure_protection: bool = False
    windows_protection: bool = False

    default_sheet_protection: bool = False
    allow_select_locked_cells: bool = True
    allow_select_unlocked_cells: bool = True

    fit_to_page: bool = True
    fit_to_width: int = 1
    fit_to_height: int = 0
    landscape: bool = True
    center_horizontally: bool = True
    center_vertically: bool = False

    def __post_init__(self) -> None:
        """Validate configuration values.

        Raises:
            TypeError:
                If a value has an unexpected type.

            ValueError:
                If a value falls outside the supported range.
        """
        text_fields = (
            ("workbook_title", self.workbook_title),
            ("author", self.author),
            ("company", self.company),
            ("tournament_name", self.tournament_name),
            ("default_font", self.default_font),
            ("theme_name", self.theme_name),
            ("date_format", self.date_format),
            ("time_format", self.time_format),
            ("datetime_format", self.datetime_format),
        )

        for name, value in text_fields:
            if not isinstance(value, str):
                raise TypeError(f"{name} must be a string.")

            if not value.strip() and name not in {"company", "tournament_name"}:
                raise ValueError(f"{name} cannot be empty.")

        if self.default_font_size <= 0:
            raise ValueError("default_font_size must be greater than zero.")

        if self.default_row_height <= 0:
            raise ValueError("default_row_height must be greater than zero.")

        if self.default_column_width <= 0:
            raise ValueError("default_column_width must be greater than zero.")

        if self.fit_to_width < 0:
            raise ValueError("fit_to_width cannot be negative.")

        if self.fit_to_height < 0:
            raise ValueError("fit_to_height cannot be negative.")


###############################################################################
# Workbook builder skeleton
###############################################################################


class WorkbookBuilder:
    """High-level workbook orchestration class.

    The builder coordinates construction of the complete Excel workbook while
    delegating specialized responsibilities to lower-level modules.

    Only the foundational object model is established here. Workbook creation,
    worksheet creation, styling, validation, formulas, named ranges, printing,
    protection, and export are implemented in later parts of this module.
    """

    def __init__(
        self,
        config: WorkbookConfig | None = None,
    ) -> None:
        """Initialize the workbook builder.

        Args:
            config:
                Optional workbook configuration. When omitted, a default
                configuration is created.
        """
        self._config: WorkbookConfig = config or WorkbookConfig()

        self._workbook: WorkbookType | None = None
        self._worksheet_registry: dict[str, WorksheetType] = {}

        self._logger: logging.Logger = logger

        self._initialize_helpers()

    @property
    def config(self) -> WorkbookConfig:
        """Return the active workbook configuration."""
        return self._config

    @property
    def workbook(self) -> WorkbookType | None:
        """Return the managed workbook instance.

        Returns:
            The workbook once created, otherwise ``None``.
        """
        return self._workbook

    @property
    def worksheet_registry(self) -> dict[str, WorksheetType]:
        """Return the internal worksheet registry.

        Future workbook construction stages will populate this mapping with
        logical worksheet names and their corresponding worksheet objects.
        """
        return self._worksheet_registry

    @property
    def logger(self) -> logging.Logger:
        """Return the module logger."""
        return self._logger

    def _initialize_helpers(self) -> None:
        """Initialize internal helper infrastructure.

        The initial implementation prepares extension points required by later
        workbook construction stages. Future parts may populate caches,
        registries, style managers, validation managers, formula helpers, and
        workbook services.
        """
        self._helper_registry: dict[str, object] = {}



###############################################################################
# WorkbookBuilder - workbook lifecycle methods
###############################################################################

    def create_workbook(self) -> WorkbookType:
        """Create and initialize a new workbook.

        This method is the primary entry point for workbook construction.
        Individual initialization steps are intentionally delegated to smaller
        methods so subclasses may override specific phases without replacing the
        complete workflow.

        The current implementation performs only foundational initialization.
        Worksheet population, styling, formulas, validations, and other
        application-specific features are implemented in later parts of this
        module.

        Returns:
            The newly created workbook.

        Raises:
            RuntimeError:
                If workbook creation fails.
        """
        self._logger.info("%s Creating workbook.", _LOG_MESSAGE_PREFIX)

        try:
            self._workbook = Workbook()

            self.initialize_internal_state()
            self.initialize_workbook()
            self.remove_default_sheet()
            self.set_workbook_properties()
            self.create_core_sheets()
            self.create_hidden_sheets()

            self._logger.info(
                "%s Workbook initialization completed successfully.",
                _LOG_MESSAGE_PREFIX,
            )

            return self._workbook

        except Exception as exc:
            self._logger.exception(
                "%s Workbook creation failed.",
                _LOG_MESSAGE_PREFIX,
            )
            self._workbook = None
            raise RuntimeError("Unable to create workbook.") from exc

    def initialize_workbook(self) -> None:
        """Initialize workbook-wide settings.

        This method performs only high-level initialization that does not depend
        upon worksheet creation.

        Future implementations may configure workbook calculation mode,
        workbook security, themes, custom properties, named styles, and other
        workbook-wide resources.

        Raises:
            RuntimeError:
                If no workbook exists.
        """
        if self._workbook is None:
            raise RuntimeError("Workbook has not been created.")

        self._logger.debug(
            "%s Initializing workbook.",
            _LOG_MESSAGE_PREFIX,
        )

    def remove_default_sheet(self) -> None:
        """Remove the automatically created worksheet.

        ``openpyxl`` creates a worksheet named ``Sheet`` whenever a workbook is
        instantiated. This application manages worksheet creation explicitly, so
        the automatically generated sheet is removed before custom worksheets
        are added.

        Raises:
            RuntimeError:
                If the workbook has not yet been created.
        """
        if self._workbook is None:
            raise RuntimeError("Workbook has not been created.")

        self._logger.debug(
            "%s Removing default worksheet if present.",
            _LOG_MESSAGE_PREFIX,
        )

        try:
            if len(self._workbook.worksheets) != 1:
                return

            default_sheet = self._workbook.active

            if default_sheet is not None:
                self._workbook.remove(default_sheet)

        except Exception:
            self._logger.exception(
                "%s Failed while removing default worksheet.",
                _LOG_MESSAGE_PREFIX,
            )
            raise

    def set_workbook_properties(self) -> None:
        """Populate workbook document properties.

        Metadata originates from :class:`WorkbookConfig`.

        Future implementations may also configure custom document properties,
        categories, keywords, revision information, and application metadata.

        Raises:
            RuntimeError:
                If no workbook exists.
        """
        if self._workbook is None:
            raise RuntimeError("Workbook has not been created.")

        self._logger.debug(
            "%s Applying workbook properties.",
            _LOG_MESSAGE_PREFIX,
        )

        try:
            properties = self._workbook.properties

            properties.title = self._config.workbook_title
            properties.creator = self._config.author
            properties.lastModifiedBy = self._config.author
            properties.company = self._config.company
            properties.subject = self._config.tournament_name
            properties.category = "Cricket Tournament"

        except Exception:
            self._logger.exception(
                "%s Failed to apply workbook properties.",
                _LOG_MESSAGE_PREFIX,
            )
            raise

    def create_core_sheets(self) -> None:
        """Create all visible worksheets.

        The foundational implementation intentionally performs no worksheet
        creation. Future parts of this module will construct every visible
        worksheet in the required order and register each sheet inside the
        worksheet registry.

        Raises:
            RuntimeError:
                If the workbook has not been created.
        """
        if self._workbook is None:
            raise RuntimeError("Workbook has not been created.")

        self._logger.debug(
            "%s Core worksheet creation phase started.",
            _LOG_MESSAGE_PREFIX,
        )

    def create_hidden_sheets(self) -> None:
        """Create internal worksheets.

        Hidden worksheets will later store lookup values, validation lists,
        configuration data, and other internal resources required by the
        workbook.

        Raises:
            RuntimeError:
                If the workbook has not been created.
        """
        if self._workbook is None:
            raise RuntimeError("Workbook has not been created.")

        self._logger.debug(
            "%s Hidden worksheet creation phase started.",
            _LOG_MESSAGE_PREFIX,
        )

    def initialize_internal_state(self) -> None:
        """Reset internal builder state.

        This method prepares internal registries and caches before workbook
        construction begins. Calling it allows a single builder instance to be
        safely reused for constructing multiple workbooks sequentially.

        Raises:
            RuntimeError:
                If the workbook has not been created.
        """
        if self._workbook is None:
            raise RuntimeError("Workbook has not been created.")

        self._logger.debug(
            "%s Initializing internal state.",
            _LOG_MESSAGE_PREFIX,
        )

        self._worksheet_registry.clear()

        helper_registry = getattr(self, "_helper_registry", None)
        if isinstance(helper_registry, dict):
            helper_registry.clear()


###############################################################################
# WorkbookBuilder - style initialization
###############################################################################

    def initialize_styles(self) -> None:
        """Initialize workbook styling infrastructure.

        This method prepares all styling resources required by the workbook.

        The implementation intentionally delegates styling responsibilities to
        the project's :mod:`styles` module rather than embedding formatting
        logic directly inside :class:`WorkbookBuilder`. This keeps workbook
        orchestration independent from visual presentation.

        Raises:
            RuntimeError:
                If the workbook has not been created.

            ImportError:
                If the project's ``styles`` module cannot be imported.
        """
        if self._workbook is None:
            raise RuntimeError("Workbook has not been created.")

        self._logger.debug(
            "%s Initializing workbook styles.",
            _LOG_MESSAGE_PREFIX,
        )

        try:
            import styles
        except ImportError:
            self._logger.exception(
                "%s Unable to import styles module.",
                _LOG_MESSAGE_PREFIX,
            )
            raise

        self._helper_registry["styles"] = styles

        self.register_named_styles()

    def register_named_styles(self) -> None:
        """Register workbook named styles.

        Registration is delegated to ``styles.py`` whenever a compatible
        registration function is available.

        Supported entry points (checked in order):

        * ``register_named_styles(workbook)``
        * ``initialize_styles(workbook)``

        If neither entry point exists, this method simply records the condition
        in the log without failing. This allows the styling module to evolve
        independently while maintaining backward compatibility.

        Raises:
            RuntimeError:
                If the workbook has not been created.
        """
        if self._workbook is None:
            raise RuntimeError("Workbook has not been created.")

        styles_module = self._helper_registry.get("styles")

        if styles_module is None:
            raise RuntimeError(
                "Style subsystem has not been initialized."
            )

        self._logger.debug(
            "%s Registering named styles.",
            _LOG_MESSAGE_PREFIX,
        )

        register = getattr(styles_module, "register_named_styles", None)

        if callable(register):
            register(self._workbook)
            return

        initialize = getattr(styles_module, "initialize_styles", None)

        if callable(initialize):
            initialize(self._workbook)
            return

        self._logger.info(
            "%s No named-style registration function found.",
            _LOG_MESSAGE_PREFIX,
        )

    def apply_default_style(
        self,
        worksheet: WorksheetType,
    ) -> None:
        """Apply the project's default worksheet style.

        Args:
            worksheet:
                Worksheet receiving the default appearance.

        Styling is delegated to ``styles.py`` using one of the following
        supported entry points:

        * ``apply_default_style(worksheet, config)``
        * ``apply_default_style(worksheet)``
        * ``style_worksheet(worksheet)``
        """
        styles_module = self._helper_registry.get("styles")

        if styles_module is None:
            raise RuntimeError(
                "Style subsystem has not been initialized."
            )

        self._logger.debug(
            "%s Applying default style to worksheet '%s'.",
            _LOG_MESSAGE_PREFIX,
            worksheet.title,
        )

        apply = getattr(styles_module, "apply_default_style", None)

        if callable(apply):
            try:
                apply(worksheet, self._config)
            except TypeError:
                apply(worksheet)
            return

        worksheet_style = getattr(styles_module, "style_worksheet", None)

        if callable(worksheet_style):
            worksheet_style(worksheet)
            return

        self._logger.info(
            "%s No worksheet default style function available.",
            _LOG_MESSAGE_PREFIX,
        )

    def style_cell(
        self,
        cell: CellType,
        style_name: str,
    ) -> None:
        """Apply a named style to a single cell.

        Args:
            cell:
                Target cell.

            style_name:
                Logical style identifier defined by ``styles.py``.
        """
        styles_module = self._helper_registry.get("styles")

        if styles_module is None:
            raise RuntimeError(
                "Style subsystem has not been initialized."
            )

        styler = getattr(styles_module, "style_cell", None)

        if not callable(styler):
            raise AttributeError(
                "styles.style_cell() is not available."
            )

        styler(cell, style_name)

    def style_range(
        self,
        worksheet: WorksheetType,
        cell_range: str,
        style_name: str,
    ) -> None:
        """Apply a named style to a rectangular cell range.

        Args:
            worksheet:
                Target worksheet.

            cell_range:
                Excel range notation.

            style_name:
                Logical style identifier.
        """
        styles_module = self._helper_registry.get("styles")

        if styles_module is None:
            raise RuntimeError(
                "Style subsystem has not been initialized."
            )

        styler = getattr(styles_module, "style_range", None)

        if not callable(styler):
            raise AttributeError(
                "styles.style_range() is not available."
            )

        styler(worksheet, cell_range, style_name)

    def style_row(
        self,
        worksheet: WorksheetType,
        row_index: int,
        style_name: str,
    ) -> None:
        """Apply a named style to an entire worksheet row.

        Args:
            worksheet:
                Target worksheet.

            row_index:
                One-based Excel row number.

            style_name:
                Logical style identifier.
        """
        styles_module = self._helper_registry.get("styles")

        if styles_module is None:
            raise RuntimeError(
                "Style subsystem has not been initialized."
            )

        styler = getattr(styles_module, "style_row", None)

        if not callable(styler):
            raise AttributeError(
                "styles.style_row() is not available."
            )

        styler(worksheet, row_index, style_name)

    def style_column(
        self,
        worksheet: WorksheetType,
        column: int | str,
        style_name: str,
    ) -> None:
        """Apply a named style to an entire worksheet column.

        Args:
            worksheet:
                Target worksheet.

            column:
                Excel column letter or one-based column index.

            style_name:
                Logical style identifier.
        """
        styles_module = self._helper_registry.get("styles")

        if styles_module is None:
            raise RuntimeError(
                "Style subsystem has not been initialized."
            )

        styler = getattr(styles_module, "style_column", None)

        if not callable(styler):
            raise AttributeError(
                "styles.style_column() is not available."
            )

        styler(worksheet, column, style_name)


###############################################################################
# WorkbookBuilder - Theme support
###############################################################################

    def apply_theme(self) -> None:
        """Apply the configured workbook theme.

        This method prepares workbook-wide visual defaults used by later
        worksheet builders. It does not directly format worksheets or cells.
        Instead, it creates a centralized theme definition that other builder
        methods and the project's ``styles`` module can consume.

        Theme application is intentionally lightweight:

        * validates the configured theme
        * integrates with ``constants.py`` where available
        * reuses ``styles.py`` for theme registration if supported
        * prepares cached default style definitions

        Raises:
            RuntimeError:
                If the workbook has not yet been created.

            ValueError:
                If the configured theme is invalid.
        """
        if self._workbook is None:
            raise RuntimeError("Workbook has not been created.")

        self._logger.debug(
            "%s Applying workbook theme '%s'.",
            _LOG_MESSAGE_PREFIX,
            self._config.theme_name,
        )

        theme_name = self._config.theme_name.strip()

        if not theme_name:
            raise ValueError("Theme name cannot be empty.")

        theme: dict[str, object] = {
            "name": theme_name,
            "colors": self._theme_color_mapping(),
            "fonts": self._theme_font_defaults(),
            "borders": self._theme_border_defaults(),
            "fills": self._theme_fill_defaults(),
            "alignment": self._theme_alignment_defaults(),
            "number_formats": self._theme_number_format_defaults(),
        }

        self._helper_registry["theme"] = theme

        styles_module = self._helper_registry.get("styles")

        if styles_module is not None:
            register = getattr(styles_module, "apply_theme", None)

            if callable(register):
                try:
                    register(self._workbook, theme)
                except TypeError:
                    register(theme)

        self._logger.info(
            "%s Theme '%s' initialized.",
            _LOG_MESSAGE_PREFIX,
            theme_name,
        )

    def _theme_color_mapping(self) -> dict[str, str]:
        """Build the workbook theme colour mapping.

        The method first attempts to reuse colour definitions exposed by
        ``constants.py``. If no compatible mapping exists, conservative Excel
        defaults are used.

        Returns:
            Mapping of logical colour names to RGB values.
        """
        mapping = getattr(constants, "THEME_COLORS", None)

        if isinstance(mapping, dict):
            return dict(mapping)

        return {
            "primary": "1F4E78",
            "secondary": "D9EAD3",
            "accent": "4F81BD",
            "success": "70AD47",
            "warning": "FFC000",
            "danger": "C00000",
            "header_background": "D9E2F3",
            "header_text": "000000",
            "grid": "D9D9D9",
            "text": "000000",
            "background": "FFFFFF",
        }

    def _theme_font_defaults(self) -> dict[str, object]:
        """Create workbook font defaults.

        Returns:
            Dictionary describing default workbook font settings.
        """
        return {
            "family": self._config.default_font,
            "size": self._config.default_font_size,
            "bold": False,
            "italic": False,
            "underline": None,
            "strike": False,
            "color": "000000",
        }

    def _theme_border_defaults(self) -> dict[str, object]:
        """Create workbook border defaults.

        Returns:
            Border configuration consumed by later styling phases.
        """
        return {
            "style": "thin",
            "color": "D9D9D9",
            "outline": True,
        }

    def _theme_fill_defaults(self) -> dict[str, object]:
        """Create workbook fill defaults.

        Returns:
            Fill configuration for reusable workbook styles.
        """
        return {
            "pattern": "solid",
            "foreground": "FFFFFF",
            "background": "FFFFFF",
        }

    def _theme_alignment_defaults(self) -> dict[str, object]:
        """Create workbook alignment defaults.

        Returns:
            Alignment configuration shared throughout the workbook.
        """
        return {
            "horizontal": "center",
            "vertical": "center",
            "wrap_text": True,
            "shrink_to_fit": False,
        }

    def _theme_number_format_defaults(self) -> dict[str, str]:
        """Create workbook number-format defaults.

        Returns:
            Mapping of logical format names to Excel number formats.
        """
        return {
            "date": self._config.date_format,
            "time": self._config.time_format,
            "datetime": self._config.datetime_format,
            "integer": "0",
            "decimal": "0.00",
            "percentage": "0.00%",
            "currency": "#,##0.00",
            "text": "@",
        }


###############################################################################
# WorkbookBuilder - Named style registration
###############################################################################

from openpyxl.styles import NamedStyle


    def register_header_style(self) -> NamedStyle:
        """Register the standard header style.

        The style is registered only once per workbook. Subsequent calls return
        the existing style instance.

        Returns:
            The registered :class:`~openpyxl.styles.NamedStyle`.

        Raises:
            RuntimeError:
                If the workbook has not been created.
        """
        return self._register_named_style(
            style_name="Header",
            number_format="General",
            style_callback="create_header_named_style",
        )

    def register_title_style(self) -> NamedStyle:
        """Register the workbook title style.

        Returns:
            The registered title style.
        """
        return self._register_named_style(
            style_name="Title",
            number_format="General",
            style_callback="create_title_named_style",
        )

    def register_normal_style(self) -> NamedStyle:
        """Register the default workbook style.

        Returns:
            The registered normal style.
        """
        return self._register_named_style(
            style_name="Normal",
            number_format="General",
            style_callback="create_normal_named_style",
        )

    def register_number_style(self) -> NamedStyle:
        """Register the standard numeric style.

        Returns:
            The registered numeric style.
        """
        return self._register_named_style(
            style_name="Number",
            number_format=self._theme_number_format_defaults()["integer"],
            style_callback="create_number_named_style",
        )

    def register_percentage_style(self) -> NamedStyle:
        """Register the percentage style.

        Returns:
            The registered percentage style.
        """
        return self._register_named_style(
            style_name="Percentage",
            number_format=self._theme_number_format_defaults()["percentage"],
            style_callback="create_percentage_named_style",
        )

    def register_date_style(self) -> NamedStyle:
        """Register the date style.

        Returns:
            The registered date style.
        """
        return self._register_named_style(
            style_name="Date",
            number_format=self._config.date_format,
            style_callback="create_date_named_style",
        )

    def register_time_style(self) -> NamedStyle:
        """Register the time style.

        Returns:
            The registered time style.
        """
        return self._register_named_style(
            style_name="Time",
            number_format=self._config.time_format,
            style_callback="create_time_named_style",
        )

    def register_currency_style(self) -> NamedStyle:
        """Register the currency style.

        Returns:
            The registered currency style.
        """
        return self._register_named_style(
            style_name="Currency",
            number_format=self._theme_number_format_defaults()["currency"],
            style_callback="create_currency_named_style",
        )

    def register_hidden_style(self) -> NamedStyle:
        """Register the hidden/internal worksheet style.

        Returns:
            The registered hidden style.
        """
        return self._register_named_style(
            style_name="Hidden",
            number_format="General",
            style_callback="create_hidden_named_style",
        )

    def _register_named_style(
        self,
        *,
        style_name: str,
        number_format: str,
        style_callback: str,
    ) -> NamedStyle:
        """Register a named style if it does not already exist.

        The method first attempts to reuse style factory functions from
        ``styles.py``. If no compatible factory exists, a minimal
        :class:`~openpyxl.styles.NamedStyle` is created using the active workbook
        theme defaults.

        Args:
            style_name:
                Excel named style name.

            number_format:
                Default Excel number format.

            style_callback:
                Factory function expected inside ``styles.py``.

        Returns:
            The registered named style.

        Raises:
            RuntimeError:
                If the workbook has not been created.
        """
        if self._workbook is None:
            raise RuntimeError("Workbook has not been created.")

        for existing in self._workbook.named_styles:
            if isinstance(existing, NamedStyle):
                if existing.name == style_name:
                    return existing
            elif existing == style_name:
                for candidate in self._workbook._named_styles:
                    if candidate.name == style_name:
                        return candidate

        styles_module = self._helper_registry.get("styles")

        style: NamedStyle | None = None

        if styles_module is not None:
            factory = getattr(styles_module, style_callback, None)

            if callable(factory):
                try:
                    style = factory(self._config)
                except TypeError:
                    style = factory()

        if style is None:
            theme = self._helper_registry.get("theme", {})

            fonts = theme.get("fonts", {})
            borders = theme.get("borders", {})
            fills = theme.get("fills", {})
            alignment = theme.get("alignment", {})

            style = NamedStyle(name=style_name)

            style.font = Font(
                name=fonts.get("family", self._config.default_font),
                size=fonts.get("size", self._config.default_font_size),
                bold=fonts.get("bold", False),
                italic=fonts.get("italic", False),
                color=fonts.get("color", "000000"),
            )

            style.fill = PatternFill(
                fill_type=fills.get("pattern", "solid"),
                fgColor=fills.get("foreground", "FFFFFF"),
                bgColor=fills.get("background", "FFFFFF"),
            )

            side = Side(
                style=borders.get("style", "thin"),
                color=borders.get("color", "D9D9D9"),
            )

            style.border = Border(
                left=side,
                right=side,
                top=side,
                bottom=side,
            )

            style.alignment = Alignment(
                horizontal=alignment.get("horizontal", "center"),
                vertical=alignment.get("vertical", "center"),
                wrap_text=alignment.get("wrap_text", True),
                shrink_to_fit=alignment.get("shrink_to_fit", False),
            )

            style.number_format = number_format

        if style.name != style_name:
            style.name = style_name

        self._workbook.add_named_style(style)

        self._logger.debug(
            "%s Registered named style '%s'.",
            _LOG_MESSAGE_PREFIX,
            style_name,
        )

        return style


###############################################################################
# WorkbookBuilder - Reusable column utilities
###############################################################################

    def set_column_width(
        self,
        worksheet: WorksheetType,
        column: int | str,
        width: float,
    ) -> None:
        """Set the width of a single worksheet column.

        Args:
            worksheet:
                Target worksheet.

            column:
                Excel column reference. May be either a one-based integer column
                index or an Excel column letter.

            width:
                Desired column width.

        Raises:
            ValueError:
                If the column reference or width is invalid.
        """
        column_letter = self.validate_column_reference(column)

        if width <= 0:
            raise ValueError("Column width must be greater than zero.")

        worksheet.column_dimensions[column_letter].width = width

        self._logger.debug(
            "%s Set width of column '%s' to %.2f.",
            _LOG_MESSAGE_PREFIX,
            column_letter,
            width,
        )

    def set_multiple_column_widths(
        self,
        worksheet: WorksheetType,
        widths: dict[int | str, float],
    ) -> None:
        """Apply widths to multiple worksheet columns.

        Args:
            worksheet:
                Target worksheet.

            widths:
                Mapping of column reference to width.
        """
        for column, width in widths.items():
            self.set_column_width(
                worksheet=worksheet,
                column=column,
                width=width,
            )

    def autosize_columns(
        self,
        worksheet: WorksheetType,
        *,
        minimum_width: float = 8.43,
        maximum_width: float = 60.0,
        padding: float = 2.0,
    ) -> None:
        """Automatically size populated worksheet columns.

        The calculation is based on the string representation of each cell
        value. This helper intentionally provides a generic implementation and
        does not contain worksheet-specific assumptions.

        Args:
            worksheet:
                Target worksheet.

            minimum_width:
                Minimum permitted column width.

            maximum_width:
                Maximum permitted column width.

            padding:
                Extra width added after measuring the largest cell value.
        """
        self._logger.debug(
            "%s Autosizing worksheet '%s'.",
            _LOG_MESSAGE_PREFIX,
            worksheet.title,
        )

        for column_cells in worksheet.columns:
            try:
                first_cell = next(iter(column_cells))
            except StopIteration:
                continue

            column_letter = first_cell.column_letter

            max_length = 0

            for cell in column_cells:
                if cell.value is None:
                    continue

                length = len(str(cell.value))

                if length > max_length:
                    max_length = length

            calculated_width = max(
                minimum_width,
                min(maximum_width, max_length + padding),
            )

            worksheet.column_dimensions[column_letter].width = calculated_width

    def hide_columns(
        self,
        worksheet: WorksheetType,
        *columns: int | str,
    ) -> None:
        """Hide one or more worksheet columns.

        Args:
            worksheet:
                Target worksheet.

            *columns:
                Column references.
        """
        for column in columns:
            column_letter = self.validate_column_reference(column)
            worksheet.column_dimensions[column_letter].hidden = True

    def unhide_columns(
        self,
        worksheet: WorksheetType,
        *columns: int | str,
    ) -> None:
        """Unhide one or more worksheet columns.

        Args:
            worksheet:
                Target worksheet.

            *columns:
                Column references.
        """
        for column in columns:
            column_letter = self.validate_column_reference(column)
            worksheet.column_dimensions[column_letter].hidden = False

    def group_columns(
        self,
        worksheet: WorksheetType,
        start_column: int | str,
        end_column: int | str,
        *,
        hidden: bool = False,
        outline_level: int = 1,
    ) -> None:
        """Create an Excel outline group for a range of columns.

        Args:
            worksheet:
                Target worksheet.

            start_column:
                First column in the group.

            end_column:
                Last column in the group.

            hidden:
                Whether the grouped columns should initially be hidden.

            outline_level:
                Excel outline level.

        Raises:
            ValueError:
                If the outline level is invalid.
        """
        if outline_level < 1:
            raise ValueError("outline_level must be at least 1.")

        start = self.validate_column_reference(start_column)
        end = self.validate_column_reference(end_column)

        worksheet.column_dimensions.group(
            start=start,
            end=end,
            hidden=hidden,
            outline_level=outline_level,
        )

    def copy_column_widths(
        self,
        source: WorksheetType,
        destination: WorksheetType,
    ) -> None:
        """Copy all explicit column widths between worksheets.

        Args:
            source:
                Worksheet providing the widths.

            destination:
                Worksheet receiving the widths.
        """
        for column_letter, dimension in source.column_dimensions.items():
            destination.column_dimensions[column_letter].width = dimension.width
            destination.column_dimensions[column_letter].hidden = (
                dimension.hidden
            )
            destination.column_dimensions[column_letter].bestFit = (
                dimension.bestFit
            )
            destination.column_dimensions[column_letter].outlineLevel = (
                dimension.outlineLevel
            )

        self._logger.debug(
            "%s Copied column widths from '%s' to '%s'.",
            _LOG_MESSAGE_PREFIX,
            source.title,
            destination.title,
        )

    def validate_column_reference(
        self,
        column: int | str,
    ) -> str:
        """Validate and normalize a column reference.

        Accepted formats are:

        * one-based integer column indices
        * Excel column letters (for example ``A`` or ``AA``)

        Args:
            column:
                Column reference.

        Returns:
            Normalized uppercase Excel column letter.

        Raises:
            TypeError:
                If the supplied type is unsupported.

            ValueError:
                If the reference is outside Excel's supported limits.
        """
        if isinstance(column, int):
            if not (1 <= column <= EXCEL_MAX_COLUMNS):
                raise ValueError(
                    f"Column index must be between 1 and {EXCEL_MAX_COLUMNS}."
                )

            return get_column_letter(column)

        if isinstance(column, str):
            normalized = column.strip().upper()

            if not normalized.isalpha():
                raise ValueError(
                    "Column reference must contain only letters."
                )

            return normalized

        raise TypeError(
            "Column reference must be an integer or Excel column letter."
        )