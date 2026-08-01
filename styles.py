"""
styles.py
=========

Reusable NamedStyle registration and styling helpers for the Cricket
Tournament Excel system.

Overview
--------
This module is responsible for defining and registering the reusable
``openpyxl`` styling building blocks — fonts, fills, borders, alignments,
and ``NamedStyle`` objects — used across the workbook.

Unlike higher-level modules such as ``workbook.py``, ``create_match.py``,
or ``scorecard.py``, which orchestrate worksheet construction and business
logic, this module focuses solely on the visual presentation layer. It
provides the shared styling vocabulary that other modules apply to cells,
ranges, and worksheets rather than each module defining its own ad-hoc
formatting.

Responsibilities
-----------------
The responsibilities of this module include (implemented incrementally
across future parts of this file):

* NamedStyle definition and registration
* Font, fill, border, and alignment presets
* Header, title, and label styling helpers
* Conditional formatting style presets
* Theme/color palette constants
* Style application helpers for cells and ranges
* Protection style presets
* Number format presets

Compatibility
--------------
This module has been designed to integrate with:

* ``constants.py`` -- application constants and configuration values
* ``workbook.py`` -- delegates NamedStyle registration and cell styling
  to this module
* ``create_match.py`` -- consumes styling helpers via ``from styles
  import *``
* ``scorecard.py`` -- consumes styling helpers for scorecard formatting

This module has no dependency on ``formulas.py`` or ``validation.py``
and does not perform any workbook orchestration, business logic, or data
validation itself.

Design Philosophy
------------------
The architecture intentionally separates styling from orchestration.

Every visual style used in the workbook should be defined once, here, and
reused everywhere it is needed rather than being redefined inline inside
worksheet-building code. This design provides:

* a single source of truth for visual presentation
* consistent look and feel across all worksheets
* easier theme-wide changes
* simpler testing of style definitions in isolation
* reduced duplication across modules

Nothing in this initial module foundation performs style registration.
This file currently establishes the shared imports and module-level
scaffolding used throughout later sections.
"""

from __future__ import annotations

###############################################################################
# Third-party imports
###############################################################################

from openpyxl.styles import Alignment
from openpyxl.styles import Border
from openpyxl.styles import Font
from openpyxl.styles import NamedStyle
from openpyxl.styles import PatternFill
from openpyxl.styles import Protection
from openpyxl.styles import Side

###############################################################################
# Public API
###############################################################################

__all__: list[str] = []