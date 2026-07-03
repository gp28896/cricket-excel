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