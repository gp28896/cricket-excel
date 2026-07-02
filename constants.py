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

__author__: Final[str] = "Your Name"
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