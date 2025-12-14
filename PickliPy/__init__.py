"""PicklyPy: Picklist generation for Echo 650 and Plate:Works.

This package contains two primary entry points:

- PicklyPy.Assay: Tool-library / combinatorics / dose-response / multi-addition workflows.
- PicklyPy.Screen: Library screening + reformatting workflows (layout auto-filled from LIB).

The implementation is a Python translation of the original Wolfram Language scripts
`assaypicklist.wls` and `screeningpicklist.wls`.

Written for the Gerencser Lab HTS workflows.
"""

from __future__ import annotations

__all__ = [
    "__version__",
]

__version__ = "0.1.0"
