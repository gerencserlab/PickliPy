from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Any, List, Optional, Sequence, Tuple

from .common import PicklyPyConfigError, as_number, as_str, round_to_increment


BACKFILL_LABEL = "backfill"


@dataclass(frozen=True)
class BackfillOverride:
    compound_name: str
    target_volume_nl: float


def normalize_backfill_override(
    compound_name: Optional[str],
    target_volume_nl: Optional[float],
) -> Optional[BackfillOverride]:
    """Validate and normalize the paired command/API backfill override."""

    name = as_str(compound_name).strip() if compound_name is not None else ""
    has_name = name != ""
    has_volume = target_volume_nl is not None
    if has_name != has_volume:
        raise PicklyPyConfigError(
            "Backfill override requires both --backfill_name and --backfill_volume."
        )
    if not has_name:
        return None

    volume = as_number(target_volume_nl)
    if volume is None or not math.isfinite(float(volume)) or float(volume) <= 0:
        raise PicklyPyConfigError(
            "--backfill_volume must be a positive finite volume in nL."
        )

    rounded = float(round_to_increment(float(volume), 2.5))
    if rounded <= 0:
        raise PicklyPyConfigError(
            "--backfill_volume rounds to 0 nL at the Echo 2.5 nL increment."
        )
    return BackfillOverride(name, rounded)


def reserved_backfill_row_indices(
    label_rows: Sequence[Sequence[Any]],
    used_plate_map_labels: Sequence[str],
) -> List[int]:
    """Return reserved Backfill rows when Backfill is absent from the plate map.

    A case-insensitive ``backfill`` token used in the plate map keeps every
    matching Labels-table row in its ordinary treatment-label role.
    """

    if any(as_str(label).strip().casefold() == BACKFILL_LABEL for label in used_plate_map_labels):
        return []
    return [
        index
        for index, row in enumerate(label_rows)
        if row and as_str(row[0]).strip().casefold() == BACKFILL_LABEL
    ]


def validate_single_workbook_backfill(
    label_rows: Sequence[Sequence[Any]],
    reserved_indices: Sequence[int],
    *,
    context: str,
) -> Optional[Tuple[str, float]]:
    """Read ``Backfill | compound | mock final concentration`` configuration."""

    if not reserved_indices:
        return None
    if len(reserved_indices) != 1:
        raise PicklyPyConfigError(
            f"{context} has more than one unused Backfill row in its Labels table. "
            "Keep exactly one Backfill configuration row."
        )

    row = label_rows[reserved_indices[0]]
    compound = as_str(row[1] if len(row) > 1 else "").strip()
    mock_final = as_number(row[2] if len(row) > 2 else None)
    if compound == "":
        raise PicklyPyConfigError(
            f"{context} Backfill row must name a compound present in SRC."
        )
    if mock_final is None or not math.isfinite(float(mock_final)) or float(mock_final) <= 0:
        raise PicklyPyConfigError(
            f"{context} Backfill row must provide a positive finite mock final concentration."
        )
    return compound, float(mock_final)


def target_volume_from_mock_concentration(
    *,
    well_volume_ul: float,
    stock_concentration: float,
    mock_final_concentration: float,
    context: str,
) -> Tuple[float, float]:
    """Return finite desired/rounded target volumes for workbook backfill."""

    values = {
        "destination well volume": float(well_volume_ul),
        "source stock concentration": float(stock_concentration),
        "mock final concentration": float(mock_final_concentration),
    }
    invalid = [name for name, value in values.items() if not math.isfinite(value) or value <= 0]
    if invalid:
        raise PicklyPyConfigError(
            f"{context} requires a positive finite {', '.join(invalid)}."
        )

    desired = (
        1000.0
        * values["destination well volume"]
        * values["mock final concentration"]
        / values["source stock concentration"]
    )
    if not math.isfinite(desired) or desired <= 0:
        raise PicklyPyConfigError(
            f"{context} produces a non-finite or nonpositive target transfer volume."
        )
    rounded = float(round_to_increment(desired, 2.5))
    if rounded <= 0:
        raise PicklyPyConfigError(
            f"{context} target rounds to 0 nL at the Echo 2.5 nL increment."
        )
    return desired, rounded


def backfill_remainders(
    normal_events: Sequence[Tuple[str, Sequence[str], float]],
    target_volume_nl: float,
) -> List[Tuple[str, float]]:
    """Calculate top-up volumes for wells with positive normal transfers only."""

    totals: dict[str, float] = {}
    order: List[str] = []
    for dst_well, _source_wells, volume_nl in normal_events:
        volume = float(volume_nl)
        if volume <= 0:
            continue
        if dst_well not in totals:
            totals[dst_well] = 0.0
            order.append(dst_well)
        totals[dst_well] += volume

    out: List[Tuple[str, float]] = []
    for dst_well in order:
        remaining = float(target_volume_nl) - totals[dst_well]
        if remaining <= 0:
            continue
        rounded = float(round_to_increment(remaining, 2.5))
        if rounded > 0:
            out.append((dst_well, rounded))
    return out
