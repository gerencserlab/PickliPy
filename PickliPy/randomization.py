from __future__ import annotations

"""Well randomization utilities for PickliPy.Assay.

This module is intentionally self-contained so that alternative randomization
strategies can be implemented later without touching the main Assay workflow.

The core concept is a *well map* (a permutation): a dict mapping
    original_well -> randomized_well

The Assay workflow applies this mapping consistently across all plate maps that
target the same destination plate barcode so that the *entire treatment*
(potentially spanning multiple sequential additions) moves together.

Randomization is restricted to the set of wells that receive any treatment.
If a layout map is provided, wells are randomized *within each layout label*
independently (no movement across different layout labels).

The implemented heuristic tries to spread replicates (wells with identical
treatment) apart by approximately maximizing their pairwise Euclidean distance
and preventing orthogonally adjacent replicates.
"""

import hashlib
import random
from typing import Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

from .common import PLATE_COLS_384, PLATE_ROWS_384, euclidean, parse_well_name, well_name


def seed_from_text(text: str) -> int:
    """Create a stable, platform-independent 64-bit integer seed from text."""

    # Using sha256 avoids Python's randomized hash() salt.
    digest = hashlib.sha256(text.encode("utf-8")).digest()
    return int.from_bytes(digest[:8], byteorder="big", signed=False)


def apply_well_map_to_str_grid(grid: Sequence[Sequence[str]], well_map: Mapping[str, str]) -> List[List[str]]:
    """Apply a well mapping to a 16x24 string grid.

    Each cell content moves together with its original well.

    Parameters
    ----------
    grid:
        16x24 table indexed [row][col] for a 384-well plate.
    well_map:
        Mapping original_well -> new_well. Missing keys are treated as identity.
    """

    out: List[List[str]] = [["" for _ in range(PLATE_COLS_384)] for _ in range(PLATE_ROWS_384)]
    for r in range(PLATE_ROWS_384):
        for c in range(PLATE_COLS_384):
            old_well = well_name(r + 1, c + 1)
            new_well = well_map.get(old_well, old_well)
            rr, cc = parse_well_name(new_well)  # 1-based
            out[rr - 1][cc - 1] = grid[r][c]
    return out


def apply_well_map_to_well_dict(values_by_well: Mapping[str, float], well_map: Mapping[str, str]) -> Dict[str, float]:
    """Apply a well mapping to a dict keyed by well names."""
    return {well_map.get(w, w): v for w, v in values_by_well.items()}


def _orthogonally_adjacent(a: Tuple[int, int], b: Tuple[int, int]) -> bool:
    """Return True if wells a and b are orthogonally adjacent (sharing an edge)."""
    dr = abs(a[0] - b[0])
    dc = abs(a[1] - b[1])
    return (dr + dc) == 1


def _min_pairwise_distance(coords: List[Tuple[int, int]]) -> float:
    if len(coords) < 2:
        return 0.0
    best = float("inf")
    for i in range(len(coords)):
        for j in range(i + 1, len(coords)):
            d = euclidean(coords[i], coords[j])
            if d < best:
                best = d
    return 0.0 if best == float("inf") else float(best)


def _has_adjacent_pair(coords: List[Tuple[int, int]]) -> bool:
    if len(coords) < 2:
        return False
    for i in range(len(coords)):
        for j in range(i + 1, len(coords)):
            if _orthogonally_adjacent(coords[i], coords[j]):
                return True
    return False


def _score_assignment(
    *,
    assigned_wells_by_treatment: Mapping[str, Sequence[str]],
    coords_by_well: Mapping[str, Tuple[int, int]],
) -> Tuple[int, float]:
    """Return (adjacency_violations, score).

    score is the sum of minimum pairwise distances for each replicated treatment.
    adjacency_violations counts how many replicated treatments end up with at
    least one orthogonally adjacent replicate pair.
    """

    violations = 0
    score = 0.0
    for _treat, wells in assigned_wells_by_treatment.items():
        if len(wells) <= 1:
            continue
        coords = [coords_by_well[w] for w in wells]
        if _has_adjacent_pair(coords):
            violations += 1
        score += _min_pairwise_distance(coords)
    return violations, score


def _choose_farthest_well(
    *,
    candidates: Sequence[str],
    chosen: Sequence[str],
    coords_by_well: Mapping[str, Tuple[int, int]],
) -> str:
    """Choose candidate maximizing its minimum distance to already chosen wells."""
    if not chosen:
        return candidates[0]
    best_well = candidates[0]
    best_d = -1.0
    for w in candidates:
        dmin = min(euclidean(coords_by_well[w], coords_by_well[x]) for x in chosen)
        if dmin > best_d:
            best_d = dmin
            best_well = w
    return best_well


def _attempt_group_assignment(
    *,
    wells: Sequence[str],
    treatment_by_well: Mapping[str, str],
    rng: random.Random,
    forbid_adjacent_replicates: bool,
) -> Optional[Dict[str, List[str]]]:
    """Greedy attempt to assign treatments to wells.

    Returns mapping treatment -> list of assigned wells, or None if the
    orthogonal-adjacency constraint cannot be satisfied in this attempt.
    """

    coords_by_well = {w: parse_well_name(w) for w in wells}

    # Build treatment -> original wells (instances)
    orig_by_treat: Dict[str, List[str]] = {}
    for w in wells:
        t = treatment_by_well.get(w, "")
        orig_by_treat.setdefault(t, []).append(w)

    # Stable list of treatments
    treatments = list(orig_by_treat.keys())

    # Split into replicated and singleton treatments
    replicated: List[str] = [t for t in treatments if len(orig_by_treat[t]) > 1]
    singletons: List[str] = [t for t in treatments if len(orig_by_treat[t]) == 1]

    # Process replicated treatments first (largest groups first, randomized tie-break)
    replicated.sort(key=lambda t: (-len(orig_by_treat[t]), rng.random()))

    available = list(wells)
    rng.shuffle(available)

    assigned: Dict[str, List[str]] = {}

    # Place replicates using farthest-point sampling with an adjacency constraint
    for t in replicated:
        k = len(orig_by_treat[t])
        if k > len(available):
            return None

        chosen: List[str] = []

        # First pick: random
        first = rng.choice(available)
        chosen.append(first)
        available.remove(first)

        while len(chosen) < k:
            # Candidate wells not adjacent (orthogonal) to any already chosen
            cand = [
                w
                for w in available
                if not any(_orthogonally_adjacent(coords_by_well[w], coords_by_well[x]) for x in chosen)
            ]

            if not cand:
                if forbid_adjacent_replicates:
                    return None
                cand = list(available)

            # Choose well maximizing distance to the already chosen replicate wells
            cand.sort()  # deterministic iteration order for stable tie-break
            w_best = _choose_farthest_well(candidates=cand, chosen=chosen, coords_by_well=coords_by_well)
            chosen.append(w_best)
            available.remove(w_best)

        assigned[t] = chosen

    # Place singleton treatments randomly in remaining wells
    rng.shuffle(singletons)
    rng.shuffle(available)
    if len(singletons) != len(available):
        # This should not happen; indicates a bug in counts.
        return None
    for t, w in zip(singletons, available):
        assigned[t] = [w]

    return assigned


def optimize_plate_well_map(
    *,
    treatment_by_well: Mapping[str, str],
    layout_by_well: Optional[Mapping[str, str]] = None,
    seed: Optional[int] = None,
    attempts: int = 200,
    forbid_adjacent_replicates: bool = True,
) -> Dict[str, str]:
    """Create a well randomization map for a plate.

    Only wells with a non-empty treatment are permuted. Empty (unused) wells map
    to themselves.

    If layout_by_well is provided, randomization is done independently inside
    each layout label value.

    Parameters
    ----------
    treatment_by_well:
        Dict well -> treatment signature (string). Empty or missing means unused.
    layout_by_well:
        Dict well -> layout group label. If None, everything is one group.
    seed:
        RNG seed. If None, a non-deterministic seed is used.
    attempts:
        Number of randomized greedy restarts per layout group.
    forbid_adjacent_replicates:
        If True, replicate wells (same treatment) are forbidden to be
        orthogonally adjacent.
    """

    rng = random.Random(seed)

    # Determine wells that actually receive any treatment.
    used_wells: List[str] = [w for w, t in treatment_by_well.items() if str(t).strip() != ""]
    used_wells = sorted(set(used_wells), key=lambda w: (parse_well_name(w)[0], parse_well_name(w)[1]))

    if not used_wells:
        return {}

    # Group used wells by layout label.
    if layout_by_well is None:
        groups: Dict[str, List[str]] = {"__ALL__": used_wells}
    else:
        groups = {}
        for w in used_wells:
            lab = str(layout_by_well.get(w, ""))
            groups.setdefault(lab, []).append(w)

    final_map: Dict[str, str] = {}

    for _lab, wells in groups.items():
        if len(wells) <= 1:
            # Nothing to permute.
            continue

        # Precompute coords for scoring.
        coords_by_well = {w: parse_well_name(w) for w in wells}

        best_map: Optional[Dict[str, str]] = None
        best_score = -1.0

        # If we cannot satisfy the adjacency constraint, we will fall back to the
        # best (highest score) assignment without the strict constraint.
        best_relaxed_map: Optional[Dict[str, str]] = None
        best_relaxed = (-1, -1.0)  # (violations, score)

        # We reuse the group-level seed stream so the whole plate remains stable.
        for _ in range(max(1, int(attempts))):
            assigned = _attempt_group_assignment(
                wells=wells,
                treatment_by_well=treatment_by_well,
                rng=rng,
                forbid_adjacent_replicates=forbid_adjacent_replicates,
            )
            if assigned is None:
                continue

            violations, score = _score_assignment(
                assigned_wells_by_treatment=assigned,
                coords_by_well=coords_by_well,
            )

            # Build old->new mapping (pair instances per treatment)
            orig_by_treat: Dict[str, List[str]] = {}
            for w in wells:
                orig_by_treat.setdefault(treatment_by_well.get(w, ""), []).append(w)

            mapping: Dict[str, str] = {}
            for t, old_list in orig_by_treat.items():
                new_list = list(assigned.get(t, []))
                if len(old_list) != len(new_list):
                    mapping = {}
                    break
                old_shuf = list(old_list)
                new_shuf = list(new_list)
                rng.shuffle(old_shuf)
                rng.shuffle(new_shuf)
                mapping.update(dict(zip(old_shuf, new_shuf)))

            if not mapping or len(mapping) != len(wells):
                continue

            if forbid_adjacent_replicates and violations == 0:
                if score > best_score:
                    best_score = score
                    best_map = mapping
            else:
                # Relaxed tracking (prefer fewer violations, then higher score)
                if (violations < best_relaxed[0]) or (violations == best_relaxed[0] and score > best_relaxed[1]):
                    best_relaxed = (violations, score)
                    best_relaxed_map = mapping

        if best_map is None:
            if forbid_adjacent_replicates and best_relaxed_map is None:
                # As a last resort, do a pure shuffle (may still violate adjacency).
                shuffled = list(wells)
                rng.shuffle(shuffled)
                best_relaxed_map = dict(zip(wells, shuffled))
                best_relaxed = (0, 0.0)

            if forbid_adjacent_replicates:
                # Still nothing satisfied the strict constraint.
                # Bubble this up as a *user actionable* error.
                raise ValueError(
                    "Could not find a non-adjacent replicate randomization within the requested attempts. "
                    "Consider reducing replicate counts per layout group, loosening the constraint, "
                    "or providing more distinct layout groups." 
                )
            best_map = best_relaxed_map

        if best_map is None:
            continue

        final_map.update(best_map)

    return final_map
