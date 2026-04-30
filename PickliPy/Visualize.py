from __future__ import annotations

import argparse
import csv
import math
import re
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Mapping, MutableMapping, Optional, Sequence, Tuple

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import colors as mcolors
from matplotlib import patheffects as pe
from matplotlib.cm import ScalarMappable
from matplotlib.patches import Circle, FancyBboxPatch, Wedge


# -----------------------------------------------------------------------------
# Optional PicklyPy compatibility imports
# -----------------------------------------------------------------------------

try:  # pragma: no cover - exercised when the module is installed inside picklipy
    from .common import (  # type: ignore
        PicklyPyConfigError,
        PicklyPyError,
        run_with_user_facing_errors,
    )
except Exception:  # pragma: no cover - exercised in standalone use

    class PicklyPyError(RuntimeError):
        """Base class for user-facing errors."""


    class PicklyPyConfigError(PicklyPyError):
        """Raised when the picklist CSV is malformed or unsupported."""


    def run_with_user_facing_errors(func, *, pause_on_exit: bool = True) -> int:
        try:
            func()
            if pause_on_exit:
                try:
                    input("Press Enter to exit...")
                except EOFError:
                    pass
            return 0
        except PicklyPyError as exc:
            print(str(exc))
            if pause_on_exit:
                try:
                    input("Press Enter to exit...")
                except EOFError:
                    pass
            return 2
        except Exception:
            import traceback

            print("Unexpected error:")
            traceback.print_exc()
            if pause_on_exit:
                try:
                    input("Press Enter to exit...")
                except EOFError:
                    pass
            return 1


# -----------------------------------------------------------------------------
# Constants / helpers
# -----------------------------------------------------------------------------

PLATE_FORMATS: Dict[int, Tuple[int, int]] = {
    96: (8, 12),
    384: (16, 24),
}

_WELL_RE = re.compile(r"^(?P<row>[A-Za-z]+)(?P<col>\d+)$")
_SAFE_FILENAME_RE = re.compile(r"[^A-Za-z0-9._-]+")


@dataclass(frozen=True)
class PlateGeometry:
    wells: int
    rows: int
    cols: int

    @property
    def row_labels(self) -> List[str]:
        return [number_to_excel_col(i) for i in range(1, self.rows + 1)]


GEOMETRY_96 = PlateGeometry(96, 8, 12)
GEOMETRY_384 = PlateGeometry(384, 16, 24)


@dataclass(frozen=True)
class DispenseEvent:
    global_index: int
    row_index: int
    in_row_index: int
    src_plate: str
    src_well: str
    dst_plate: str
    dst_well: str
    volume_nl: float


@dataclass(frozen=True)
class SourceVisit:
    source_plate_visit_index: int
    global_row_index: int
    src_plate: str
    src_well: str
    dst_plate: str
    events: Tuple[DispenseEvent, ...]


@dataclass(frozen=True)
class ParsedPicklist:
    csv_path: Path
    events: Tuple[DispenseEvent, ...]
    visits: Tuple[SourceVisit, ...]

    @property
    def source_plates(self) -> List[str]:
        return list(dict.fromkeys(v.src_plate for v in self.visits))

    @property
    def destination_plates(self) -> List[str]:
        return list(dict.fromkeys(e.dst_plate for e in self.events))


@dataclass(frozen=True)
class VisualizationResult:
    output_dir: Path
    destination_images: Mapping[str, Path]
    source_images: Mapping[str, Path]
    event_count: int


@dataclass(frozen=True)
class RenderStyle:
    cmap_name: str = "viridis"
    plate_face: str = "#262626"
    plate_edge: str = "#111111"
    empty_well_face: str = "#d6f0fb"
    empty_well_edge: str = "#7aaec8"
    neutral_multi_fill: str = "#f7fbff"
    path_color: str = "#38d17a"
    path_width: float = 1.9
    visit_marker_size: float = 10.0
    well_radius: float = 0.41
    multi_inner_radius: float = 0.22
    base_fontsize_96: float = 10.5
    base_fontsize_384: float = 6.2


# -----------------------------------------------------------------------------
# Generic helpers
# -----------------------------------------------------------------------------


def as_str(value: object) -> str:
    if value is None:
        return ""
    return str(value)



def excel_col_to_number(col: str) -> int:
    s = col.strip().upper()
    if not s or not re.fullmatch(r"[A-Z]+", s):
        raise ValueError(f"Invalid Excel column: {col!r}")
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n



def number_to_excel_col(n: int) -> str:
    if n <= 0:
        raise ValueError("n must be >= 1")
    out: List[str] = []
    x = n
    while x > 0:
        x, rem = divmod(x - 1, 26)
        out.append(chr(ord("A") + rem))
    return "".join(reversed(out))



def parse_well_name(well: str) -> Tuple[int, int]:
    m = _WELL_RE.match(well.strip())
    if not m:
        raise ValueError(f"Invalid well name: {well!r}")
    row = excel_col_to_number(m.group("row"))
    col = int(m.group("col"))
    return row, col



def sanitize_filename(text: str) -> str:
    cleaned = _SAFE_FILENAME_RE.sub("_", text.strip())
    cleaned = cleaned.strip("._")
    return cleaned or "plate"



def geometry_from_format(fmt: int | str) -> PlateGeometry:
    if fmt in (96, "96"):
        return GEOMETRY_96
    if fmt in (384, "384"):
        return GEOMETRY_384
    raise PicklyPyConfigError(f"Unsupported plate format: {fmt!r}. Use 96 or 384.")



def infer_plate_geometry(wells: Iterable[str], *, override: str | int = "auto") -> PlateGeometry:
    if override != "auto":
        return geometry_from_format(override)

    max_row = 0
    max_col = 0
    any_wells = False
    for well in wells:
        any_wells = True
        row, col = parse_well_name(well)
        max_row = max(max_row, row)
        max_col = max(max_col, col)

    if not any_wells:
        raise PicklyPyConfigError("Cannot infer plate geometry because no wells were found.")

    if max_row <= GEOMETRY_96.rows and max_col <= GEOMETRY_96.cols:
        return GEOMETRY_96
    if max_row <= GEOMETRY_384.rows and max_col <= GEOMETRY_384.cols:
        return GEOMETRY_384

    raise PicklyPyConfigError(
        "Only 96- and 384-well geometries are supported. "
        f"Observed wells extend to row {number_to_excel_col(max_row)} / column {max_col}."
    )



def well_to_xy(well: str, geometry: PlateGeometry) -> Tuple[float, float]:
    row, col = parse_well_name(well)
    if row > geometry.rows or col > geometry.cols:
        raise PicklyPyConfigError(
            f"Well {well!r} does not fit on a {geometry.wells}-well plate "
            f"({geometry.rows}x{geometry.cols})."
        )
    x = float(col)
    y = float(geometry.rows - row + 1)
    return x, y



def figsize_for_geometry(geometry: PlateGeometry) -> Tuple[float, float]:
    if geometry.wells == 96:
        return (9.2, 6.8)
    return (13.0, 9.0)



def text_size_for_geometry(geometry: PlateGeometry, style: RenderStyle) -> float:
    return style.base_fontsize_96 if geometry.wells == 96 else style.base_fontsize_384


# -----------------------------------------------------------------------------
# Picklist parsing
# -----------------------------------------------------------------------------


def parse_picklist_csv(csv_path: str | Path) -> ParsedPicklist:
    path = Path(csv_path)
    if not path.exists():
        raise PicklyPyConfigError(f"Picklist CSV not found: {path}")

    events: List[DispenseEvent] = []
    visits: List[SourceVisit] = []
    visit_count_by_src_plate: Dict[str, int] = defaultdict(int)

    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.reader(handle)
        global_index = 1
        for row_index, row in enumerate(reader, start=1):
            # Trim trailing blanks only. Blank lines are ignored.
            row = [as_str(cell).strip() for cell in row]
            while row and row[-1] == "":
                row.pop()
            if not row:
                continue

            if len(row) < 5:
                raise PicklyPyConfigError(
                    f"Malformed picklist row {row_index}: expected at least 5 columns, found {len(row)}."
                )
            if (len(row) - 3) % 2 != 0:
                raise PicklyPyConfigError(
                    f"Malformed picklist row {row_index}: after the first 3 columns, remaining values "
                    "must be volume/well pairs."
                )

            src_plate, src_well, dst_plate = row[:3]
            if not src_plate or not src_well or not dst_plate:
                raise PicklyPyConfigError(
                    f"Malformed picklist row {row_index}: source plate, source well, and destination plate "
                    "must all be present."
                )

            try:
                parse_well_name(src_well)
            except ValueError as exc:
                raise PicklyPyConfigError(f"Malformed source well in row {row_index}: {exc}") from exc

            row_events: List[DispenseEvent] = []
            for in_row_index, col in enumerate(range(3, len(row), 2), start=1):
                volume_txt = row[col]
                dst_well = row[col + 1]
                if not volume_txt or not dst_well:
                    raise PicklyPyConfigError(
                        f"Malformed picklist row {row_index}: volume/well pairs cannot be blank."
                    )
                try:
                    volume_nl = float(volume_txt)
                except ValueError as exc:
                    raise PicklyPyConfigError(
                        f"Malformed dispense volume in row {row_index}: {volume_txt!r}"
                    ) from exc
                try:
                    parse_well_name(dst_well)
                except ValueError as exc:
                    raise PicklyPyConfigError(f"Malformed destination well in row {row_index}: {exc}") from exc

                event = DispenseEvent(
                    global_index=global_index,
                    row_index=row_index,
                    in_row_index=in_row_index,
                    src_plate=src_plate,
                    src_well=src_well,
                    dst_plate=dst_plate,
                    dst_well=dst_well,
                    volume_nl=volume_nl,
                )
                row_events.append(event)
                events.append(event)
                global_index += 1

            visit_count_by_src_plate[src_plate] += 1
            visits.append(
                SourceVisit(
                    source_plate_visit_index=visit_count_by_src_plate[src_plate],
                    global_row_index=row_index,
                    src_plate=src_plate,
                    src_well=src_well,
                    dst_plate=dst_plate,
                    events=tuple(row_events),
                )
            )

    if not events:
        raise PicklyPyConfigError(f"No dispense events were found in {path}.")

    return ParsedPicklist(csv_path=path, events=tuple(events), visits=tuple(visits))


# -----------------------------------------------------------------------------
# Color assignment and index maps
# -----------------------------------------------------------------------------


def build_event_color_map(
    events: Sequence[DispenseEvent],
    *,
    cmap_name: str = "viridis",
    scope: str = "global",
) -> Tuple[Mapping[int, Tuple[float, float, float, float]], ScalarMappable]:
    cmap = matplotlib.colormaps.get_cmap(cmap_name)
    scope = scope.lower().strip()

    if scope not in {"global", "destination", "destination_plate", "per_destination_plate"}:
        raise PicklyPyConfigError(
            "color_scope must be 'global' or 'destination' (alias: per_destination_plate)."
        )

    if scope == "global":
        vmax = max(2, len(events))
        norm = mcolors.Normalize(vmin=1, vmax=vmax)
        color_map = {event.global_index: cmap(norm(event.global_index)) for event in events}
        return color_map, ScalarMappable(norm=norm, cmap=cmap)

    by_dst: Dict[str, List[DispenseEvent]] = defaultdict(list)
    for event in events:
        by_dst[event.dst_plate].append(event)

    color_map: Dict[int, Tuple[float, float, float, float]] = {}
    # This mappable is only approximate for colorbar display in per-destination mode.
    norm = mcolors.Normalize(vmin=1, vmax=max(2, max(len(v) for v in by_dst.values())))
    for plate_events in by_dst.values():
        local_norm = mcolors.Normalize(vmin=1, vmax=max(2, len(plate_events)))
        for local_index, event in enumerate(plate_events, start=1):
            color_map[event.global_index] = cmap(local_norm(local_index))
    return color_map, ScalarMappable(norm=norm, cmap=cmap)



def build_destination_event_numbers(events: Sequence[DispenseEvent]) -> Mapping[int, int]:
    by_dst: Dict[str, int] = defaultdict(int)
    numbers: Dict[int, int] = {}
    for event in events:
        by_dst[event.dst_plate] += 1
        numbers[event.global_index] = by_dst[event.dst_plate]
    return numbers



def build_source_visit_numbers(visits: Sequence[SourceVisit]) -> Mapping[Tuple[str, int], int]:
    # SourceVisit.source_plate_visit_index is already local to the plate. This helper makes lookup explicit.
    return {(visit.src_plate, visit.global_row_index): visit.source_plate_visit_index for visit in visits}


# -----------------------------------------------------------------------------
# Drawing helpers
# -----------------------------------------------------------------------------


def _average_rgba(colors: Sequence[Tuple[float, float, float, float]]) -> Tuple[float, float, float, float]:
    if not colors:
        return mcolors.to_rgba("#ffffff")
    n = float(len(colors))
    return (
        sum(c[0] for c in colors) / n,
        sum(c[1] for c in colors) / n,
        sum(c[2] for c in colors) / n,
        sum(c[3] for c in colors) / n,
    )



def _text_positions(x: float, y: float, count: int, radius: float) -> List[Tuple[float, float]]:
    if count <= 1:
        return [(x, y)]

    positions: List[Tuple[float, float]] = []
    if count <= 8:
        for i in range(count):
            angle = math.pi / 2 - (2 * math.pi * i / count)
            positions.append((x + radius * math.cos(angle), y + radius * math.sin(angle)))
        return positions

    # Two rings for denser revisits.
    inner_count = math.ceil(count / 2)
    outer_count = count - inner_count
    positions.extend(_text_positions(x, y, inner_count, radius * 0.72))
    positions.extend(_text_positions(x, y, outer_count, radius * 1.25))
    return positions



def _draw_plate_background(ax, geometry: PlateGeometry, style: RenderStyle) -> None:
    x0 = 0.35
    y0 = 0.35
    width = geometry.cols + 1.05
    height = geometry.rows + 0.95
    plate = FancyBboxPatch(
        (x0, y0),
        width,
        height,
        boxstyle="round,pad=0.03,rounding_size=0.35",
        linewidth=2.2,
        edgecolor=style.plate_edge,
        facecolor=style.plate_face,
        zorder=0,
    )
    ax.add_patch(plate)

    # Column numbers.
    for col in range(1, geometry.cols + 1):
        ax.text(
            col,
            geometry.rows + 0.65,
            str(col),
            ha="center",
            va="center",
            color="#c7c7c7",
            fontsize=8 if geometry.wells == 96 else 6,
            fontweight="bold",
            zorder=10,
        )

    # Row letters.
    for row in range(1, geometry.rows + 1):
        ax.text(
            0.55,
            geometry.rows - row + 1,
            number_to_excel_col(row),
            ha="center",
            va="center",
            color="#c7c7c7",
            fontsize=8 if geometry.wells == 96 else 6,
            fontweight="bold",
            zorder=10,
        )

    # Base empty wells.
    for row in range(1, geometry.rows + 1):
        for col in range(1, geometry.cols + 1):
            cx = float(col)
            cy = float(geometry.rows - row + 1)
            ax.add_patch(
                Circle(
                    (cx, cy),
                    radius=style.well_radius,
                    facecolor=style.empty_well_face,
                    edgecolor=style.empty_well_edge,
                    linewidth=1.0,
                    zorder=0.5,
                )
            )



def _draw_path(ax, points: Sequence[Tuple[float, float]], style: RenderStyle) -> None:
    if len(points) < 2:
        return
    xs = [p[0] for p in points]
    ys = [p[1] for p in points]
    ax.plot(
        xs,
        ys,
        color=style.path_color,
        linewidth=style.path_width,
        alpha=0.95,
        solid_capstyle="round",
        solid_joinstyle="round",
        zorder=1.2,
    )
    ax.scatter(xs, ys, s=style.visit_marker_size, c=style.path_color, zorder=1.3)



def _draw_event_well(
    ax,
    *,
    x: float,
    y: float,
    colors: Sequence[Tuple[float, float, float, float]],
    numbers: Sequence[int],
    geometry: PlateGeometry,
    style: RenderStyle,
) -> None:
    if not colors:
        return

    if len(colors) == 1:
        ax.add_patch(
            Circle(
                (x, y),
                radius=style.well_radius * 0.96,
                facecolor=colors[0],
                edgecolor="#0f0f0f",
                linewidth=1.4,
                zorder=2.2,
            )
        )
    else:
        # Neutral center for readability, segmented outer ring for chronological colors.
        ax.add_patch(
            Circle(
                (x, y),
                radius=style.well_radius * 0.96,
                facecolor=_average_rgba(colors),
                edgecolor="#0f0f0f",
                linewidth=1.2,
                alpha=0.35,
                zorder=2.0,
            )
        )
        n = len(colors)
        for i, color in enumerate(colors):
            theta1 = 90.0 - (360.0 * i / n)
            theta2 = 90.0 - (360.0 * (i + 1) / n)
            ax.add_patch(
                Wedge(
                    center=(x, y),
                    r=style.well_radius * 0.98,
                    theta1=theta2,
                    theta2=theta1,
                    width=(style.well_radius - style.multi_inner_radius) * 0.96,
                    facecolor=color,
                    edgecolor="#0f0f0f",
                    linewidth=0.4,
                    zorder=2.25,
                )
            )
        ax.add_patch(
            Circle(
                (x, y),
                radius=style.multi_inner_radius,
                facecolor=style.neutral_multi_fill,
                edgecolor="#0f0f0f",
                linewidth=0.8,
                zorder=2.4,
            )
        )

    font_size = text_size_for_geometry(geometry, style)
    positions = _text_positions(x, y, len(numbers), style.well_radius * 0.33)
    for (tx, ty), number in zip(positions, numbers):
        text = ax.text(
            tx,
            ty,
            str(number),
            ha="center",
            va="center",
            color="black",
            fontsize=max(4.0, font_size - 0.45 * max(0, len(numbers) - 1)),
            fontweight="bold",
            zorder=3.5,
        )
        text.set_path_effects([pe.withStroke(linewidth=2.0, foreground="white")])



def _finish_plate_figure(
    fig,
    ax,
    *,
    geometry: PlateGeometry,
    title: str,
    subtitle: str,
    color_mappable: ScalarMappable,
    color_scope: str,
    output_path: Path,
    dpi: int,
) -> None:
    ax.set_xlim(0.15, geometry.cols + 1.05)
    ax.set_ylim(0.15, geometry.rows + 1.05)
    ax.set_aspect("equal")
    ax.axis("off")
    ax.set_title(title, fontsize=15 if geometry.wells == 96 else 13, fontweight="bold", pad=14)
    if subtitle:
        fig.text(0.5, 0.02, subtitle, ha="center", va="bottom", fontsize=9)

    cbar = fig.colorbar(color_mappable, ax=ax, fraction=0.030, pad=0.02)
    if color_scope == "global":
        cbar.set_label("Global dispense order", fontsize=9)
    else:
        cbar.set_label("Dispense order within destination plate", fontsize=9)
    cbar.ax.tick_params(labelsize=8)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(output_path, dpi=dpi, bbox_inches="tight")
    plt.close(fig)


# -----------------------------------------------------------------------------
# Rendering
# -----------------------------------------------------------------------------


def render_destination_plate(
    *,
    plate: str,
    events: Sequence[DispenseEvent],
    geometry: PlateGeometry,
    event_colors: Mapping[int, Tuple[float, float, float, float]],
    event_numbers: Mapping[int, int],
    color_mappable: ScalarMappable,
    color_scope: str,
    output_path: Path,
    style: RenderStyle,
    dpi: int,
) -> None:
    fig, ax = plt.subplots(figsize=figsize_for_geometry(geometry))
    _draw_plate_background(ax, geometry, style)

    path_points: List[Tuple[float, float]] = [well_to_xy(event.dst_well, geometry) for event in events]
    _draw_path(ax, path_points, style)

    colors_by_well: Dict[str, List[Tuple[float, float, float, float]]] = defaultdict(list)
    numbers_by_well: Dict[str, List[int]] = defaultdict(list)
    for event in events:
        colors_by_well[event.dst_well].append(event_colors[event.global_index])
        numbers_by_well[event.dst_well].append(event_numbers[event.global_index])

    for well in colors_by_well:
        x, y = well_to_xy(well, geometry)
        _draw_event_well(
            ax,
            x=x,
            y=y,
            colors=colors_by_well[well],
            numbers=numbers_by_well[well],
            geometry=geometry,
            style=style,
        )

    subtitle = (
        f"{len(events)} dispense events from {len(set(e.src_plate for e in events))} source plate(s)"
    )
    _finish_plate_figure(
        fig,
        ax,
        geometry=geometry,
        title=f"Destination plate {plate}",
        subtitle=subtitle,
        color_mappable=color_mappable,
        color_scope=color_scope,
        output_path=output_path,
        dpi=dpi,
    )



def render_source_plate(
    *,
    plate: str,
    visits: Sequence[SourceVisit],
    geometry: PlateGeometry,
    event_colors: Mapping[int, Tuple[float, float, float, float]],
    visit_numbers: Mapping[Tuple[str, int], int],
    color_mappable: ScalarMappable,
    color_scope: str,
    output_path: Path,
    style: RenderStyle,
    dpi: int,
) -> None:
    fig, ax = plt.subplots(figsize=figsize_for_geometry(geometry))
    _draw_plate_background(ax, geometry, style)

    visit_points: List[Tuple[float, float]] = [well_to_xy(visit.src_well, geometry) for visit in visits]
    _draw_path(ax, visit_points, style)

    colors_by_well: Dict[str, List[Tuple[float, float, float, float]]] = defaultdict(list)
    numbers_by_well: Dict[str, List[int]] = defaultdict(list)

    for visit in visits:
        numbers_by_well[visit.src_well].append(visit_numbers[(visit.src_plate, visit.global_row_index)])
        for event in visit.events:
            colors_by_well[visit.src_well].append(event_colors[event.global_index])

    for well in set(list(colors_by_well.keys()) + list(numbers_by_well.keys())):
        x, y = well_to_xy(well, geometry)
        _draw_event_well(
            ax,
            x=x,
            y=y,
            colors=colors_by_well.get(well, []),
            numbers=numbers_by_well.get(well, []),
            geometry=geometry,
            style=style,
        )

    total_events = sum(len(visit.events) for visit in visits)
    subtitle = f"{len(visits)} source-well visits, {total_events} dispense events"
    _finish_plate_figure(
        fig,
        ax,
        geometry=geometry,
        title=f"Source plate {plate}",
        subtitle=subtitle,
        color_mappable=color_mappable,
        color_scope=color_scope,
        output_path=output_path,
        dpi=dpi,
    )


# -----------------------------------------------------------------------------
# Public API
# -----------------------------------------------------------------------------


def visualize_picklist(
    picklist_csv: str | Path,
    *,
    output_dir: str | Path | None = None,
    src_format: str | int = "auto",
    dst_format: str | int = "auto",
    image_format: str = "png",
    dpi: int = 300,
    cmap_name: str = "viridis",
    color_scope: str = "global",
    style: Optional[RenderStyle] = None,
) -> VisualizationResult:
    """Create source-plate and destination-plate dispense-path images from a PicklyPy picklist CSV.

    Parameters
    ----------
    picklist_csv:
        Path to the picklist CSV produced by picklipy.assay or picklipy.screen.
    output_dir:
        Folder for rendered images. Defaults to <picklist_stem>_movement_viz next to the CSV.
    src_format / dst_format:
        'auto', 96, or 384. Auto selects the smallest compatible format from the observed wells.
        When only top-left wells are used, a physically 384-well source plate may still need an
        explicit override such as src_format=384.
    image_format:
        File extension/format accepted by matplotlib, typically png, svg, or pdf.
    dpi:
        Output resolution.
    cmap_name:
        Matplotlib colormap name used to encode dispense order.
    color_scope:
        'global' uses one gradient across the full picklist so source and destination images share
        the same colors. 'destination' restarts the gradient per destination plate.
    style:
        Optional RenderStyle override.
    """
    parsed = parse_picklist_csv(picklist_csv)
    style = style or RenderStyle(cmap_name=cmap_name)
    image_format = image_format.lstrip(".").lower()

    out_dir = Path(output_dir) if output_dir is not None else parsed.csv_path.with_name(
        f"{parsed.csv_path.stem}_movement_viz"
    )
    out_dir.mkdir(parents=True, exist_ok=True)

    event_colors, color_mappable = build_event_color_map(
        parsed.events,
        cmap_name=cmap_name,
        scope=color_scope,
    )
    dst_event_numbers = build_destination_event_numbers(parsed.events)
    src_visit_numbers = build_source_visit_numbers(parsed.visits)

    destination_images: Dict[str, Path] = {}
    source_images: Dict[str, Path] = {}

    events_by_dst: Dict[str, List[DispenseEvent]] = defaultdict(list)
    for event in parsed.events:
        events_by_dst[event.dst_plate].append(event)

    visits_by_src: Dict[str, List[SourceVisit]] = defaultdict(list)
    for visit in parsed.visits:
        visits_by_src[visit.src_plate].append(visit)

    for dst_plate, plate_events in events_by_dst.items():
        geometry = infer_plate_geometry((event.dst_well for event in plate_events), override=dst_format)
        output_path = out_dir / f"{parsed.csv_path.stem}__destination__{sanitize_filename(dst_plate)}.{image_format}"
        render_destination_plate(
            plate=dst_plate,
            events=plate_events,
            geometry=geometry,
            event_colors=event_colors,
            event_numbers=dst_event_numbers,
            color_mappable=color_mappable,
            color_scope=color_scope,
            output_path=output_path,
            style=style,
            dpi=dpi,
        )
        destination_images[dst_plate] = output_path

    for src_plate, plate_visits in visits_by_src.items():
        geometry = infer_plate_geometry((visit.src_well for visit in plate_visits), override=src_format)
        output_path = out_dir / f"{parsed.csv_path.stem}__source__{sanitize_filename(src_plate)}.{image_format}"
        render_source_plate(
            plate=src_plate,
            visits=plate_visits,
            geometry=geometry,
            event_colors=event_colors,
            visit_numbers=src_visit_numbers,
            color_mappable=color_mappable,
            color_scope=color_scope,
            output_path=output_path,
            style=style,
            dpi=dpi,
        )
        source_images[src_plate] = output_path

    return VisualizationResult(
        output_dir=out_dir,
        destination_images=destination_images,
        source_images=source_images,
        event_count=len(parsed.events),
    )


# -----------------------------------------------------------------------------
# CLI
# -----------------------------------------------------------------------------


def _print_banner() -> None:
    print(
        "PicklyPy dispense-path visualizer. "
        "Creates one path image per destination plate and one per source plate."
    )
    print(
        "Destination wells are numbered in dispense order per destination plate; "
        "source wells are numbered in visit order per source plate."
    )
    print(
        "Colors encode dispense order and are shared between source and destination views "
        "when color_scope='global'."
    )



def _run_cli(args: argparse.Namespace) -> None:
    _print_banner()
    result = visualize_picklist(
        args.picklist_csv,
        output_dir=args.output_dir,
        src_format=args.src_format,
        dst_format=args.dst_format,
        image_format=args.image_format,
        dpi=args.dpi,
        cmap_name=args.cmap,
        color_scope=args.color_scope,
    )

    print(f"Loaded {result.event_count} dispense events.")
    print(f"Saved images to: {result.output_dir}")
    print("Destination plate images:")
    for plate, path in result.destination_images.items():
        print(f"  {plate}: {path.name}")
    print("Source plate images:")
    for plate, path in result.source_images.items():
        print(f"  {plate}: {path.name}")



def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        prog="picklipy.visualize",
        description=(
            "Render dispense-order path images from a PicklyPy picklist CSV. "
            "The output includes one image per destination plate and one image per source plate."
        ),
    )
    parser.add_argument("picklist_csv", help="Path to the picklist CSV created by picklipy.assay or picklipy.screen")
    parser.add_argument(
        "--output-dir",
        default=None,
        help="Folder for output images. Defaults to <picklist_stem>_movement_viz next to the CSV.",
    )
    parser.add_argument(
        "--src-format",
        default="auto",
        choices=["auto", "96", "384"],
        help="Source plate format. Use 384 when only top-left wells are used but the source plate is physically 384-well.",
    )
    parser.add_argument(
        "--dst-format",
        default="auto",
        choices=["auto", "96", "384"],
        help="Destination plate format.",
    )
    parser.add_argument(
        "--image-format",
        default="png",
        help="Matplotlib output format, usually png, svg, or pdf.",
    )
    parser.add_argument("--dpi", type=int, default=300, help="Output resolution in dots per inch.")
    parser.add_argument("--cmap", default="viridis", help="Matplotlib colormap name for dispense order.")
    parser.add_argument(
        "--color-scope",
        default="global",
        choices=["global", "destination"],
        help="Use one color gradient across the full picklist ('global') or restart the gradient per destination plate ('destination').",
    )
    parser.add_argument("--no-pause", action="store_true", help="Do not wait for Enter before exiting.")

    args = parser.parse_args(list(argv) if argv is not None else None)

    # Convert numeric options to ints where applicable while keeping the 'auto' sentinel.
    args.src_format = args.src_format if args.src_format == "auto" else int(args.src_format)
    args.dst_format = args.dst_format if args.dst_format == "auto" else int(args.dst_format)

    return run_with_user_facing_errors(lambda: _run_cli(args), pause_on_exit=not args.no_pause)


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
