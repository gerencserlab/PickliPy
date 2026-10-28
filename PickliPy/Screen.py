from __future__ import annotations

import copy
import math
import re
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Any, Dict, List, Mapping, Optional, Sequence, Tuple

import openpyxl

import warnings
warnings.filterwarnings(
    "ignore",
    message="Data Validation extension is not supported",
)

from .common import (
    PLATE_COLS_384,
    PLATE_COLS_1536,
    PLATE_ROWS_384,
    PLATE_ROWS_1536,
    PicklyPyConfigError,
    PicklyPyError,
    PicklyPyVolumeError,
    as_number,
    as_str,
    excel_col_to_number,
    format_csv_rows,
    integer_chop,
    load_survey_volumes,
    make_inventory_row,
    parse_well_name,
    quote_csv_field,
    reorder_destinations_within_group,
    reorder_source_groups,
    split_labels,
    tally_ordered,
    write_text,
    well_name
)
from .randomization import optimize_plate_well_map, seed_from_text
from .backfill import (
    backfill_remainders,
    normalize_backfill_override,
    reserved_backfill_row_indices,
    target_volume_from_mock_concentration,
    validate_single_workbook_backfill,
)


@dataclass
class _ScreenHeader:
    inventory_src_name: str
    inventory_dst_name: str
    inventory_src_rack: int
    inventory_dst_rack: int
    process_src_name: str
    process_dst_name: str
    picklist_name: str
    barcodes_dst: List[str]
    well_volume_ul: float
    min_volume_ul: float


@dataclass(frozen=True)
class _LibReferenceColumns:
    well: Tuple[int, ...]
    stock: Tuple[int, ...]
    barcode: Tuple[int, ...]


@dataclass(frozen=True)
class _ReformattedLibDispense:
    lib_row_index: int
    destination_barcode: str
    destination_well: str
    volume_nl: float
    stock_concentration: float
    source_barcode: str
    source_well: str


def _at_reference_columns(value: Any, lib_column_count: int) -> Tuple[int, ...]:
    """Return zero-based LIB columns referenced by an SRC cell."""

    columns: List[int] = []
    for column_letters in re.findall(r"@([A-Za-z]+)", as_str(value)):
        column_index = excel_col_to_number(column_letters) - 1
        if 0 <= column_index < lib_column_count and column_index not in columns:
            columns.append(column_index)
    return tuple(columns)


def _lib_references_for_src_row(
    src_row: Sequence[Any],
    lib_column_count: int,
) -> _LibReferenceColumns:
    def cell(column_index: int) -> Any:
        return src_row[column_index] if column_index < len(src_row) else ""

    return _LibReferenceColumns(
        well=_at_reference_columns(cell(0), lib_column_count),
        stock=_at_reference_columns(cell(2), lib_column_count),
        barcode=_at_reference_columns(cell(5), lib_column_count),
    )


def _common_lib_references(
    src_rows: Sequence[Sequence[Any]],
    lib_column_count: int,
    *,
    using_backfill: bool,
) -> _LibReferenceColumns:
    """Resolve one unambiguous LIB column mapping for the export schema."""

    # addcompound consumes a LIB row whenever any SRC cell contains an @
    # reference. Fully fixed explicit controls are irrelevant to the export
    # schema; every LIB-consuming row, including a partial mapping, must agree.
    references = [
        _lib_references_for_src_row(row, lib_column_count)
        for row in src_rows
        if any("@" in as_str(cell) for cell in row)
    ]

    def common_columns(role: str, display_name: str) -> Tuple[int, ...]:
        mappings: List[Tuple[int, ...]] = []
        for reference in references:
            columns = getattr(reference, role)
            if columns not in mappings:
                mappings.append(columns)
        if len(mappings) > 1:
            human_mappings = [
                tuple(column + 1 for column in mapping)
                for mapping in mappings
            ]
            raise PicklyPyConfigError(
                "Reformatted LIB export requires one consistent SRC @ mapping "
                f"for {display_name}; found LIB column sets {human_mappings}."
            )
        return mappings[0] if mappings else ()

    return _LibReferenceColumns(
        well=common_columns("well", "Source Well"),
        stock=(
            common_columns("stock", "Stock concentration")
            if using_backfill
            else ()
        ),
        barcode=common_columns("barcode", "Plate barcode"),
    )


def _write_reformatted_lib_workbook(
    output_path: Path,
    *,
    lib: Sequence[Sequence[Any]],
    dispenses: Sequence[_ReformattedLibDispense],
    total_volume_nl_by_destination: Mapping[Tuple[str, str], float],
    destination_order: Sequence[str],
    using_backfill: bool,
    references: _LibReferenceColumns,
) -> None:
    """Write the daughter-library view of actual Screen dispense events."""

    lib_rows = [list(row) for row in lib]
    if not lib_rows:
        lib_rows = [[]]
    lib_column_count = max((len(row) for row in lib_rows), default=0)
    original_headers = list(lib_rows[0]) + [""] * (
        lib_column_count - len(lib_rows[0])
    )

    all_well_columns = list(references.well)
    all_barcode_columns = list(references.barcode)
    all_stock_columns = list(references.stock)

    header_annotations: Dict[int, List[str]] = {}
    for column in all_well_columns:
        header_annotations.setdefault(column, []).append("Dispense destination well")
    for column in all_barcode_columns:
        header_annotations.setdefault(column, []).append(
            "Dispense destination barcode"
        )
    if using_backfill:
        for column in all_stock_columns:
            header_annotations.setdefault(column, []).append(
                "Concentration after dispense"
            )

    output_headers = list(original_headers)
    for column, annotations in header_annotations.items():
        original_header = as_str(original_headers[column]).strip()
        if original_header == "":
            original_header = f"LIB column {column + 1}"
        output_headers[column] = f"{original_header} ({'; '.join(annotations)})"

    # A Screen design normally references LIB columns for source well and
    # source barcode. Preserve useful destination coordinates even for a
    # non-standard template that consumes LIB rows through another @ field.
    add_destination_well_column = not all_well_columns
    add_destination_barcode_column = not all_barcode_columns
    if add_destination_well_column:
        output_headers.append("Destination well (Dispense destination well)")
    if add_destination_barcode_column:
        output_headers.append(
            "Destination plate barcode (Dispense destination barcode)"
        )

    output_headers.extend(
        [
            "Dispense volume (nL)",
            "Original source well",
            "Original source plate barcode",
        ]
    )

    destination_rank = {
        barcode: index for index, barcode in enumerate(destination_order)
    }
    indexed_dispenses = list(enumerate(dispenses))

    def dispense_sort_key(
        item: Tuple[int, _ReformattedLibDispense],
    ) -> Tuple[int, int, int, int]:
        sequence, event = item
        try:
            row, column = parse_well_name(event.destination_well)
        except ValueError:
            row, column = (10**9, 10**9)
        return (
            destination_rank.get(event.destination_barcode, 10**9),
            row,
            column,
            sequence,
        )

    output_rows: List[List[Any]] = []
    for _sequence, event in sorted(indexed_dispenses, key=dispense_sort_key):
        if event.lib_row_index <= 0 or event.lib_row_index >= len(lib_rows):
            raise PicklyPyConfigError(
                "Cannot create reformatted LIB row: resolved LIB row index "
                f"{event.lib_row_index} is outside the imported LIB worksheet."
            )

        original_row = list(lib_rows[event.lib_row_index]) + [""] * (
            lib_column_count - len(lib_rows[event.lib_row_index])
        )
        output_row = list(original_row)

        for column in references.well:
            output_row[column] = event.destination_well
        for column in references.barcode:
            output_row[column] = event.destination_barcode

        if using_backfill and references.stock:
            total_volume_nl = total_volume_nl_by_destination.get(
                (event.destination_barcode, event.destination_well),
                0.0,
            )
            resulting_concentration = 0.0
            if total_volume_nl > 0:
                resulting_concentration = (
                    event.stock_concentration
                    * event.volume_nl
                    / total_volume_nl
                )
            for column in references.stock:
                output_row[column] = integer_chop(resulting_concentration)

        if add_destination_well_column:
            output_row.append(event.destination_well)
        if add_destination_barcode_column:
            output_row.append(event.destination_barcode)

        output_row.extend(
            [
                integer_chop(event.volume_nl),
                event.source_well,
                event.source_barcode,
            ]
        )
        output_rows.append(output_row)

    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "LIB"
    worksheet.append(output_headers)
    for row in output_rows:
        worksheet.append(row)
    workbook.save(output_path)


def _print_banner() -> None:
    print(
        "Picklist generator for acoustic dispensing using Beckman Coulter Echo 650 "
        "controlled by Revvity PlateWorks. Written by Akos A. Gerencser."
    )
    
    print(
        "\nPicklyPy.Screen (Python translation of screeningpicklist.wls v6.6 10/25/2025)\n"
        "Generates screening picklists for a generic destination layout, auto-filled\n"
        "from a multi-plate compound library (LIB worksheet).\n"
    )
    print(
        "Version 6.7 10/28/2026"
    )

def _matrix_nonempty_first_col(matrix: List[List[Any]]) -> List[List[Any]]:
    out: List[List[Any]] = []
    for row in matrix:
        if len(row) == 0:
            continue
        if as_str(row[0]) == "":
            continue
        out.append(row)
    return out


def _detect_plate_format(dst: List[List[Any]], plate_map_row: int) -> Tuple[int, int]:
    """Auto-detect plate format from the column headers in the Plate Map row.

    Counts numeric column headers in the row containing 'Plate Map:'.
    Returns (plate_rows, plate_cols).
    """
    header_row = dst[plate_map_row]
    col_count = 0
    for c in range(1, len(header_row)):
        v = as_number(header_row[c])
        if v is not None and int(v) == col_count + 1:
            col_count += 1
        else:
            break
    if col_count == PLATE_COLS_1536:
        return PLATE_ROWS_1536, PLATE_COLS_1536
    return PLATE_ROWS_384, PLATE_COLS_384

def _load_design_sheets(xlsx_path: Path) -> Tuple[List[List[Any]], List[List[Any]], List[List[Any]], Optional[List[List[Any]]]]:
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    # Prefer conventional sheet names if available, else fall back to positional order.
    def _get_by_name(names: Sequence[str]) -> Optional[openpyxl.worksheet.worksheet.Worksheet]:
        for nm in names:
            if nm in wb.sheetnames:
                return wb[nm]
        return None

    ws_src = _get_by_name(["SRC", "Src", "src"]) or wb.worksheets[0]
    ws_dst = _get_by_name(["DST", "Dst", "dst"]) or (wb.worksheets[1] if len(wb.worksheets) > 1 else None)
    ws_lib = _get_by_name(["LIB", "Lib", "lib"]) or (wb.worksheets[2] if len(wb.worksheets) > 2 else None)
    ws_blk = None
    for name in wb.sheetnames:
        name_lower = name.lower()
        if "blacklist" in name_lower or "blk" in name_lower:
            ws_blk = wb[name]
            break
    if (ws_blk is None and len(wb.worksheets) > 3
            and not wb.worksheets[3].title.lower().startswith('intermediate')):
        ws_blk = wb.worksheets[3]

    if ws_dst is None or ws_lib is None:
        raise PicklyPyConfigError(
            "The design workbook must contain SRC, DST and LIB worksheets (or at least 3 worksheets)."
        )

    from .common import sheet_to_matrix

    src = sheet_to_matrix(ws_src, min_cols=8)
    dst = sheet_to_matrix(ws_dst, min_cols=25)
    lib = sheet_to_matrix(ws_lib, min_cols=1)
    blk = sheet_to_matrix(ws_blk, min_cols=25) if ws_blk is not None else None

    # Match Wolfram behavior: drop rows where first column is empty in SRC.
    src = _matrix_nonempty_first_col(src)

    # Match Wolfram behavior: LIB keeps the header and drops rows with empty first column.
    if len(lib) >= 1:
        header = lib[0]
        data = [r for r in lib[1:] if as_str(r[0]) != ""]
        lib = [header] + data

    return src, dst, lib, blk


def _parse_header(dst: List[List[Any]], src: List[List[Any]]) -> _ScreenHeader:
    # Wolfram: header = Map[#[[1]]->#[[2]]&, dst[[1;;12,{1,2}]]]
    header_map: Dict[str, Any] = {}
    for r in range(len(dst)):
        key = as_str(dst[r][0]) if len(dst[r]) > 0 else ""
        val = dst[r][1] if len(dst[r]) > 1 else ""
        if key != "":
            header_map[key] = val

    def _get_str(k: str, default: str) -> str:
        v = header_map.get(k, default)
        s = as_str(v)
        return s if s != "" else default

    def _get_int(k: str, default: int) -> int:
        v = header_map.get(k, default)
        n = as_number(v)
        if n is None:
            return default
        return int(integer_chop(n))

    def _get_float(k: str, default: float) -> float:
        v = header_map.get(k, default)
        n = as_number(v)
        if n is None:
            return default
        return float(n)

    inventory_src = _get_str("Inventory_SRC:", "Inventory")
    inventory_dst = _get_str("Inventory_DST:", "Inventory")
    rack_src = _get_int("Inventory_SRC_Rack#:", 1)
    rack_dst = _get_int("Inventory_DST_Rack#:", 2)
    process_src = _get_str("Process_SRC:", "Process")
    process_dst = _get_str("Process_DST:", "Process")

    picklist_name = _get_str("Picklist Name:", "picklist")
    barcodes_dst_raw = _get_str("Barcode_DST:", "PlateDST")
    barcodes_dst = [b.strip() for b in barcodes_dst_raw.split(",") if b.strip()]

    wellvol = _get_float("Well volume (ul):", _get_float("Well volume (uL):", 50.0))

    # min volume in SRC cell G1 (row 1 col 7; 1-based). src in Wolfram may have been read with first row intact.
    minvol = 0.0
    if len(src) >= 1:
        # Ensure at least 7 cols.
        row0 = src[0] + [""] * max(0, 7 - len(src[0]))
        minvol = float(as_number(row0[6]) or 0.0)
    if minvol <= 0:
        raise PicklyPyConfigError(
            "SRC sheet cell G1 must contain a positive minimum source volume (ul)."
        )

    if not barcodes_dst:
        raise PicklyPyConfigError(
            "DST header field 'Barcode_DST:' must contain at least one destination barcode."
        )

    return _ScreenHeader(
        inventory_src_name=inventory_src,
        inventory_dst_name=inventory_dst,
        inventory_src_rack=rack_src,
        inventory_dst_rack=rack_dst,
        process_src_name=process_src,
        process_dst_name=process_dst,
        picklist_name=picklist_name,
        barcodes_dst=barcodes_dst,
        well_volume_ul=float(wellvol),
        min_volume_ul=float(minvol),
    )


def _find_row(dst: List[List[Any]], key: str) -> Optional[int]:
    for i, row in enumerate(dst):
        if as_str(row[0]) == key:
            return i
    return None


def _extract_plate_grid(dst: List[List[Any]], top_row: int, left_col: int = 1,
                        n_rows: int = PLATE_ROWS_384, n_cols: int = PLATE_COLS_384) -> List[List[str]]:
    grid: List[List[str]] = []
    for r in range(top_row, top_row + n_rows):
        if r >= len(dst):
            raise PicklyPyConfigError(
                f"DST worksheet ended while reading a {n_rows}x{n_cols} plate map at row {top_row+1}."
            )
        row_vals = dst[r]
        row_out: List[str] = []
        for c in range(left_col, left_col + n_cols):
            v = row_vals[c] if c < len(row_vals) else ""
            # Screening script coerces numeric to string.
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                v = integer_chop(float(v))
            row_out.append(as_str(v))
        grid.append(row_out)
    return grid


def _plate_map_pairs_from_grid(grid: List[List[str]]) -> List[Tuple[str, str]]:
    pairs: List[Tuple[str, str]] = []
    for r in range(len(grid)):
        for c in range(len(grid[0])):
            cell = grid[r][c]
            if cell.strip() == "":
                continue
            dstwell = well_name(r+1,c+1)
            pairs.append((dstwell, cell))
    return pairs


@dataclass
class _Blacklist:
    groups_by_label: Dict[str, List[str]]
    label_by_well: Dict[str, str]
    blacklisted_wells_by_dst: Dict[str, List[str]]


def _parse_blacklist(blk: Optional[List[List[Any]]], plate_rows: int = PLATE_ROWS_384, plate_cols: int = PLATE_COLS_384) -> Optional[_Blacklist]:
    if blk is None:
        return None

    if plate_rows > PLATE_ROWS_384:
        print("Note: Blacklist is not yet supported for 1536-well plate format. Skipping blacklist.")
        return None

    # Convert entire sheet to string matrix for key scanning.
    blk_str: List[List[str]] = [[as_str(v) for v in row] for row in blk]

    # Find "Groups Map:" and read the next 16 rows x 24 cols from B..Y.
    groups_row = None
    for i, row in enumerate(blk_str):
        if row and row[0] == "Groups:":
            groups_row = i
            break

    if groups_row is None:
        return None

    groups_grid = _extract_plate_grid(blk_str, groups_row + 1, left_col=1)
    

    groups_by_label: Dict[str, List[str]] = {}
    label_by_well: Dict[str, str] = {}
    for r in range(PLATE_ROWS_384):
        for c in range(PLATE_COLS_384):
            g = groups_grid[r][c].strip()
            if g == "":
                continue
            well = well_name(r+1,c+1)
            groups_by_label.setdefault(g, []).append(well)
            label_by_well[well] = g

    # Find "Well Blacklist:" table: two columns A..B until blank in A.
    blk_row = None
    for i, row in enumerate(blk_str):
        for j, cell in enumerate(row):
            cell_clean = cell.strip().lower().replace("_", " ").rstrip(":")
            if cell_clean in ("well blacklist", "well_blacklist"):
                blk_row = i
                break
        if blk_row is not None:
            break

    blacklisted_by_dst: Dict[str, List[str]] = {}
    if blk_row is not None:
        # Find which columns hold barcode and wells by checking the header row
        barcode_col = None
        wells_col = None
        for j, cell in enumerate(blk_str[blk_row]):
            cell_clean = cell.strip().lower().replace("_", " ").rstrip(":")
            if "barcode" in cell_clean:
                barcode_col = j
            if "blacklist" in cell_clean or "well" in cell_clean:
                wells_col = j
        if barcode_col is None:
            barcode_col = 0
        if wells_col is None:
            wells_col = 1

        j = blk_row + 1
        while j < len(blk_str):
            barcode = blk_str[j][barcode_col].strip() if len(blk_str[j]) > barcode_col else ""
            if barcode == "":
                break
            wells_raw = blk_str[j][wells_col].strip() if len(blk_str[j]) > wells_col else ""
            wells = [w.strip() for w in wells_raw.split(",") if w.strip()]
            blacklisted_by_dst[barcode] = wells
            j += 1

    return _Blacklist(groups_by_label=groups_by_label, label_by_well=label_by_well, blacklisted_wells_by_dst=blacklisted_by_dst)


def _round_dispense_nl(desired_nl: float) -> float:
    # Mathematica Round[..., 2.5] (half-away-from-zero) is implemented in common for Assay;
    # For Screen we can reuse same logic, but we keep it here to match script.
    from .common import round_to_increment

    return round_to_increment(desired_nl, 2.5)


def _allocate_randomized_blacklist_plate(
    *,
    plate_pairs: Sequence[Tuple[str, str]],
    valid_wells: Sequence[str],
    blacklisted_wells: set[str],
    layout_by_well: Mapping[str, str],
    labels_data: Sequence[Sequence[Any]],
    explicit_compounds: set[str],
    barcode: str,
) -> List[Tuple[str, str]]:
    """Reserve controls, then fit complete library sets within physical Layout pools.

    Group-map text does not define pools here. As with randomization itself,
    absent/blank Layout labels form one pool. Library cells sharing a compound
    slot are indivisible, including dose aliases and connected combinations.
    LIB is resolved only after allocation, so omitted slots consume no LIB rows.
    """
    compounds_by_label: Dict[str, set[str]] = {}
    for row in labels_data:
        compounds_by_label.setdefault(as_str(row[0]), set()).add(as_str(row[1]))

    controls: List[Tuple[str, str]] = []
    library: List[Tuple[str, str]] = []
    library_compounds: List[set[str]] = []
    includes_control: List[bool] = []
    for well, cell in plate_pairs:
        compounds: set[str] = set()
        for label in split_labels(cell):
            if label not in compounds_by_label:
                raise PicklyPyConfigError(
                    f"Plate-map label {label!r} is not defined in the Labels table."
                )
            compounds.update(compounds_by_label[label])
        if compounds and compounds <= explicit_compounds:
            controls.append((well, cell))
        else:
            library.append((well, cell))
            library_compounds.append(compounds - explicit_compounds)
            includes_control.append(bool(compounds & explicit_compounds))

    control_wells = {well for well, _ in controls}
    pools: Dict[str, List[str]] = {}
    for well in dict.fromkeys(valid_wells):
        if well not in blacklisted_wells and well not in control_wells:
            pools.setdefault(layout_by_well.get(well, ""), []).append(well)

    result: List[Tuple[str, str]] = []
    for well, cell in controls:
        if well not in blacklisted_wells:
            result.append((well, cell))
            continue
        pool = pools.get(layout_by_well.get(well, ""), [])
        if not pool:
            raise PicklyPyConfigError(
                f"Cannot preserve control {cell!r} from {well} on destination plate "
                f"{barcode}: no healthy replacement in its Layout group "
                f"{layout_by_well.get(well, '')!r}. A uniquely pinned control cannot move."
            )
        replacement = pool.pop(0)
        result.append((replacement, cell))
        print(f"Replacing blacklisted control {well} ({cell}) with {replacement} on {barcode}")

    # Union cells that share any library compound. Multiple Labels rows and
    # multi-label cells can connect otherwise separate replicate/dose sets.
    parents = list(range(len(library)))

    def root(index: int) -> int:
        while parents[index] != index:
            parents[index] = parents[parents[index]]
            index = parents[index]
        return index

    first_cell_by_compound: Dict[str, int] = {}
    for index, compounds in enumerate(library_compounds):
        for compound in compounds:
            previous = first_cell_by_compound.setdefault(compound, index)
            parents[root(index)] = root(previous)
    bundles: Dict[int, List[int]] = {}
    for index in range(len(library)):
        bundles.setdefault(root(index), []).append(index)

    assigned_library_wells = 0
    for indices in bundles.values():
        required: Dict[str, int] = {}
        for index in indices:
            group = layout_by_well.get(library[index][0], "")
            required[group] = required.get(group, 0) + 1
        if any(len(pools.get(group, [])) < count for group, count in required.items()):
            if any(includes_control[index] for index in indices):
                raise PicklyPyConfigError(
                    f"Cannot preserve controls co-dispensed with a complete library "
                    f"replicate/dose set on destination plate {barcode}: insufficient "
                    "healthy wells in its Layout groups."
                )
            continue
        for index in indices:
            original_well, cell = library[index]
            group = layout_by_well.get(original_well, "")
            result.append((pools[group].pop(0), cell))
            assigned_library_wells += 1

    if library and not assigned_library_wells:
        raise PicklyPyConfigError(
            f"No complete library replicate/dose set fits on destination plate {barcode} "
            "after preserving controls within Layout groups."
        )
    return result


def _build_treatment_signatures_for_randomization(
    *,
    plate_pairs: Sequence[Tuple[str, str]],
    signature_parts_by_label: Mapping[str, Sequence[str]],
) -> Dict[str, str]:
    """Build per-well treatment signatures for Screen randomization.

    ``plate_pairs`` contains the template position and its plate-map label cell.
    A label can resolve to multiple dispense definitions (for example a vehicle
    backfill plus a library compound), so every resolved identifier and final
    concentration is included in the signature. Labels without a resolved
    source definition are omitted; this keeps unused positions unused on a
    partially filled final destination plate.
    """

    treatment_by_well: Dict[str, str] = {}
    for dstwell, label_cell in plate_pairs:
        parts: List[str] = []
        for label in split_labels(label_cell):
            parts.extend(signature_parts_by_label.get(label, ()))
        if not parts:
            continue
        signature = repr(tuple(sorted(parts)))
        previous = treatment_by_well.get(dstwell, "")
        treatment_by_well[dstwell] = previous + (";" if previous else "") + signature
    return treatment_by_well


def generate_screening_picklist(
    xlsx_path: str | Path,
    *,
    randomize_wells: bool = False,
    backfill_name: Optional[str] = None,
    backfill_volume: Optional[float] = None,
    intermediate_dilution: bool = False,
    flush_dst_after_src: bool = False,
) -> None:
    """Generate screening picklists and metadata files.

    This function mirrors `screeningpicklist.wls` behavior by default.
    With flush_dst_after_src, each destination takes LIB rows from only one
    resolved original source barcode; controls and intermediate sources are exempt.
    """

    xlsx_path = Path(xlsx_path)
    if not xlsx_path.exists():
        raise PicklyPyConfigError(f"Design file not found: {xlsx_path}")

    _print_banner()

    src, dst, lib, blk = _load_design_sheets(xlsx_path)
    header = _parse_header(dst, src)
    intermediates = None
    dilution_aliases: Dict[str, Tuple[str, float]] = {}
    if intermediate_dilution:
        from .intermediate import IntermediateDilutions, split_dilution

        intermediates = IntermediateDilutions(xlsx_path, header.min_volume_ul)
        if intermediates.barcode in header.barcodes_dst:
            raise PicklyPyConfigError('Barcode_INT must differ from every final destination barcode.')
    backfill_override = normalize_backfill_override(backfill_name, backfill_volume)
    src_data_rows = src[1:] if len(src) > 1 else []
    dmso_controls = {
        as_str(row[1]) for row in src_data_rows
        if len(row) > 1 and as_str(row[1]).casefold() == "dmso"
        and not any("@" in as_str(value) for value in row)
    }
    wildcard_reference_rows = [
        row for row in src_data_rows if len(row) > 1 and as_str(row[1]) == "*"
    ]
    if wildcard_reference_rows:
        reference_schema_rows = [wildcard_reference_rows[0]] + [
            row
            for row in src_data_rows
            if not as_str(row[0]).startswith("@")
        ]
    else:
        reference_schema_rows = src_data_rows
    # Survey volumes (optional)
    survey = load_survey_volumes(xlsx_path.parent)
    vols_ul = survey.volumes_ul
    if vols_ul:
        print(f"Using survey volumes from {len(survey.source_files)} file(s).")

    # Plate map and labels list parsing (supports only one plate map, per Wolfram comment).
    plate_map_row = _find_row(dst, "Plate Map:")
    if plate_map_row is None:
        raise PicklyPyConfigError(
            "DST worksheet is missing a 'Plate Map:' section.")

    plate_rows, plate_cols = _detect_plate_format(dst, plate_map_row)
    print(f"Detected plate format: {plate_rows} rows x {plate_cols} cols")

    labels_row = _find_row(dst[plate_map_row + 1:], "Labels")
    if labels_row is None:
        raise PicklyPyConfigError("DST worksheet is missing a 'Labels' section after the plate map.")
    labels_row = plate_map_row + 1 + labels_row

    # Parameters between plate map and labels are read from col A..B, but we only need barcodes, picklist name, well volume.
    # (In screeningpicklist.wls, Picklist Name, Barcode_DST, etc are typically in header anyway.)

    plate_grid = _extract_plate_grid(
        dst, plate_map_row + 1, left_col=1, n_rows=plate_rows, n_cols=plate_cols)
    b0 = _plate_map_pairs_from_grid(plate_grid)

    # Optional Layout map. Its labels define independent randomization pools:
    # a treatment can move only to another used well carrying the same Layout
    # label. In the absence of Layout:, all used wells form one pool.
    layout_by_well: Dict[str, str] = {}
    layout_row = _find_row(dst, "Layout:")
    if layout_row is not None:
        layout_grid = _extract_plate_grid(
            dst,
            layout_row + 1,
            left_col=1,
            n_rows=plate_rows,
            n_cols=plate_cols,
        )
        for row_index in range(plate_rows):
            for col_index in range(plate_cols):
                layout_by_well[well_name(row_index + 1, col_index + 1)] = (
                    layout_grid[row_index][col_index].strip()
                )

    if randomize_wells:
        print(
            "****** Well randomization enabled (per destination plate barcode) - "
            "this may take a minute - random seed is based on design file name ******"
        )

    # Labels table starts at labels_row+1 and runs until blank in col A.
    labellists00: List[List[List[Any]]] = [[]]
    labels_data_raw: List[List[Any]] = []
    r = labels_row + 1
    while r < len(dst):
        key = as_str(dst[r][0])
        if key == "":
            break
        row = dst[r]
        label = as_str(row[0])
        compound = as_str(row[1]) if len(row) > 1 else ""
        conc = row[2] if len(row) > 2 else ""
        labels_data_raw.append([label, compound, conc])
        r += 1

    plate_map_labels = [
        label
        for _well, label_cell in b0
        for label in split_labels(label_cell)
    ]
    reserved_indices = reserved_backfill_row_indices(
        labels_data_raw,
        plate_map_labels,
    )
    workbook_backfill = (
        None
        if backfill_override is not None
        else validate_single_workbook_backfill(
            labels_data_raw,
            reserved_indices,
            context="Screen design",
        )
    )
    using_backfill = backfill_override is not None or workbook_backfill is not None
    reformatted_lib_references = _common_lib_references(
        reference_schema_rows,
        len(lib[0]) if lib else 0,
        using_backfill=using_backfill,
    )
    intermediate_lib_references = (
        _common_lib_references(
            reference_schema_rows, len(lib[0]) if lib else 0, using_backfill=True,
        )
        if intermediates is not None else None
    )
    configured_backfill_compound = (
        backfill_override.compound_name
        if backfill_override is not None
        else (workbook_backfill[0] if workbook_backfill is not None else None)
    )
    explicit_backfill_wells: set[str] = set()
    configured_backfill_stock: Optional[float] = None
    if configured_backfill_compound is not None:
        explicit_backfill_wells = {
            as_str(row[0]).strip()
            for row in src[1:]
            if len(row) > 1
            and as_str(row[1]).strip() == configured_backfill_compound
            and as_str(row[0]).strip() != ""
            and not as_str(row[0]).strip().startswith("@")
        }
        if not explicit_backfill_wells:
            raise PicklyPyConfigError(
                f"Backfill compound '{configured_backfill_compound}' must be present in "
                "an explicit SRC row with a concrete source well."
            )
        if workbook_backfill is not None:
            configured_stocks: set[float] = set()
            for row in src[1:]:
                if (
                    len(row) <= 2
                    or as_str(row[1]).strip() != configured_backfill_compound
                    or as_str(row[0]).strip() not in explicit_backfill_wells
                ):
                    continue
                stock = as_number(row[2])
                if stock is None or not math.isfinite(float(stock)) or float(stock) <= 0:
                    raise PicklyPyConfigError(
                        f"Every explicit SRC row for Backfill compound "
                        f"'{configured_backfill_compound}' must have a positive finite "
                        "stock concentration."
                    )
                configured_stocks.add(float(stock))
            if len(configured_stocks) != 1:
                raise PicklyPyConfigError(
                    f"Backfill compound '{configured_backfill_compound}' must have one "
                    "consistent positive stock concentration across its explicit SRC rows."
                )
            configured_backfill_stock = next(iter(configured_stocks))

    # Unused Backfill rows are operational configuration. All other 0.0
    # concentration rows retain the legacy Screen behavior and are removed.
    reserved_index_set = set(reserved_indices)
    labels_data: List[List[Any]] = []
    for row_index, row in enumerate(labels_data_raw):
        if row_index in reserved_index_set:
            continue
        conc_num = as_number(row[2] if len(row) > 2 else None)
        if conc_num is not None and float(conc_num) == 0.0:
            continue
        labels_data.append(row)
    labellists00[0] = labels_data

    if intermediates is not None:
        # The suffix is a source-stock alias, not another LIB slot. Validate
        # precedence against the full definitions, including unused direct doses.
        preceding = set()
        for row in labels_data_raw:
            compound = as_str(row[1])
            alias = split_dilution(compound)
            if alias is not None:
                base, factor = alias
                if base not in preceding:
                    raise PicklyPyConfigError(
                        f"Intermediate compound {compound!r} requires a preceding "
                        f"base compound {base!r} in the DST Labels definitions."
                    )
                dilution_aliases[compound] = (base, factor)
            else:
                preceding.add(compound)

    if not labels_data:
        raise PicklyPyConfigError("DST 'Labels' list is empty.")

    def base_compound(compound: str) -> str:
        return dilution_aliases.get(compound, (compound, 1.0))[0]

    allocation_labels_data = [
        [row[0], base_compound(as_str(row[1])), *row[2:]]
        for row in labels_data
    ]

    # Blacklist parsing
    blacklist = _parse_blacklist(
        blk, plate_rows=plate_rows, plate_cols=plate_cols)
    

    if blacklist is not None:
        print(f"Blacklisted_wells_by_dst = {blacklist.blacklisted_wells_by_dst}")
    else:
        print("No blacklisted wells.")

    # Precompute labeldict and reverse mapping (used by blacklist handler)
    labeldict: Dict[str, str] = {as_str(row[0]): as_str(row[1]) for row in labels_data}
    labeldictrev: Dict[str, List[str]] = {}
    for row in labels_data:
        pm_label = as_str(row[0])
        comp_label = as_str(row[1])
        labeldictrev.setdefault(comp_label, []).append(pm_label)

    # Global state across destination plates (matches WL behavior)
    lib_index = 0  # 1-based row index for LIB data rows; 0 means none consumed yet.
    dst_count = 0
    loopsrc: List[str] = []
    loopdst: List[str] = []
    csv_rows: List[List[Any]] = []
    warnings = False

    # volumes tracker across all used wells
    volumes_ul: Dict[Tuple[str, str], float] = {}

    # table merge and condition labels
    table_merge_rows: List[List[Any]] = []
    table_merge_rows_blk: List[List[Any]] = []
    conditionlabels: Dict[str, List[List[Any]]] = {}  # dst barcode -> 16x24 table of "" or (name, concs)
    reformatted_lib_dispenses: List[_ReformattedLibDispense] = []
    total_volume_nl_by_destination: Dict[Tuple[str, str], float] = {}
    lib_references_by_index: Dict[int, _LibReferenceColumns] = {}

    # Support variables for blacklist within each destination plate
    def fix_one_blacklisted_well(
        bad_well: str,
        *,
        b_pairs: List[Tuple[str, str]],
        free_wells: Dict[str, List[str]],
        blacklisted_wells: List[str],
    ) -> List[Tuple[str, str]]:
        """Apply the Wolfram fixOneWellBlackList logic."""
        if blacklist is None:
            return b_pairs

        if bad_well not in blacklist.label_by_well:
            # Exclude controls or fixed wells from groups.
            return b_pairs

        # Find the {well,label} pair.
        matches = [p for p in b_pairs if p[0] == bad_well]
        if not matches:
            print(
                f"Blacklisting: {bad_well} - it is not on the dispense list any more, skipping..."
            )
            return b_pairs
        bad_pair = matches[0]
        bad_label = bad_pair[1]

        # Remove all instances of the corresponding *compound label*.
        compound_label = labeldict.get(bad_label, "")
        remove_labels = labeldictrev.get(compound_label, [])
        remove_wells = [p[0] for p in b_pairs if p[1] in remove_labels]

        group_label = blacklist.label_by_well[bad_well]
        replacement_candidates = free_wells.get(group_label, [])

        if replacement_candidates:
            replacement = replacement_candidates[0]
            free_wells[group_label] = replacement_candidates[1:]
            # Replace the bad well with replacement, keep same label.
            b_pairs = [(replacement, bad_label) if p == bad_pair else p for p in b_pairs]
            print(f"Switching {bad_well} to {replacement}")
            return b_pairs

        # No replacement within group: remove all instances of that compound.
        spare_wells = [w for w in remove_wells if w != bad_well]
        spare_wells = [w for w in spare_wells if w not in blacklisted_wells]
        if spare_wells:
            for w in spare_wells:
                g = blacklist.label_by_well.get(w)
                if g is None:
                    continue
                free_wells.setdefault(g, []).append(w)

        b_pairs = [p for p in b_pairs if p[0] not in remove_wells]
        print(
            f"Removing all instances of {remove_labels} wells: {', '.join(remove_wells)}"
        )
        return b_pairs

    # Helper to fill a SRC row from LIB using @Column references (addcompound)
    def addcompound(srcline: List[Any]) -> Optional[List[Any]]:
        nonlocal lib_index, destination_library_barcode, source_boundary_reached

        # Leave the first row of the next source group unconsumed. Explicit
        # controls still belong on this destination even when listed last.
        if source_boundary_reached and any("@" in as_str(value) for value in srcline):
            return None

        # Randomized runs must still resolve explicit controls after the last
        # LIB-backed SRC row, including named templates with controls listed last.
        if lib_index >= len(lib) - 1 and (
            not (randomize_wells or intermediate_dilution or flush_dst_after_src)
            or any("@" in as_str(value) for value in srcline)
        ):
            return None

        # Copy row, pad to at least 7 columns (to preserve 'new' flag column).
        row = list(srcline) + [""] * max(0, 7 - len(srcline))

        # Identify @-references and substitute from the next library row.
        notyet = True
        lib_row_used: Optional[List[Any]] = None

        def _sub_cell(val: Any) -> Any:
            nonlocal notyet, lib_index, lib_row_used

            s = as_str(val)
            if "@" not in s:
                return val

            def repl(match: re.Match) -> str:
                nonlocal notyet, lib_index, lib_row_used
                col_letters = match.group(1)
                if notyet:
                    notyet = False
                    if lib_index + 1 >= len(lib):
                        return ""
                    lib_row_used = lib[lib_index + 1]
                assert lib_row_used is not None
                col_idx_1b = excel_col_to_number(col_letters)
                col_idx_0b = col_idx_1b - 1
                if col_idx_0b < 0 or col_idx_0b >= len(lib_row_used):
                    return ""
                return as_str(lib_row_used[col_idx_0b])

            return re.sub(r"@([A-Za-z]+)", repl, s)

        new_row = [_sub_cell(v) for v in row]
        if lib_row_used is not None:
            if flush_dst_after_src:
                barcode = as_str(new_row[5])
                if barcode in ("", "*"):
                    raise PicklyPyConfigError(
                        "flush_dst_after_src requires a resolved source barcode "
                        f"for LIB worksheet row {lib_index + 2}."
                    )
                if destination_library_barcode is None:
                    destination_library_barcode = barcode
                elif barcode != destination_library_barcode:
                    source_boundary_reached = True
                    return None
            lib_index += 1
        # Only rows that actually dereferenced LIB receive LIB provenance.
        # Explicit controls otherwise inherit the running index and can be
        # mistaken for a library row on later destination plates.
        resolved_lib_index = lib_index if lib_row_used is not None else 0
        new_row.append(resolved_lib_index)
        if resolved_lib_index > 0:
            lib_references_by_index[resolved_lib_index] = (
                _lib_references_for_src_row(
                    srcline,
                    len(lib[0]) if lib else 0,
                )
            )
        return new_row

    lastbarcode: str = ""

    # Main loop: one destination plate per iteration, until the library is exhausted.
    # Build valid well order from groups grid (if available), else from plate map template
    if blacklist is not None and blacklist.groups_by_label:
        # Valid wells = any well with a group label, in plate order (top-left to bottom-right)
        valid_well_template: List[str] = []
        for r in range(plate_rows):
            for c in range(plate_cols):
                w = well_name(r + 1, c + 1)
                if w in blacklist.label_by_well:
                    valid_well_template.append(w)
    else:
        # No groups grid — valid wells are those in the plate map template
        valid_well_template = [w for w, _ in b0]

    # Label sequence from plate map template, in plate-map order
    label_sequence: List[str] = [label for _, label in b0]

    # Identify control wells: wells in b0 whose labels are NOT in the library-filled set.
    # These are fixed to their positions regardless of blacklist.
    # Library-filled labels use wildcard rows (compound = "*" in SRC) or @-references.
    src_rows_check = src[1:] if len(src) > 1 else []
    explicit_compounds = set()
    has_wildcard = False
    for r in src_rows_check:
        if as_str(r[1]) == "*":
            has_wildcard = True
        elif not as_str(r[0]).startswith("@"):
            explicit_compounds.add(as_str(r[1]))

    # Map label -> compound from labels_data
    label_to_compound: Dict[str, str] = {as_str(row[0]): as_str(row[1]) for row in labels_data}

    # Control labels = those whose compound is explicitly defined in SRC (not wildcard-filled)
    control_labels: set = set()
    library_labels: List[str] = []
    for lab in label_sequence:
        comp = label_to_compound.get(lab, "")
        if comp in explicit_compounds:
            control_labels.add(lab)
        else:
            library_labels.append(lab)

    # Control well assignments from original plate map (fixed positions)
    control_pairs: List[Tuple[str, str]] = [(w, lab) for w, lab in b0 if lab in control_labels]
    control_wells: set = {w for w, _ in control_pairs}

    # Main loop
    carry_forward_labels: List[str] = []
    

    while lib_index < len(lib) - 1 or carry_forward_labels:
        destination_library_barcode: Optional[str] = None
        source_boundary_reached = False
        dst_count += 1
        if dst_count > len(header.barcodes_dst):
            raise PicklyPyConfigError(
                "Error: please provide more destination bar codes in the DST worksheet Barcode_DST: field"
            )
        barcode_dst = header.barcodes_dst[dst_count - 1]

        # Get blacklisted wells for this plate
        clname_blk_entries: List[List[Any]] = []
        blacklisted_set: set = set()
        if blacklist is not None:
            blacklisted_wells = blacklist.blacklisted_wells_by_dst.get(barcode_dst, [])
            if blacklisted_wells:
                print(f"****** Preparing Blacklist for: {barcode_dst} ******")
                blacklisted_set = set(blacklisted_wells)
                for bad in blacklisted_wells:
                    clname_blk_entries.append([bad, barcode_dst, "rejected", 0])
                print(f"Blacklisted {len(blacklisted_wells)} wells on {barcode_dst}")

        # Available wells = valid template minus blacklisted minus control wells
        available_wells: List[str] = [
            w for w in valid_well_template
            if w not in blacklisted_set and w not in control_wells
        ]

        # Non-randomized allocation retains the legacy fixed-control behavior.
        # With randomization and rejected wells, reserve replacements first and
        # allocate complete library sets before consuming any LIB rows.
        b_pairs: List[Tuple[str, str]] = []

        if randomize_wells and blacklisted_set:
            b_pairs = _allocate_randomized_blacklist_plate(
                plate_pairs=b0,
                valid_wells=valid_well_template,
                blacklisted_wells=blacklisted_set,
                layout_by_well=layout_by_well,
                labels_data=allocation_labels_data,
                explicit_compounds=explicit_compounds,
                barcode=barcode_dst,
            )
        else:
            # Add control pairs (skip if control well is blacklisted on this plate)
            for w, lab in control_pairs:
                if w not in blacklisted_set:
                    b_pairs.append((w, lab))
                else:
                    print(f"Warning: control well {w} ({lab}) is blacklisted on {barcode_dst}, skipping control")

            # Labels to assign: carried-forward first, then fresh from template
            labels_to_assign: List[str] = list(carry_forward_labels)
            carry_forward_labels = []

            # Fill remaining wells from the repeating label template
            template_idx = 0
            while len(labels_to_assign) < len(available_wells) and template_idx < len(library_labels):
                labels_to_assign.append(library_labels[template_idx])
                template_idx += 1

            # Assign labels to available wells
            for i, w in enumerate(available_wells):
                if i < len(labels_to_assign):
                    b_pairs.append((w, labels_to_assign[i]))

            # Any leftover labels carry forward
            if len(labels_to_assign) > len(available_wells):
                carry_forward_labels = labels_to_assign[len(available_wells):]
                print(f"Carrying forward {len(carry_forward_labels)} displaced labels to next plate")

        print(f"Plate {barcode_dst}: {len(available_wells)} available wells, {len([p for p in b_pairs if p[1] not in control_labels])} library compounds assigned")

        if not b_pairs:
            break

        # Build c list (dstwell,label) by splitting any multi-label cells
        c_pairs: List[Tuple[str, str]] = []
        for dstwell, label_cell in b_pairs:
            for lab in split_labels(label_cell):
                c_pairs.append((dstwell, lab))

        used_labels = sorted(set(lab for _, lab in c_pairs))

        # Reduce label list to only used labels
        labels_data_used = [row for row in labels_data if as_str(row[0]) in used_labels]

        missing_labels = [lab for lab in used_labels if lab not in {as_str(r[0]) for r in labels_data_used}]
        if missing_labels:
            raise PicklyPyConfigError(
                "Error: One or more labels in the plate map are not defined in the label list: "
                + ", ".join(missing_labels)
            )

        compounds_used = [base_compound(as_str(r[1])) for r in labels_data_used]

        # Expand SRC inventory for this destination plate
        src_rows = src[1:] if len(src) > 1 else []
        wildcard_rows = [r for r in src_rows if as_str(r[1]) == "*"]
        if wildcard_rows:
            wildcard_row = wildcard_rows[0]
            explicit_rows = [r for r in src_rows if not as_str(r[0]).startswith("@")]
            explicit_compounds = {as_str(r[1]) for r in explicit_rows}
            blank_compounds = []
            for c in compounds_used:
                if c in explicit_compounds:
                    continue
                if c not in blank_compounds:
                    blank_compounds.append(c)

            blank_rows = []
            for comp in blank_compounds:
                # Replace every literal "*" in the template row with the compound label
                new = [comp if as_str(v) == "*" else v for v in wildcard_row]
                blank_rows.append(new)

            src2_seed = explicit_rows + blank_rows
        else:
            src2_seed = src_rows
            if flush_dst_after_src or intermediate_dilution or (randomize_wells and blacklisted_set):
                # Named LIB-backed SRC templates, like wildcard expansions,
                # must not consume a LIB row for a slot omitted by allocation.
                src2_seed = [
                    row for row in src_rows
                    if as_str(row[1]) in compounds_used
                    or not any("@" in as_str(value) for value in row)
                ]

        # Apply addcompound to seed rows
        src2: List[List[Any]] = []
        for r0 in src2_seed:
            filled = addcompound(r0)
            if filled is None:
                continue
            src2.append(filled)

        if source_boundary_reached:
            # Restart the template for the next source, just as for a full
            # destination. Never carry an old source's displaced slot labels.
            carry_forward_labels = []

        if flush_dst_after_src:
            # A combined treatment cannot be split at a source boundary.
            # Detect both multi-label cells and duplicate Labels-table entries
            # before producing incomplete combinations or allocating dilutions.
            library_slots = {
                as_str(row[1]) for row in src2_seed
                if any("@" in as_str(value) for value in row)
            }
            resolved_slots = {
                as_str(row[1]) for row in src2
                if isinstance(row[-1], int) and row[-1] > 0
            }
            slots_by_label: Dict[str, set] = {}
            for label, compound, *_ in labels_data_used:
                slot = base_compound(as_str(compound))
                if slot in library_slots:
                    slots_by_label.setdefault(as_str(label), set()).add(slot)
            for well, label_cell in b_pairs:
                required_slots = set().union(*(
                    slots_by_label.get(label, set()) for label in split_labels(label_cell)
                ))
                if required_slots & resolved_slots and required_slots - resolved_slots:
                    raise PicklyPyConfigError(
                        f"flush_dst_after_src cannot fill the complete library combination "
                        f"at {barcode_dst}:{well} from one source group. "
                        "Keep all combined LIB slots in the same barcode group, "
                        "use separate destination slots, or disable this option."
                    )

        # Replace any "*" barcode in column 6 with the first non-* barcode in this src2 (Wolfram behavior)
        # Column indices: 0 well, 1 compound label, 2 stock, 3 identifier, 4 volume, 5 barcode, 6 newflag, 7? ...
        barcodes_in_src2 = [as_str(r[5]) for r in src2 if len(r) > 5 and as_str(r[5]) != "*"]
        if barcodes_in_src2:
            first_real_barcode = barcodes_in_src2[0]
            for r in src2:
                if len(r) <= 5:
                    r.extend([""] * (6 - len(r)))
                if as_str(r[5]) == "*":
                    r[5] = first_real_barcode

        # Apply survey volumes (if available)
        srccol = 4  # 0-based; WL uses column 5
        srcorig = copy.deepcopy(src2)
        if vols_ul:
            unknown_mapped_to: Optional[str] = None
            has_unknown = any(bc == 'UnknownBarCode' for (bc, _w) in vols_ul.keys())
            if has_unknown:
                used_barcodes = sorted({as_str(r[5]) for r in src2 if len(r) > 5 and as_str(r[5]) not in ('', '*')})
                if len(used_barcodes) > 1:
                    raise PicklyPyConfigError(
                        'Error: Echo survey file contains UnknownBarCode but multiple source plate barcodes are used: '
                        + ', '.join(used_barcodes)
                        + '. Rename/re-export the survey with the correct barcode.'
                    )
                if len(used_barcodes) == 1:
                    unknown_mapped_to = used_barcodes[0]
                    print('Survey barcode is UnknownBarCode; mapping survey volumes to source plate barcode {}.'.format(unknown_mapped_to))

            missing_any = False
            for r in src2:
                if len(r) <= max(srccol, 6):
                    r.extend([''] * (max(srccol, 6) + 1 - len(r)))
                plate = as_str(r[5])
                well = as_str(r[0])
                vol = vols_ul.get((plate, well))
                if vol is None and unknown_mapped_to == plate:
                    vol = vols_ul.get(('UnknownBarCode', well))
                if vol is None:
                    vol = 0.0
                    missing_any = True
                r[srccol] = float(vol)
            if missing_any:
                print('Warning: some wells were not found in the survey file(s); they were set to 0 uL.')

            newly_added = [i for i, r in enumerate(src2) if len(r) > 6 and as_str(r[6]).lower() == 'new']
            for i in newly_added:
                src2[i][srccol] = srcorig[i][srccol]

        # Build srcr: plate barcode -> compound label -> (wellsList, stockConc, libIndex)
        srcr: Dict[str, Dict[str, Tuple[List[str], float, int]]] = {}
        src_name: Dict[str, Dict[str, str]] = {}

        for r in src2:
            if len(r) < 6:
                continue
            well = as_str(r[0])
            comp_label = as_str(r[1])
            stock = float(as_number(r[2]) or 0.0)
            ident = as_str(r[3])
            vol = float(as_number(r[4]) or 0.0)
            plate = as_str(r[5])
            libi = int(r[-1]) if r and isinstance(r[-1], (int, float)) else int(as_number(r[-1]) or 0)

            if plate == "" or well == "" or comp_label == "":
                continue

            srcr.setdefault(plate, {})
            if comp_label not in srcr[plate]:
                srcr[plate][comp_label] = ([well], stock, libi)
            else:
                srcr[plate][comp_label][0].append(well)

            src_name.setdefault(plate, {})
            src_name[plate][well] = ident

            # Update volumes tracker (preserve previous remaining volumes across iterations)
            key = (plate, well)
            if key not in volumes_ul:
                volumes_ul[key] = vol

        # barcodesSRC in original order (Tally preserves order)
        barcodes_src_ordered = [b for b, _ in tally_ordered([as_str(r[5]) for r in src2 if len(r) > 5])]

        # Duplicate explicit compound wells overlap / duplicate library well definitions per barcode
        seen_wells: Dict[Tuple[str, str], str] = {}
        overlaps: List[Tuple[str, str, str, str]] = []
        for r in src2:
            if len(r) < 6:
                continue
            plate = as_str(r[5])
            well = as_str(r[0])
            ident = as_str(r[3])
            k = (plate, well)
            if k in seen_wells and seen_wells[k] != ident:
                overlaps.append((plate, well, seen_wells[k], ident))
            else:
                seen_wells[k] = ident
        if overlaps:
            msg_lines = [
                "Error in SRC sheet - explicit compound (control) wells overlap with library wells, "
                "or library has duplicate well definitions for a barcode!",
                "Conflicts:",
            ]
            for plate, well, a, b in overlaps:
                msg_lines.append(f"  {plate} {well}: {a} vs {b}")
            raise PicklyPyConfigError("\n".join(msg_lines))

        if intermediates is not None:
            if intermediates.barcode in barcodes_src_ordered:
                raise PicklyPyConfigError('Barcode_INT must differ from every original source barcode.')
            # Survey replacement and original volume initialization have already
            # happened. Synthetic intermediate wells must not be zeroed by an
            # original-source survey, nor consume another original LIB row.
            for compound in dict.fromkeys(as_str(row[1]) for row in labels_data_used):
                if compound not in dilution_aliases:
                    continue
                base, factor = dilution_aliases[compound]
                matches = [(plate, srcr[plate][base]) for plate in barcodes_src_ordered
                           if base in srcr.get(plate, {})]
                if not matches:
                    # Unresolved slot on the final partial destination plate.
                    continue
                if len(matches) != 1:
                    raise PicklyPyConfigError(
                        f'Intermediate base {base!r} resolves to more than one source plate.'
                    )
                original_plate, (wells, stock, libi) = matches[0]
                if libi <= 0:
                    raise PicklyPyConfigError(
                        f'Intermediate base {base!r} must resolve to a LIB-backed source.'
                    )
                allocation = intermediates.allocate(
                    source_plate=original_plate, wells=wells, stock=stock,
                    lib_index=libi, identifier=src_name[original_plate][wells[0]],
                    factor=factor, volumes_ul=volumes_ul,
                    source_minimum_ul=header.min_volume_ul,
                )
                srcr.setdefault(intermediates.barcode, {})[compound] = (
                    [allocation.well], allocation.stock, libi,
                )
                src_name.setdefault(intermediates.barcode, {})[allocation.well] = allocation.identifier
            if srcr.get(intermediates.barcode):
                barcodes_src_ordered.append(intermediates.barcode)

        # If switching SRC plate barcode between destination plates, Plate:Works may need extra loop.
        if lastbarcode and barcodes_src_ordered and lastbarcode != barcodes_src_ordered[0]:
            loopdst.append(barcode_dst)
        if barcodes_src_ordered:
            lastbarcode = barcodes_src_ordered[-1]

        # Build labellists replicated per source plate, and apply indirect concentration references (@Column)
        # Start with base list (used labels only)
        base_labels = [[as_str(r[0]), as_str(r[1]), r[2]] for r in labels_data_used]

        # Build compound->libindex lookup for currently used source plates (flatten srcr)
        comp_to_libindex: Dict[str, int] = {}
        for plate in barcodes_src_ordered:
            for comp_label, (_wells, _stock, libi) in srcr.get(plate, {}).items():
                if comp_label not in comp_to_libindex:
                    comp_to_libindex[comp_label] = libi

        # Duplicate label rows support conditional co-dispenses such as DMSO
        # backfill followed by the wildcard library compound for the same
        # plate-map label.  On a partial final destination plate, the template
        # can contain more library slots than remain in LIB.  Suppress an
        # explicit co-dispense when that label's primary (final) compound
        # mapping was not populated from LIB on this destination plate.
        available_compounds = {
            comp_label
            for plate_map in srcr.values()
            for comp_label in plate_map
        }

        def resolve_conc_value(comp_label: str, conc_cell: Any) -> Any:
            s = as_str(conc_cell)
            if not s.startswith("@"):
                return conc_cell
            if comp_label not in comp_to_libindex:
                return conc_cell
            libi = comp_to_libindex[comp_label]
            col_letters = s[1:]
            col_idx = excel_col_to_number(col_letters) - 1
            lib_row_idx = libi
            if lib_row_idx < 0 or lib_row_idx >= len(lib):
                return conc_cell
            row = lib[lib_row_idx]
            if col_idx < 0 or col_idx >= len(row):
                return conc_cell
            return row[col_idx]

        base_labels_resolved: List[List[Any]] = []
        for lab, comp, conc in base_labels:
            base_labels_resolved.append([lab, comp, resolve_conc_value(comp, conc)])

        # For each source plate in order, build a label list whose compound entries are replaced with srcr mapping.
        labellists_by_src: List[List[Tuple[str, Tuple[List[str], float, int], float]]] = []
        for plate in barcodes_src_ordered:
            defs: List[Tuple[str, Tuple[List[str], float, int], float]] = []
            plate_map = srcr.get(plate, {})
            for lab, comp, conc in base_labels_resolved:
                if comp not in plate_map:
                    continue
                primary_compound = labeldict.get(lab, comp)
                if (
                    comp in explicit_compounds
                    and comp != primary_compound
                    and primary_compound not in available_compounds
                ):
                    continue
                wells, stock, libi = plate_map[comp]
                conc_num = as_number(conc)
                if conc_num is None:
                    raise PicklyPyConfigError(
                        f"Label list concentration for '{lab}' (compound '{comp}') must be numeric."
                    )
                defs.append((lab, (wells, float(stock), int(libi)), float(conc_num)))
            labellists_by_src.append(defs)

        # Build one destination-specific permutation after all actual LIB-backed
        # definitions are known. This makes aliases with the same resolved
        # identifier/concentration count as replicates and excludes empty slots
        # on a partially filled final plate.
        c_pairs_use = c_pairs
        if randomize_wells:
            signature_parts_by_label: Dict[str, List[str]] = {}
            adjacency_allowed_by_label: Dict[str, List[bool]] = {}
            exempt_control_sources = {
                (as_str(row[5]), as_str(row[0])) for row in src2
                if row[-1] == 0 and as_str(row[1]) in dmso_controls
            }
            for plate_i, plate in enumerate(barcodes_src_ordered):
                for lab, (wells, _stock, _libi), final in labellists_by_src[plate_i]:
                    ident = src_name.get(plate, {}).get(wells[0], wells[0])
                    signature_ident = ident or f"{plate}:{wells[0]}"
                    final_text = as_str(integer_chop(float(final)))
                    signature_parts_by_label.setdefault(lab, []).append(
                        f"{signature_ident}_{final_text}"
                    )
                    adjacency_allowed_by_label.setdefault(lab, []).append(
                        _libi == 0 and all((plate, well) in exempt_control_sources for well in wells)
                    )

            treatment_by_well = _build_treatment_signatures_for_randomization(
                plate_pairs=b_pairs,
                signature_parts_by_label=signature_parts_by_label,
            )
            # Exempt only treatments composed entirely of explicit DMSO
            # controls. A compound mixed with DMSO remains strict, as does a
            # LIB compound whose identifier happens to match a control name.
            allowed_by_treatment: Dict[str, bool] = {}
            for well, label_cell in b_pairs:
                treatment = treatment_by_well.get(well)
                if not treatment:
                    continue
                allowed_parts = [
                    allowed for label in split_labels(label_cell)
                    for allowed in adjacency_allowed_by_label.get(label, [])
                ]
                allowed_by_treatment[treatment] = (
                    allowed_by_treatment.get(treatment, True)
                    and bool(allowed_parts) and all(allowed_parts)
                )
            dmso_treatments = {
                treatment for treatment, allowed in allowed_by_treatment.items() if allowed
            }
            seed = seed_from_text(f"{xlsx_path.name}|{barcode_dst}")
            randomization_kwargs = dict(
                treatment_by_well=treatment_by_well,
                layout_by_well=layout_by_well if layout_by_well else None,
                seed=seed,
                attempts=100,
                forbid_adjacent_replicates=True,
            )
            try:
                well_map = optimize_plate_well_map(**randomization_kwargs)
            except ValueError as e:
                if not dmso_treatments:
                    raise PicklyPyConfigError(
                        f"Well randomization failed for destination plate {barcode_dst}: {e}"
                    ) from e
                print(
                    f"Warning: strict well randomization failed for destination plate {barcode_dst}: {e} "
                    "Retrying with adjacent DMSO-only control wells allowed. "
                    "Compound-replicate spacing, occupied wells and Layout groups remain unchanged."
                )
                try:
                    well_map = optimize_plate_well_map(
                        **randomization_kwargs, adjacency_exempt_treatments=dmso_treatments,
                    )
                except ValueError as fallback_error:
                    raise PicklyPyConfigError(
                        f"Well randomization failed for destination plate {barcode_dst} "
                        f"even with adjacent DMSO controls allowed: {fallback_error}"
                    ) from fallback_error

            c_pairs_use = [
                (well_map.get(dstwell, dstwell), lab)
                for dstwell, lab in c_pairs
            ]
            print(
                f"Randomization map for {barcode_dst}: {len(well_map)} wells mapped "
                f"({len(treatment_by_well)} used wells)."
            )

        if barcode_dst not in conditionlabels:
            conditionlabels[barcode_dst] = [
                ["" for _ in range(plate_cols)] for _ in range(plate_rows)
            ]

        # Rejected coordinates are physical metadata and are recorded once per
        # destination plate, independently of how many source plates feed it.
        for bad, _bcd, _rej, _zero in clname_blk_entries:
            row_index, col_index = parse_well_name(bad)
            conditionlabels[barcode_dst][row_index - 1][col_index - 1] = (
                "rejected",
                0,
            )
        table_merge_rows_blk.extend(clname_blk_entries)

        # Plan normal dispenses for every source plate first. Backfill must use
        # the sum across all source plates, but is inserted into one existing
        # source-plate plan so Plate:Works never has to revisit a source plate.
        event_plans: List[
            Tuple[
                List[Tuple[str, List[str], float]],
                List[Tuple[str, str, float]],
            ]
        ] = []
        normal_events_all: List[Tuple[str, List[str], float]] = []
        plate_reformatted_lib_dispenses: List[_ReformattedLibDispense] = []
        provenance_by_event: Dict[Tuple[Any, ...], List[_ReformattedLibDispense]] = {}
        for plate_i, plate in enumerate(barcodes_src_ordered):
            label_defs = {}
            for lab, srcinfo, final in labellists_by_src[plate_i]:
                label_defs.setdefault(lab, []).append((srcinfo, final))

            cl_events: List[Tuple[str, List[str], float]] = []
            clname_events: List[Tuple[str, str, float]] = []

            for dstwell, lab in c_pairs_use:
                if lab not in label_defs:
                    continue
                for (wells, stock, libi), final in label_defs[lab]:
                    if intermediates is not None and (
                        not math.isfinite(stock) or stock <= 0
                        or not math.isfinite(final) or final <= 0
                    ):
                        raise PicklyPyConfigError(
                            f'Positive finite stock and final concentration required for {lab!r}.'
                        )
                    desired_nl = 1000.0 * header.well_volume_ul * (final / stock)
                    vol_nl = _round_dispense_nl(desired_nl)
                    if intermediates is not None and vol_nl <= 0:
                        raise PicklyPyConfigError(
                            f'Cell transfer for {lab!r} rounds to zero nL; change its stock or dose.'
                        )
                    # Rounding warning (matches WL, uses Max[2.5, Round] in message)
                    if desired_nl != 0:
                        possible = max(2.5, _round_dispense_nl(desired_nl))
                        if abs(desired_nl - possible) / abs(desired_nl) > 0.1:
                            ident = src_name.get(plate, {}).get(wells[0], wells[0])
                            print(
                                "Warning: Significant rounding error at low volume dispensing! "
                                f"Compound: {ident} Stock: {stock} Final: {final} "
                                f"Desired dispense volume: {desired_nl} Possible dispense volume: {possible}"
                            )
                            warnings = True

                    cl_events.append((dstwell, wells, vol_nl))

                    lib_references = lib_references_by_index.get(libi)
                    if vol_nl > 0 and libi > 0 and lib_references is not None:
                        plate_reformatted_lib_dispenses.append(
                            _ReformattedLibDispense(
                                lib_row_index=libi,
                                destination_barcode=barcode_dst,
                                destination_well=dstwell,
                                volume_nl=vol_nl,
                                stock_concentration=stock,
                                source_barcode=plate,
                                source_well=wells[0],
                            )
                        )
                        if intermediates is not None:
                            provenance_by_event.setdefault(
                                (plate, dstwell, tuple(wells), vol_nl), []
                            ).append(plate_reformatted_lib_dispenses[-1])

                    # Merge table / plate map metadata uses first source well -> identifier
                    ident = src_name.get(plate, {}).get(wells[0], wells[0])
                    clname_events.append((dstwell, ident, final))

            event_plans.append((cl_events, clname_events))
            normal_events_all.extend(cl_events)

        active_backfill_compound: Optional[str] = None
        active_backfill_target_nl: Optional[float] = None
        if backfill_override is not None:
            active_backfill_compound = backfill_override.compound_name
            active_backfill_target_nl = backfill_override.target_volume_nl
        elif workbook_backfill is not None:
            active_backfill_compound, mock_final = workbook_backfill

        if active_backfill_compound is not None:
            source_matches: List[
                Tuple[int, str, Tuple[List[str], float, int]]
            ] = []
            for plate_i, plate in enumerate(barcodes_src_ordered):
                source_info = srcr.get(plate, {}).get(active_backfill_compound)
                if source_info is not None:
                    explicit_wells = [
                        well
                        for well in source_info[0]
                        if well in explicit_backfill_wells
                    ]
                    if explicit_wells:
                        source_matches.append(
                            (
                                plate_i,
                                plate,
                                (explicit_wells, source_info[1], source_info[2]),
                            )
                        )
            if not source_matches:
                raise PicklyPyConfigError(
                    f"Backfill compound '{active_backfill_compound}' is not present in the "
                    f"resolved SRC inventory for destination plate '{barcode_dst}'."
                )

            backfill_plate_i, backfill_plate, backfill_source = source_matches[0]
            backfill_wells, backfill_stock, _backfill_libi = backfill_source
            if active_backfill_target_nl is None:
                if configured_backfill_stock is None:
                    raise PicklyPyConfigError(
                        f"Backfill compound '{active_backfill_compound}' on source plate "
                        f"'{backfill_plate}' is missing its validated stock concentration."
                    )
                backfill_stock = configured_backfill_stock
                desired_nl, active_backfill_target_nl = (
                    target_volume_from_mock_concentration(
                        well_volume_ul=header.well_volume_ul,
                        stock_concentration=backfill_stock,
                        mock_final_concentration=mock_final,
                        context=(
                            f"Backfill compound '{active_backfill_compound}' on source plate "
                            f"'{backfill_plate}'"
                        ),
                    )
                )
                if abs(desired_nl - active_backfill_target_nl) / abs(desired_nl) > 0.1:
                    print(
                        "Warning: Significant rounding error for Backfill target! "
                        f"Compound: {active_backfill_compound} Stock: {backfill_stock} "
                        f"Mock final: {mock_final} Desired volume: {desired_nl} "
                        f"Possible volume: {active_backfill_target_nl}"
                    )
                    warnings = True

            topups = backfill_remainders(
                normal_events_all,
                active_backfill_target_nl,
            )
            event_plans[backfill_plate_i][0].extend(
                (dstwell, backfill_wells, volume_nl)
                for dstwell, volume_nl in topups
            )
            print(
                f"Backfill enabled for {barcode_dst}: {active_backfill_compound} from "
                f"{backfill_plate}, target total transfer {active_backfill_target_nl:g} nL "
                f"({len(topups)} wells topped up)."
            )

        # Final daughter-well volumes include every actual positive transfer,
        # including combinations, explicit co-dispenses, and native backfill.
        # LIB-backed normal events use these totals for post-backfill dilution.
        for cl_events, _clname_events in event_plans:
            for dstwell, _wells, volume_nl in cl_events:
                if volume_nl <= 0:
                    continue
                key = (barcode_dst, dstwell)
                total_volume_nl_by_destination[key] = (
                    total_volume_nl_by_destination.get(key, 0.0) + volume_nl
                )
        if intermediates is None:
            reformatted_lib_dispenses.extend(plate_reformatted_lib_dispenses)

        # Generate metadata and picklist rows in the original source-plate order.
        for plate_i, plate in enumerate(barcodes_src_ordered):
            cl_events, clname_events = event_plans[plate_i]

            print(f"****** {header.picklist_name}.csv ******")
            loopsrc.append(plate)
            loopdst.append(barcode_dst)

            print(f"Dispensing from: {plate}")
            print(f"Dispensing into: {barcode_dst}")

            # Merge table associations + update conditionlabels
            for dstwell, ident, final in clname_events:
                table_merge_rows.append([dstwell, barcode_dst, ident, final])
                r, c = parse_well_name(dstwell)
                r0 = r - 1
                c0 = c - 1
                cell = conditionlabels[barcode_dst][r0][c0]
                if cell == "":
                    conditionlabels[barcode_dst][r0][c0] = (ident, final)
                else:
                    existing_ident, existing_conc = cell
                    if isinstance(existing_conc, list):
                        new_conc = existing_conc + [final]
                    else:
                        new_conc = [existing_conc, final]
                    conditionlabels[barcode_dst][r0][c0] = (f"{existing_ident}+{ident}", new_conc)

            # Optimize and choose source wells
            # Group by src wells list (compound), maintain order of first appearance
            grouped: Dict[Tuple[str, ...], List[Tuple[str, List[str], float]]] = {}
            group_order: List[Tuple[str, ...]] = []
            for dstwell, wells, vol in cl_events:
                key = tuple(wells)
                if key not in grouped:
                    grouped[key] = []
                    group_order.append(key)
                grouped[key].append((dstwell, wells, vol))

            groups_list = [grouped[k] for k in group_order]
            groups_list = reorder_source_groups(groups_list)

            last_pos = (1, 1)
            reordered_events: List[Tuple[str, List[str], float]] = []
            for grp in groups_list:
                grp_ordered, last_pos = reorder_destinations_within_group(grp, last_pos)
                reordered_events.extend(grp_ordered)

            # Choose source well (auto-switching) and build per-source-well groups
            chosen_events: List[Tuple[str, str, float]] = []
            for dstwell, wells, vol in reordered_events:
                chosen = None
                minimum_volume_ul = (
                    intermediates.minimum_volume_ul
                    if intermediates is not None and plate == intermediates.barcode
                    else header.min_volume_ul
                )
                for w in wells:
                    if volumes_ul.get((plate, w), 0.0) >= minimum_volume_ul + vol / 1000.0:
                        chosen = w
                        break
                if chosen is None:
                    ident0 = src_name.get(plate, {}).get(wells[0], wells[0])
                    raise PicklyPyVolumeError(f"Error: Out of Volume for {ident0}")
                volumes_ul[(plate, chosen)] = volumes_ul.get((plate, chosen), 0.0) - vol / 1000.0
                chosen_events.append((dstwell, chosen, vol))
                if intermediates is not None:
                    provenance = provenance_by_event.get((plate, dstwell, tuple(wells), vol))
                    if provenance:
                        reformatted_lib_dispenses.append(replace(provenance.pop(0), source_well=chosen))

            # Group by chosen source well, preserve order
            per_well: Dict[str, List[Tuple[str, str, float]]] = {}
            well_order: List[str] = []
            for ev in chosen_events:
                _, w, _ = ev
                if w not in per_well:
                    per_well[w] = []
                    well_order.append(w)
                per_well[w].append(ev)

            for w in well_order:
                evs = per_well[w]
                row = [plate, w, barcode_dst]
                for dstwell, _w, vol in evs:
                    if intermediates is not None and plate == intermediates.barcode:
                        # Account for the total once above, but emit acoustic
                        # transfers within the configured per-transfer limit.
                        units_left = int(round(vol / 2.5))
                        max_units = int(math.floor(intermediates.maximum_transfer_nl / 2.5))
                        while units_left:
                            units = min(units_left, max_units)
                            row.extend([units * 2.5, dstwell])
                            units_left -= units
                    else:
                        row.extend([vol, dstwell])
                csv_rows.append(row)

    # End main loop

    # Combine merge rows (normal + blacklist)
    table_merge_all = table_merge_rows + table_merge_rows_blk

    # Write outputs to the same folder as the design file.
    out_dir = xlsx_path.parent

    src_unique = [b for b, _ in tally_ordered(loopsrc)]
    dst_unique = [b for b, _ in tally_ordered(loopdst)]
    if intermediates is not None:
        # A common path is written once: preparation before cell dosing for
        # merged picklists; ordered union for shared inventories/process lists.
        intermediates.write_outputs(
            out_dir=out_dir, header=header, cell_rows=csv_rows,
            src_unique=[b for b in src_unique if b != intermediates.barcode],
            dst_unique=dst_unique, lib=lib, references=intermediate_lib_references,
            volumes_ul=volumes_ul,
        )
    else:
        picklist_path = out_dir / f"{header.picklist_name}.csv"
        write_text(picklist_path, format_csv_rows(csv_rows))
        print(f"Saving: {picklist_path}")
        inv_src_rows = [make_inventory_row(b, header.inventory_src_rack, i + 1) for i, b in enumerate(src_unique)]
        inv_dst_rows = [make_inventory_row(b, header.inventory_dst_rack, i + 1) for i, b in enumerate(dst_unique)]
        inv_src_text = format_csv_rows(inv_src_rows)
        inv_dst_text = format_csv_rows(inv_dst_rows)
        if header.inventory_src_name == header.inventory_dst_name:
            inv_path = out_dir / f"{header.inventory_src_name}.csv"
            write_text(inv_path, inv_src_text + "\n" + inv_dst_text)
            print(f"Saving: {inv_path}")
        else:
            inv_path_src = out_dir / f"{header.inventory_src_name}.csv"
            inv_path_dst = out_dir / f"{header.inventory_dst_name}.csv"
            write_text(inv_path_src, inv_src_text)
            write_text(inv_path_dst, inv_dst_text)
            print(f"Saving: {inv_path_src}")
            print(f"Saving: {inv_path_dst}")
        proc_src_path = out_dir / f"{header.process_src_name}.txt"
        proc_dst_path = out_dir / f"{header.process_dst_name}.txt"
        write_text(proc_src_path, "\n".join(src_unique))
        write_text(proc_dst_path, "\n".join(dst_unique))
        print(f"Saving: {proc_src_path}")
        print(f"Saving: {proc_dst_path}")

    # Table merge
    merge_path = out_dir / "table merge.txt"
    merge_lines = []
    for row in table_merge_all:
        fields = [quote_csv_field(x) for x in row]
        merge_lines.append(",".join(fields))
    write_text(merge_path, "\n".join(merge_lines))
    print(f"Saving: {merge_path}")

    # Plate map compounds + concentrations
    compounds_path = out_dir / "plate map compounds.txt"
    conc_path = out_dir / "plate map concentrations.txt"

    def _cell_compound(cell: Any) -> str:
        if cell == "":
            return ""
        name, _conc = cell
        return as_str(name)

    def _cell_conc(cell: Any) -> str:
        if cell == "":
            return ""
        _name, conc = cell
        if isinstance(conc, list):
            return "{" + ",".join(str(integer_chop(float(x))) if as_number(x) is not None else str(x) for x in conc) + "}"
        n = as_number(conc)
        if n is None:
            return as_str(conc)
        return str(integer_chop(float(n)))

    lines_comp: List[str] = []
    lines_conc: List[str] = []
    for bcd in dst_unique:
        lines_comp.append(bcd)
        lines_conc.append(bcd)
        table = conditionlabels.get(
            bcd) or [["" for _ in range(plate_cols)] for _ in range(plate_rows)]
        for r in range(plate_rows):
            lines_comp.append("\t".join(_cell_compound(
                table[r][c]) for c in range(plate_cols)))
            lines_conc.append("\t".join(_cell_conc(
                table[r][c]) for c in range(plate_cols)))

    write_text(compounds_path, "\n".join(lines_comp))
    write_text(conc_path, "\n".join(lines_conc))
    print(f"Saving: {compounds_path}")
    print(f"Saving: {conc_path}")

    reformatted_lib_path = (
        out_dir / f"{xlsx_path.stem}_reformatted_LIB.xlsx"
    )
    _write_reformatted_lib_workbook(
        reformatted_lib_path,
        lib=lib,
        dispenses=reformatted_lib_dispenses,
        total_volume_nl_by_destination=total_volume_nl_by_destination,
        destination_order=dst_unique,
        using_backfill=using_backfill,
        references=reformatted_lib_references,
    )
    print(f"Saving: {reformatted_lib_path}")

    # Summary
    print("\nRun complete.")
    print(f"CellPlateCount (destination plates): {len(dst_unique)}")
    if intermediates is not None:
        original_count = len([barcode for barcode in src_unique if barcode != intermediates.barcode])
        print(f"LibraryPlateCount (original source plates): {original_count}")
        print("IntermediatePlateCount: 1")
        print(f"IntermediateWellsPrepared: {len(intermediates.allocations)}")
    else:
        print(f"LibraryPlateCount (source plates): {len(src_unique)}")
    print(f"TotalDispenseEvents (plate switches): {len(loopdst)}")

    if warnings:
        print(
            "\nWARNING: Significant rounding error at low volume dispensing for one or more conditions."
        )


def main(argv: Optional[Sequence[str]] = None) -> int:
    import argparse

    def _str2bool(v: object) -> bool:
        if isinstance(v, bool):
            return v
        value = str(v).strip().lower()
        if value in {"true", "t", "1", "yes", "y"}:
            return True
        if value in {"false", "f", "0", "no", "n"}:
            return False
        raise argparse.ArgumentTypeError(
            "Expected a boolean value (true/false)."
        )

    parser = argparse.ArgumentParser(
        description="PicklyPy.Screen - generate screening picklists from SRC/DST/LIB worksheets"
    )
    parser.add_argument(
        "xlsx",
        nargs="?",
        help="Path to the design Excel file (legacy positional form).",
    )
    parser.add_argument(
        "-f",
        "--file",
        dest="xlsx_file",
        metavar="XLSX",
        help="Path to the design Excel file (preferred; matches PickliPy.Assay).",
    )

    parser.add_argument(
        "-r",
        "--randomize_wells",
        "--randomize-wells",
        dest="randomize_wells",
        type=_str2bool,
        default=False,
        metavar="{true|false}",
        help=(
            "Randomize destination well positions independently for each destination plate barcode. "
            "Only used wells are permuted. If a Layout: map is present in the DST worksheet, "
            "wells move only among positions with the same Layout label. With blacklisting, "
            "preserve control counts by reserving healthy replacements within those groups, "
            "then fit complete library replicate/dose sets; extra destination plates may be needed. "
            "Fail if controls cannot be preserved."
        ),
    )

    parser.add_argument(
        "--intermediate_dilution", "--intermediate-dilution",
        nargs="?", const=True, default=False, type=_str2bool,
        metavar="{true|false}",
        help=("Prepare -dilN compound aliases on one Intermediate_dilutions plate. "
              "Bare flag means true; false preserves normal Screen behavior."),
    )
    parser.add_argument(
        "--flush_dst_after_src", "--flush-dst-after-src",
        nargs="?", const=True, default=False, type=_str2bool,
        metavar="{true|false}",
        help=("Start a new destination plate when the resolved LIB source barcode changes. "
              "Controls and intermediate sources do not trigger a flush. "
              "Bare flag means true; default false allows source overlap."),
    )
    parser.add_argument(
        "--no-pause",
        action="store_true",
        help="Do not wait for Enter before exiting on error (useful for pipelines).",
    )
    parser.add_argument(
        "--backfill_name",
        "--backfill-name",
        dest="backfill_name",
        nargs="?",
        const=None,
        default=None,
        metavar="COMPOUND",
        help=(
            "SRC compound to use for operational backfill. Must be supplied together with "
            "--backfill_volume and overrides an unused Backfill row in the Labels table. "
            "If both backfill options are present without values, they are ignored."
        ),
    )
    parser.add_argument(
        "--backfill_volume",
        "--backfill-volume",
        dest="backfill_volume",
        type=float,
        nargs="?",
        const=None,
        default=None,
        metavar="NL",
        help=(
            "Target total transfer volume per treated destination well, in nL. "
            "Must be supplied together with --backfill_name. If both backfill "
            "options are present without values, they are ignored."
        ),
    )

    args = parser.parse_args(argv)

    if args.xlsx_file is not None and args.xlsx is not None:
        parser.error(
            "provide the design file either positionally or with -f/--file, not both"
        )
    xlsx_path = args.xlsx_file if args.xlsx_file is not None else args.xlsx
    if xlsx_path is None:
        parser.error(
            "a design Excel file is required; provide it with -f/--file or positionally"
        )

    from .common import run_with_user_facing_errors

    generate_kwargs: Dict[str, Any] = {
        "randomize_wells": args.randomize_wells,
    }
    if args.intermediate_dilution:
        generate_kwargs['intermediate_dilution'] = True
    if args.flush_dst_after_src:
        generate_kwargs['flush_dst_after_src'] = True
    if args.backfill_name is not None or args.backfill_volume is not None:
        generate_kwargs["backfill_name"] = args.backfill_name
        generate_kwargs["backfill_volume"] = args.backfill_volume

    return run_with_user_facing_errors(
        lambda: generate_screening_picklist(
            xlsx_path,
            **generate_kwargs,
        ),
        pause_on_exit=not args.no_pause,
    )

    


if __name__ == "__main__":
    raise SystemExit(main())
