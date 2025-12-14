from __future__ import annotations

import copy
import math
from dataclasses import dataclass
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
    PLATE_ROWS_384,
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


def _print_banner() -> None:
    print(
        "\nPicklyPy.Screen (Python translation of screeningpicklist.wls v6.6 10/25/2025)\n"
        "Generates screening picklists for a generic destination layout, auto-filled\n"
        "from a multi-plate compound library (LIB worksheet).\n"
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
    ws_blk = _get_by_name(["BLK", "Blk", "blk", "Blacklist", "BLACKLIST"]) or (
        wb.worksheets[3] if len(wb.worksheets) > 3 else None
    )

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


def _extract_plate_grid(dst: List[List[Any]], top_row: int, left_col: int = 1) -> List[List[str]]:
    grid: List[List[str]] = []
    for r in range(top_row, top_row + PLATE_ROWS_384):
        if r >= len(dst):
            raise PicklyPyConfigError(
                f"DST worksheet ended while reading a 16x24 plate map at row {top_row+1}."
            )
        row_vals = dst[r]
        row_out: List[str] = []
        for c in range(left_col, left_col + PLATE_COLS_384):
            v = row_vals[c] if c < len(row_vals) else ""
            # Screening script coerces numeric to string.
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                v = integer_chop(float(v))
            row_out.append(as_str(v))
        grid.append(row_out)
    return grid


def _plate_map_pairs_from_grid(grid: List[List[str]]) -> List[Tuple[str, str]]:
    pairs: List[Tuple[str, str]] = []
    for r in range(PLATE_ROWS_384):
        for c in range(PLATE_COLS_384):
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


def _parse_blacklist(blk: Optional[List[List[Any]]]) -> Optional[_Blacklist]:
    if blk is None:
        return None

    # Convert entire sheet to string matrix for key scanning.
    blk_str: List[List[str]] = [[as_str(v) for v in row] for row in blk]

    # Find "Groups Map:" and read the next 16 rows x 24 cols from B..Y.
    groups_row = None
    for i, row in enumerate(blk_str):
        if row and row[0] == "Groups Map:":
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
        if row and row[0] == "Well Blacklist:":
            blk_row = i
            break

    blacklisted_by_dst: Dict[str, List[str]] = {}
    if blk_row is not None:
        j = blk_row + 1
        while j < len(blk_str):
            barcode = blk_str[j][0].strip() if len(blk_str[j]) > 0 else ""
            if barcode == "":
                break
            wells_raw = blk_str[j][1].strip() if len(blk_str[j]) > 1 else ""
            wells = [w.strip() for w in wells_raw.split(",") if w.strip()]
            blacklisted_by_dst[barcode] = wells
            j += 1

    return _Blacklist(groups_by_label=groups_by_label, label_by_well=label_by_well, blacklisted_wells_by_dst=blacklisted_by_dst)


def _round_dispense_nl(desired_nl: float) -> float:
    # Mathematica Round[..., 2.5] (half-away-from-zero) is implemented in common for Assay;
    # For Screen we can reuse same logic, but we keep it here to match script.
    from .common import round_to_increment

    return round_to_increment(desired_nl, 2.5)


def generate_screening_picklist(
    xlsx_path: str | Path
    ) -> None:
    """Generate screening picklists and metadata files.

    This function mirrors `screeningpicklist.wls` behavior.
    """

    xlsx_path = Path(xlsx_path)
    if not xlsx_path.exists():
        raise PicklyPyConfigError(f"Design file not found: {xlsx_path}")

    _print_banner()

    src, dst, lib, blk = _load_design_sheets(xlsx_path)
    header = _parse_header(dst, src)

    # Survey volumes (optional)
    survey = load_survey_volumes(xlsx_path.parent)
    vols_ul = survey.volumes_ul
    if vols_ul:
        print(f"Using survey volumes from {len(survey.source_files)} file(s).")

    # Plate map and labels list parsing (supports only one plate map, per Wolfram comment).
    plate_map_row = _find_row(dst, "Plate Map:")
    if plate_map_row is None:
        raise PicklyPyConfigError("DST worksheet is missing a 'Plate Map:' section.")

    labels_row = _find_row(dst[plate_map_row + 1 :], "Labels")
    if labels_row is None:
        raise PicklyPyConfigError("DST worksheet is missing a 'Labels' section after the plate map.")
    labels_row = plate_map_row + 1 + labels_row

    # Parameters between plate map and labels are read from col A..B, but we only need barcodes, picklist name, well volume.
    # (In screeningpicklist.wls, Picklist Name, Barcode_DST, etc are typically in header anyway.)

    plate_grid = _extract_plate_grid(dst, plate_map_row + 1, left_col=1)
    b0 = _plate_map_pairs_from_grid(plate_grid)

    # Labels table starts at labels_row+1 and runs until blank in col A.
    labellists00: List[List[List[Any]]] = [[]]
    labels_data: List[List[Any]] = []
    r = labels_row + 1
    while r < len(dst):
        key = as_str(dst[r][0])
        if key == "":
            break
        row = dst[r]
        label = as_str(row[0])
        compound = as_str(row[1]) if len(row) > 1 else ""
        conc = row[2] if len(row) > 2 else ""
        # Remove 0.0 concentration rows (feature in v6.x).
        conc_num = as_number(conc)
        if conc_num is not None and float(conc_num) == 0.0:
            r += 1
            continue
        labels_data.append([label, compound, conc])
        r += 1
    labellists00[0] = labels_data

    if not labels_data:
        raise PicklyPyConfigError("DST 'Labels' list is empty.")

    # Blacklist parsing
    blacklist = _parse_blacklist(blk)

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
        nonlocal lib_index

        if lib_index >= len(lib) - 1:
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
                    lib_index += 1
                    if lib_index >= len(lib):
                        return ""
                    lib_row_used = lib[lib_index]
                assert lib_row_used is not None
                col_idx_1b = excel_col_to_number(col_letters)
                col_idx_0b = col_idx_1b - 1
                if col_idx_0b < 0 or col_idx_0b >= len(lib_row_used):
                    return ""
                return as_str(lib_row_used[col_idx_0b])

            return re.sub(r"@([A-Za-z]+)", repl, s)

        import re

        new_row = [_sub_cell(v) for v in row]
        # Append lib_index (matches Wolfram addcompound)
        new_row.append(lib_index)
        return new_row

    lastbarcode: str = ""

    # Main loop: one destination plate per iteration, until the library is exhausted.
    while lib_index < len(lib) - 1:
        dst_count += 1
        if dst_count > len(header.barcodes_dst):
            raise PicklyPyConfigError(
                "Error: please provide more destination bar codes in the DST worksheet Barcode_DST: field"
            )
        barcode_dst = header.barcodes_dst[dst_count - 1]

        b_pairs = list(b0)

        # Blacklist handling
        clname_blk_entries: List[List[Any]] = []
        if blacklist is not None:
            blacklisted_wells = blacklist.blacklisted_wells_by_dst.get(barcode_dst)
            if blacklisted_wells:
                print(f"****** Preparing Blacklist for: {barcode_dst} ******")

                # Free wells pool per group label
                free_wells: Dict[str, List[str]] = {}
                used_wells = [w for w, _ in b_pairs]
                for g, wells in blacklist.groups_by_label.items():
                    free_wells[g] = [w for w in wells if w not in used_wells and w not in blacklisted_wells]

                for bad in blacklisted_wells:
                    b_pairs = fix_one_blacklisted_well(
                        bad,
                        b_pairs=b_pairs,
                        free_wells=free_wells,
                        blacklisted_wells=blacklisted_wells,
                    )

                # Mark blacklisted wells as rejected in merge table
                for bad in blacklisted_wells:
                    clname_blk_entries.append([bad, barcode_dst, "rejected", 0])

                # Add remaining free wells using the last label in the labels list
                last_label = as_str(labels_data[-1][0])
                remaining_free = [w for wells in free_wells.values() for w in wells]
                b_pairs.extend([(w, last_label) for w in remaining_free])

                if remaining_free:
                    print(
                        "Adding remaining free wells (to last label): "
                        + ", ".join(remaining_free)
                    )

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

        compounds_used = [as_str(r[1]) for r in labels_data_used]

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

        # Apply addcompound to seed rows
        src2: List[List[Any]] = []
        for r0 in src2_seed:
            filled = addcompound(r0)
            if filled is None:
                continue
            src2.append(filled)

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
                wells, stock, libi = plate_map[comp]
                conc_num = as_number(conc)
                if conc_num is None:
                    raise PicklyPyConfigError(
                        f"Label list concentration for '{lab}' (compound '{comp}') must be numeric."
                    )
                defs.append((lab, (wells, float(stock), int(libi)), float(conc_num)))
            labellists_by_src.append(defs)

        # Now generate dispenses per source plate
        for plate_i, plate in enumerate(barcodes_src_ordered):
            print(f"****** {header.picklist_name}.csv ******")
            loopsrc.append(plate)
            loopdst.append(barcode_dst)

            if barcode_dst not in conditionlabels:
                conditionlabels[barcode_dst] = [["" for _ in range(PLATE_COLS_384)] for _ in range(PLATE_ROWS_384)]

            print(f"Dispensing from: {plate}")
            print(f"Dispensing into: {barcode_dst}")

            # Build dispense events for this source plate
            label_defs = {}
            for lab, srcinfo, final in labellists_by_src[plate_i]:
                label_defs.setdefault(lab, []).append((srcinfo, final))

            cl_events: List[Tuple[str, List[str], float]] = []
            clname_events: List[Tuple[str, str, float]] = []

            for dstwell, lab in c_pairs:
                if lab not in label_defs:
                    continue
                for (wells, stock, _libi), final in label_defs[lab]:
                    desired_nl = 1000.0 * header.well_volume_ul * (final / stock)
                    vol_nl = _round_dispense_nl(desired_nl)
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

                    # Merge table / plate map metadata uses first source well -> identifier
                    ident = src_name.get(plate, {}).get(wells[0], wells[0])
                    clname_events.append((dstwell, ident, final))

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

            # Apply blacklist rejected markers in conditionlabels
            for bad, _bcd, _rej, _zero in clname_blk_entries:
                r, c = parse_well_name(bad)
                conditionlabels[barcode_dst][r - 1][c - 1] = ("rejected", 0)

            # Add blacklist rows to merge table output
            for row in clname_blk_entries:
                table_merge_rows_blk.append(row)

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
                for w in wells:
                    if volumes_ul.get((plate, w), 0.0) >= header.min_volume_ul + vol / 1000.0:
                        chosen = w
                        break
                if chosen is None:
                    ident0 = src_name.get(plate, {}).get(wells[0], wells[0])
                    raise PicklyPyVolumeError(f"Error: Out of Volume for {ident0}")
                volumes_ul[(plate, chosen)] = volumes_ul.get((plate, chosen), 0.0) - vol / 1000.0
                chosen_events.append((dstwell, chosen, vol))

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
                    row.extend([vol, dstwell])
                csv_rows.append(row)

    # End main loop

    # Combine merge rows (normal + blacklist)
    table_merge_all = table_merge_rows + table_merge_rows_blk

    # Write outputs to the same folder as the design file.
    out_dir = xlsx_path.parent

    picklist_path = out_dir / f"{header.picklist_name}.csv"
    write_text(picklist_path, format_csv_rows(csv_rows))
    print(f"Saving: {picklist_path}")

    # Inventory
    src_unique = [b for b, _ in tally_ordered(loopsrc)]
    dst_unique = [b for b, _ in tally_ordered(loopdst)]

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

    # Process files
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
        table = conditionlabels.get(bcd) or [["" for _ in range(PLATE_COLS_384)] for _ in range(PLATE_ROWS_384)]
        for r in range(PLATE_ROWS_384):
            lines_comp.append("\t".join(_cell_compound(table[r][c]) for c in range(PLATE_COLS_384)))
            lines_conc.append("\t".join(_cell_conc(table[r][c]) for c in range(PLATE_COLS_384)))

    write_text(compounds_path, "\n".join(lines_comp))
    write_text(conc_path, "\n".join(lines_conc))
    print(f"Saving: {compounds_path}")
    print(f"Saving: {conc_path}")

    # Summary
    print("\nRun complete.")
    print(f"CellPlateCount (destination plates): {len(dst_unique)}")
    print(f"LibraryPlateCount (source plates): {len(src_unique)}")
    print(f"TotalDispenseEvents (plate switches): {len(loopdst)}")

    if warnings:
        print(
            "\nWARNING: Significant rounding error at low volume dispensing for one or more conditions."
        )


def main(argv: Optional[Sequence[str]] = None) -> int:
    import argparse

    parser = argparse.ArgumentParser(
        description="PicklyPy.Screen - generate screening picklists from SRC/DST/LIB worksheets"
    )
    parser.add_argument("xlsx", help="Path to the design Excel file")

    parser.add_argument(
        "--no-pause",
        action="store_true",
        help="Do not wait for Enter before exiting on error (useful for pipelines).",
    )

    args = parser.parse_args(argv)

    from .common import run_with_user_facing_errors

    return run_with_user_facing_errors(
        lambda: generate_screening_picklist(args.xlsx),
        pause_on_exit=not args.no_pause,
    )

    


if __name__ == "__main__":
    raise SystemExit(main())
