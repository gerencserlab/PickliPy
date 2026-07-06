# PickliPy Assay and Screen design reference

This reference summarizes how to design Excel workbooks for PickliPy picklist generation. Use it with `../scripts/picklipy_design_builder.py`.

## Assay versus Screen

| Situation | Use | Why |
|---|---|---|
| Known source wells and compounds | Assay | `SRC` explicitly lists every source well and stock. |
| Dose-response of one or several tool compounds | Assay | `Concentrations Map:` provides per-well final concentrations. |
| Combinatorics / multiple compounds in one destination well | Assay | A plate-map cell can contain multiple labels; labels map to compound/concentration rows. |
| Serial additions / time courses | Assay | Repeat full `Plate Map:` sections; same picklist name appends, different names separate timepoints. |
| Library reformatting from many source plates | Screen | `LIB` table is threaded through a fixed destination layout. |
| A library subset | Screen | Delete/filter `LIB` rows before generation and keep rows sorted by source plate barcode. |
| QC-based destination well rejection | Screen | Use `DST_Blacklist` with groups and per-destination well lists. |
| Library + special controls not present on library plates | Hybrid | Screen for the library, Assay for additional controls or perturbagens. |

## Common geometry

- Processed grid size: 16 rows × 24 columns.
- Rows: `A` through `P`; columns: `1` through `24`.
- Map data begins one cell below and one cell to the right of the map label.
- A `Plate Map:` label in cell `A14` means processed destination `A1` is in `B15`, destination `A2` is in `C15`, destination `B1` is in `B16`, etc.
- For 96-well mode, use only the top-left 8 rows × 12 columns of the 384 grid.

## Assay workbook schema

### Sheet 1: `SRC`

| Column | Header | Meaning |
|---|---|---|
| A | Source Well | Echo source well, e.g. `A1`. |
| B | Compound Name | Exact compound name referenced from `DST` label rows. |
| C | Stock concentration | Numeric stock concentration. Same unit as final concentrations. |
| D | Unit (same unit as in DST) | Informational only. |
| E | Volume in source plate | Available volume in µL. |
| F | Plate barcode | Source plate barcode. |
| G1 | minimum volume | Minimum retained volume in µL. |

Important rules:

- `(Source Well, Plate barcode)` must be unique.
- Repeated compound names are allowed only when the stock concentration is identical within a plate.
- The generator can switch to another source well with the same compound and stock when the first well approaches the minimum volume.
- Use suffixes for different stocks, e.g. `FCCP_2mM` and `FCCP_10mM`.

### Sheet 2: `DST`

Header labels in column A and values in column B:

- `Well volume (ul):`
- `Inventory_SRC:`
- `Inventory_DST:`
- `Inventory_SRC_Rack#:`
- `Inventory_DST_Rack#:`
- `Process_SRC:`
- `Process_DST:`

Each addition section contains:

1. `Plate Map:` followed by a 16 × 24 label grid.
2. Optional `Concentrations Map:` followed by a 16 × 24 numeric grid.
3. `Picklist Name:` value.
4. `Barcode_SRC:` value.
5. `Barcode_DST:` value. May be a single barcode, a comma-separated list, or `*` when prior explicit destination barcodes define the universe.
6. `Labels` table: rows of `label`, `compound`, `concentration`.

### Assay dose-response

Use one compound label per occupied destination well in the `Plate Map:`. Add `Concentrations Map:` with final concentrations in the same coordinates. The `Labels` table maps each label to a compound name; its concentration column is ignored in this mode.

Example:

- `Plate Map:`: `FCCP` in columns 2–11.
- `Concentrations Map:`: `0.01, 0.03, 0.1, 0.3, 1, 3, 10, 30` across those columns.
- `Labels` row: `FCCP` → `FCCP`, concentration can be `0` or a note value.

### Assay combinatorics

Do not use `Concentrations Map:`. Put multiple labels in a destination cell, e.g. `A1+B1`. Define `A1`, `B1`, etc. in `Labels` with exact final concentrations. A single plate-map label can have multiple rows in the `Labels` list when that label should dispense multiple compounds.

### Multiple Assay sections

- Same `Picklist Name:`: append to one output CSV.
- Different `Picklist Name:`: create separate output CSVs, useful for time courses.
- One source barcode per section. To use multiple source plates on the same destination plate, make separate sections with the same picklist name and same destination barcode(s).

## Screen workbook schema

### Sheet 1: `SRC`

| Column | Header | Meaning |
|---|---|---|
| A | Source Well | Explicit well or `@` reference to `LIB`. |
| B | Compound Name | Explicit control name or `*` wildcard for library slots. |
| C | Stock concentration | Numeric or `@` reference. |
| D | Identifier for merging | Name/identifier for output maps and merge files. |
| E | Volume in source plate | Numeric available volume or survey-updated value. |
| F | Plate barcode | Explicit source barcode, `*` for controls present on all plates, or `@` reference. |
| G1 | minimum volume | Minimum retained source volume in µL. |

For Screen, explicit controls are listed as real rows. Library compounds are pulled from `LIB` using `@ColumnLetter` references. The wildcard row with `Compound Name = *` is expanded for the slot labels in `DST` that are not explicitly defined controls.

### Sheet 2: `DST`

Header labels include the Assay headers plus:

- `Loop Counts_SRC:`
- `Loop Counts_DST:`

Screen uses a single `Plate Map:`. Place fixed controls and library slot labels. After the map, include:

- `Picklist Name:`
- `Barcode_DST:`: comma-separated list of available destination plate barcodes.
- `Labels`: map plate-map labels to compound item labels and final concentrations.

Library slots typically look like:

| Labels | Compounds | Concentrations |
|---|---|---|
| #1 | #1 | 10 |
| #2 | #2 | 10 |
| #3 | #3 | 10 |

Dose levels for the same library row use multiple labels pointing to the same compound item:

| Labels | Compounds | Concentrations |
|---|---|---|
| #1_d1 | #1 | 0.1 |
| #1_d2 | #1 | 1 |
| #1_d3 | #1 | 10 |

Replicates can use repeated occurrences of the same label in the `Plate Map:`. Blacklist handling then treats all instances of the same compound item as a group.

### Sheet 3: `LIB`

Use one header row and one row per source well/library compound. Only rows present in the table are dispensed. To shortlist a library, filter/delete rows before the Screen generator is run.

Always sort by source plate barcode so PlateWorks does not need to return to an earlier source plate. If rows are custom-ordered within a source plate, preserve that order.

### Sheet 4: `DST_Blacklist` (optional)

The blacklist sheet contains a group map and a destination-barcode/well list.

`Groups:` map:

- Empty cells: fixed positions. Blacklisted wells in these positions are rejected but not relocated.
- Non-empty cells: group labels. Wells relocate only within the same group.
- Do not put group labels on controls unless controls are intended to move.

`Barcode_DST` list:

| Barcode_DST | Well |
|---|---|
| Plate1 | A2,B5,C10 |
| Plate2 | D7,E8 |

If a compound has several associated wells, all are moved together to a later plate when one member is blacklisted. The final row in the Screen `Labels` table is used as the fill treatment for freed good positions.

## Operational checks before running PickliPy

- Does every label appearing in a `Plate Map:` have a row in the `Labels` table?
- Does every Assay `Labels` compound exist in `SRC` for the active `Barcode_SRC:`?
- Does every Screen wildcard/reference row resolve to valid `LIB` columns?
- Are `LIB` rows sorted by source plate barcode?
- Are there enough destination barcodes in Screen `Barcode_DST:`?
- Are stock/final concentration units consistent?
- Does the estimated transfer volume round to a valid Echo increment?
- Are source wells above minimum volume after all expected transfers?
- Are blacklisted controls intentionally fixed or intentionally movable?
- Are output file names unique enough to avoid overwriting prior PlateWorks files?
