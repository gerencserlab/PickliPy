---
name: picklipy-excel-design
description: Design, generate, validate, and troubleshoot PickliPy.Assay and PickliPy.Screen/PicklPy.Screen Excel workbooks for Echo 650 and PlateWorks picklists. Use for assay-ready plates, dose-response layouts, combinatorics, serial additions, library screening layouts, library shortlisting, destination well blacklisting, SRC/DST/LIB worksheet schemas, source volumes, barcodes, and plate-map design advice.
---

# PickliPy Excel Design Skill

Use this skill whenever the user is designing or editing Excel design workbooks for `PickliPy.Assay` or `PickliPy.Screen` / `PicklPy.Screen`, or when they ask how to arrange Echo/PlateWorks acoustic-dispensing experiments that will be converted into picklists.

This skill has two jobs:

1. Advise on the correct PickliPy workflow and layout design.
2. Generate or patch Excel workbooks using the helper script in `scripts/picklipy_design_builder.py` when deterministic workbook creation is useful.

Do not expose confidential tokens or internal access credentials in generated files, instructions, or examples.

## Start every task by classifying the design

Choose `PickliPy.Assay` when the user needs any of the following:

- A small explicit inventory of tool compounds, vehicles, dyes, or controls.
- Dose-response plates where the user controls the exact concentrations per destination well.
- Combinatorics: multiple labels/compounds in the same destination well.
- Serial additions, time courses, or multiple dispense events.
- Assay-ready plates dispensed from one or more known source plates.
- Positive controls added to plates that were otherwise populated by the screening workflow.

Choose `PickliPy.Screen` when the user needs any of the following:

- Reformatting a multi-source library into as many destination plates as required.
- A single destination layout template containing numbered slots such as `#1`, `#2`, … that will be filled from a `LIB` worksheet.
- Library shortlisting by filtering rows of a library table before dispensing.
- A `DST_Blacklist` worksheet for QC-based rejection or relocation of poor wells.
- Destination compound/concentration maps and table-merge files for downstream image analysis.

Choose a hybrid plan when a screen requires custom controls or perturbations that are not on every library plate: use `PickliPy.Screen` for the library and `PickliPy.Assay` for the additional tool-plate dispense into the same destination barcodes.

## Core geometry and placement rules

PickliPy workbooks use a 384-well grid unless the user explicitly asks for a 96-well plate. The grid is always 16 rows × 24 columns, rows `A`–`P`, columns `1`–`24`. The interpreted grid starts one cell down and one cell right of the label cell:

- `Plate Map:` in column A; map cells are immediately below and to the right, i.e. columns B:Y and 16 rows.
- `Concentrations Map:` in column A; same 16 × 24 geometry.
- `Groups:` in `DST_Blacklist`; same 16 × 24 geometry.

For a 96-well design, fill only the top-left 8 × 12 portion of the same 384-format map (`A1:H12`). Leave the remainder blank unless the user explicitly wants visible notes outside the processed region.

Labels are parsed from plate-map cells. Separators are comma, semicolon, plus, and newline. Therefore labels themselves must not contain those separators. Prefer short labels such as `FCCP`, `OLIGO`, `DMSO`, `#1`, `#1_d2`, `CTRLpos`, `veh`, or `fill`. Avoid spaces. Keep display colors/styles for human readability only; the generator reads text values.

## Dispense-volume reasoning

PickliPy computes Echo transfer volume from:

`transfer_nL = round_to_2.5_nL(1000 × destination_well_volume_uL × final_concentration / source_stock_concentration)`

The stock and final concentrations must use the same unit. The unit column is informational in Assay; the code does not convert units. Before making a design, check for:

- rounded volume of `0 nL`: invalid;
- rounded volume below the 2.5 nL increment or with large relative rounding error;
- LDV-source transfers above 500 nL per dispense, which may need a more concentrated stock or a split-dispense-aware workflow;
- total requested source volume that would drive a source well below the minimum volume in `SRC!G1`;
- evaporation risk in aqueous source plates, especially long Echo runs.

Use `python scripts/picklipy_design_builder.py estimate-volume ...` for quick volume checks.

## PickliPy.Assay workbook rules

The Assay workbook has `SRC` as the first worksheet and `DST` as the second worksheet.

### Assay `SRC` worksheet

Required columns:

1. `Source Well`
2. `Compound Name`
3. `Stock concentration`
4. `Unit (same unit as in DST)`
5. `Volume in source plate`
6. `Plate barcode`

Put the minimum retained source volume in `G1`; put a note such as `<-Min volume` in `H1`.

Rules:

- `(Plate barcode, Source Well)` must be unique.
- The same `Compound Name` may appear in multiple source wells only when stock concentration is the same; PickliPy will advance through source wells as earlier wells approach minimum volume.
- If a compound is present at multiple stock concentrations, use different compound names such as `FCCP_2mM` and `FCCP_10mM`.
- The source plate barcode referenced in each Assay plate-map section must exist in this worksheet.

### Assay `DST` worksheet

Top header rows in column A/B define:

- `Well volume (ul):`
- `Inventory_SRC:`
- `Inventory_DST:`
- `Inventory_SRC_Rack#:`
- `Inventory_DST_Rack#:`
- `Process_SRC:`
- `Process_DST:`

Each dispense event is a section containing:

1. `Plate Map:` and a 16 × 24 label grid.
2. Optional `Concentrations Map:` and a 16 × 24 concentration grid.
3. `Picklist Name:`
4. `Barcode_SRC:`
5. `Barcode_DST:`
6. `Labels` table with columns `Labels`, `Compounds`, `Concentrations`.

`Picklist Name:` controls file merging. Multiple sections with the same picklist name append to one picklist. Use this for serial additions to the same destination plates, for dose-response plus a second combinatorial treatment, or for applying multiple source plates to one destination plate. Use distinct picklist names for time-course additions that must be run separately.

`Barcode_DST:` may contain one destination barcode, a comma-separated list, or `*`. Use `*` only after the workbook contains other explicit destination barcodes that define the destination barcode universe.

### Assay dose-response

Use a `Concentrations Map:`. Each occupied well in `Plate Map:` should contain one label for one compound. The corresponding well in `Concentrations Map:` contains the final concentration. In this mode, the concentration values in the `Labels` list are ignored, but the `Labels` list still maps label → compound.

If a well needs more than one compound and each compound has a dose-response, create multiple dispense sections with the same `Picklist Name:`, `Barcode_SRC:`, and `Barcode_DST:`; put one compound/dose map in each section.

### Assay combinatorics

Do not use `Concentrations Map:` for a combinatorial section. Put multiple labels in a `Plate Map:` cell, separated by `+`, comma, semicolon, or newline. The `Labels` table gives the final concentration for each label. The same label may appear multiple times in the `Labels` table to dispense multiple compounds for one map label.

### Assay design patterns

- Single-compound dose response: put one label, e.g. `FCCP`, across the destination wells and use `Concentrations Map:` for a dilution series.
- Combination matrix: use row labels for compound A levels and column labels for compound B levels; write combined labels such as `A1+B3` in each well and define each label concentration in the `Labels` list.
- Serial addition: repeat complete sections; keep the same picklist name to append, or use timepoint-specific picklist names when the experiment must pause between additions.
- Edge-sensitive live-cell assays: reserve edges for vehicle/blank/control or leave blank; distribute treatment replicates across internal rows/columns.
- 96-well assay-ready plate: use only `A1:H12` in the 384-style map and leave all other cells blank.

## PickliPy.Screen workbook rules

The Screen workbook has `SRC`, `DST`, and `LIB` worksheets, plus optional `DST_Blacklist` as the fourth worksheet.

Use Screen when the library rows are the experimental unit. The `LIB` table should contain one header row and one row per compound/source well. The entire `LIB` table is dispensed; to dispense a subset, delete or filter out unwanted rows before generating the design file. After shortlisting, keep rows sorted by source plate barcode because the PlateWorks method cannot revisit an earlier source plate after switching away from it.

### Screen `SRC` worksheet

Required columns:

1. `Source Well`
2. `Compound Name`
3. `Stock concentration`
4. `Identifier for merging`
5. `Volume in source plate`
6. `Plate barcode`

Put the minimum retained source volume in `G1`.

Use explicit control rows for controls that exist on every library source plate. Their `Plate barcode` should be `*`, and that source well must exist on all library plates.

Use one or more `@ColumnLetter` references to pull library data from the `LIB` worksheet. A common pattern is:

- `Source Well` = `@A`
- `Compound Name` = `*`
- `Stock concentration` = `@C` or a fixed stock value
- `Identifier for merging` = `@B`
- `Volume in source plate` = fixed volume, e.g. `5`
- `Plate barcode` = `@D`

The wildcard `Compound Name` of `*` applies to any destination label-list compound that has not been explicitly defined as a control.

### Screen `DST` worksheet

Top header rows include the Assay header fields plus:

- `Loop Counts_SRC:`
- `Loop Counts_DST:`

Only one `Plate Map:` is used. It defines fixed controls plus numbered or otherwise unique library slots. `Barcode_DST:` is a comma-separated list of available destination plate barcodes; the generator uses only as many as needed to dispense the library.

The `Labels` list maps each plate-map label to an inventory item label and final concentration. A typical library slot uses the same text in `Labels` and `Compounds`, e.g. `#1` → `#1`. Multiple plate-map labels may map to the same compound item label to create dose levels for the same library row, e.g. `#1_d1`, `#1_d2`, `#1_d3` all map to `#1` with different concentrations. Repeating the same label in multiple wells creates replicates. Final concentration may be a number or an `@ColumnLetter` reference into `LIB`.

When using blacklisting, make the final row of the `Labels` table an appropriate fill control, such as `fill` → `vehicle`, because PickliPy uses the final label row to fill movable positions that become free after a rejected compound group is moved.

### Screen `LIB` worksheet and shortlisting

Keep the first row as headers. Required fields depend on the `@` references used in `SRC`, but at minimum the library table needs source well, source plate barcode, and an identifier/name. Include stock or final concentration columns when the design references them.

Shortlist by filtering the `LIB` table before writing the workbook. Do not mark unwanted library rows inside the sheet unless the Screen generator is known to ignore that marker; the Screen generator dispenses the rows present in `LIB`. Always sort the final shortlisted table by source plate barcode. Preserve any row order within a barcode if that order has operational meaning.

### Screen blacklisting

The optional `DST_Blacklist` worksheet supports quality-controlled dispensing after imaging/QC of destination plates.

It contains:

1. `Groups:` plus a 16 × 24 group map.
2. `Barcode_DST` and `Well` rows, where each destination barcode has a comma-separated list of rejected wells.

Group-map rules:

- Empty group-map cell means fixed/unmovable. Use empty cells for controls and any positions that should be rejected rather than relocated.
- A group label such as `1`, `2`, `left`, or `right` means positions may be relocated only within that group.
- Label only the used area of the plate. Do not label controls unless you intentionally allow controls to be relocated.
- If one compound appears in multiple wells, e.g. replicates or dose levels, the full compound group is moved together to the next plate; the script does not split a compound group across destination plates.
- Blacklisted wells are marked as `rejected` in merge/map outputs.

Use group maps to encode experimental constraints. For example, if the left half and right half are different cell types receiving matched library compounds, use group `left` for left-half library wells and group `right` for right-half library wells. Leave shared controls ungrouped.

## Layout advice by experimental goal

- **Uniform single-cell assay-ready plates:** use interleaved replicate placement and avoid putting all replicates in one row or column.
- **Dose response:** put concentrations along columns if the imaging/readout scans by rows; put concentrations along rows if users visually interpret by columns. Keep vehicle and positive controls on both sides of the dose series.
- **Combinatorial matrix:** use one axis per perturbation and ensure the `Labels` table contains all component labels. For multi-compound dose-response combinatorics, use multiple Assay sections with the same picklist name.
- **Library screen with controls:** reserve edge wells or regular sentinel wells for vehicle, positive, and negative controls. Thread library slots through the remaining wells in a deterministic order.
- **Two cell types on one plate:** split the map into halves or quadrants, duplicate each library slot once per group, and use `DST_Blacklist` group labels matching those halves/quadrants.
- **QC blacklisting expected:** include spare movable wells in the same group if possible, and choose the final `Labels` row as a benign fill treatment.
- **96-well destination plate:** use only the upper-left 8 × 12 coordinates. Do not accidentally populate rows I:P or columns 13:24.
- **High-volume transfers:** recommend increasing stock concentration where feasible, using PP source plates when appropriate, or confirming that the generator/method can split transfers above the LDV single-dispense limit.

## Recommended workflow for Codex

1. Interpret the user’s experimental goal and choose Assay, Screen, or hybrid.
2. Identify destination format, well volume, source plate format, minimum source volume, source stocks, destination barcodes, controls, number of replicates, and whether QC blacklisting is needed.
3. Propose the plate layout in words and identify any tradeoffs.
4. Use `scripts/picklipy_design_builder.py` to create a workbook when the user wants an actual `.xlsx` design file.
5. Run the helper’s `validate` command on generated files.
6. In the final response, report which sheets and design sections were created, what assumptions were used, and any warnings such as concentration rounding, insufficient destination barcodes, source-volume risk, or library sorting constraints.

## Helper script commands

From the skill directory:

```bash
python scripts/picklipy_design_builder.py assay-demo --output assay_demo.xlsx
python scripts/picklipy_design_builder.py screen-demo --output screen_demo.xlsx --slots 77 --dst-barcodes Plate1,Plate2,Plate3
python scripts/picklipy_design_builder.py assay-from-spec examples/assay_dose_response_spec.json --output assay_from_spec.xlsx
python scripts/picklipy_design_builder.py screen-from-spec examples/screen_shortlist_spec.json --output screen_from_spec.xlsx
python scripts/picklipy_design_builder.py validate assay_from_spec.xlsx --mode assay
python scripts/picklipy_design_builder.py validate screen_from_spec.xlsx --mode screen
python scripts/picklipy_design_builder.py estimate-volume --well-volume-ul 50 --stock 2000 --final 1
python scripts/picklipy_design_builder.py blacklist-from-csv harmony_export.csv --barcode-col Barcode --well-col Well --metric-col CellCount --min-value 500 --max-value 5000 --output blacklist.csv
```

The helper writes workbook values in the exact geometry expected by the PickliPy generators. It does not replace the production PickliPy picklist generator; it creates and validates the Excel design inputs.

## Files to consult

- `references/picklipy_schema_and_design.md` for workbook schema and design recipes.
- `references/layout_recipes.md` for examples of dose-response, combinatorics, screen slots, and blacklisting layouts.
- `references/troubleshooting.md` for common errors and fixes.
- `scripts/picklipy_design_builder.py` for deterministic workbook creation, validation, blacklisting CSV generation, and volume estimation.
