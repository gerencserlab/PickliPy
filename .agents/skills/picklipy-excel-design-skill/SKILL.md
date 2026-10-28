---
name: picklipy-excel-design
description: Design, generate, validate, troubleshoot, and visualize PickliPy.Assay, PickliPy.Screen/PicklPy.Screen, and PickliPy.Bluetable workflows for Echo 650 and PlateWorks picklists. Use for assay-ready plates, dose-response layouts, combinatorics, serial additions, blue-table treatment/definition rows, library screening layouts, library shortlisting, destination well blacklisting, source volumes, barcodes, plate-map design, and dispense-path visualization.
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

Choose `PickliPy.Bluetable` when each destination plate is described by a 96-well blue-table workbook with up to three addition blocks, one source plate, optional manual additions, and `def` / `veh` helper rows.

## Core geometry and placement rules

PickliPy Assay and Screen workbooks use a 384-well grid unless the user explicitly asks for a 96-well plate. The grid is always 16 rows × 24 columns, rows `A`–`P`, columns `1`–`24`. The interpreted grid starts one cell down and one cell right of the label cell:

- `Plate Map:` in column A; map cells are immediately below and to the right, i.e. columns B:Y and 16 rows.
- `Concentrations Map:` in column A; same 16 × 24 geometry.
- Optional `Layout:` in column A; same geometry. Its cell text defines fixed physical regions for runtime well randomization: a treatment may move only between used wells whose `Layout:` cells contain identical text. Blank cells are one shared label group, as in `PickliPy.Assay`.
- `Groups:` in `DST_Blacklist`; same 16 × 24 geometry.

`DST!Layout:` and `DST_Blacklist!Groups:` are independent. `Layout:` constrains randomization; non-empty `Groups:` cells identify positions available to the current blacklist-compaction path. Do not substitute one for the other.

For a 96-well design, fill only the top-left 8 × 12 portion of the same 384-format map (`A1:H12`). Leave the remainder blank unless the user explicitly wants visible notes outside the processed region.

Labels are parsed from plate-map cells. Separators are comma, semicolon, plus, and newline. Therefore labels themselves must not contain those separators. Prefer short labels such as `FCCP`, `OLIGO`, `DMSO`, `#1`, `#1_d2`, `CTRLpos`, or `veh`. Avoid spaces. Keep display colors/styles for human readability only; the generator reads text values.

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

### Operational backfill

Use native backfill to equalize total Echo transfer volume without placing a vehicle label in every treated map cell. Add exactly one unused Labels-table row:

| Labels | Compounds | Concentrations |
|---|---|---|
| `Backfill` | exact vehicle compound from `SRC` | mock final concentration |

`Backfill` is case-insensitive only for recognizing this reserved row. It is reserved only when no case variant of `backfill` appears in that section's `Plate Map:`; when used in the map, it remains an ordinary treatment label. The mock final and the vehicle stock use the same units. PickliPy derives the target total transfer with the normal dilution formula, then adds `target − sum(already-rounded normal transfers)` from the vehicle source.

- Top up only wells with at least one positive normal compound transfer.
- Do not fill blank, unresolved, blacklisted, or zero-only wells.
- Add nothing when a well already receives the target volume or more.
- Resolve the vehicle by exact `SRC` compound label and apply normal source minimum-volume checks. Screen uses the first resolved source plate in source order and switches among equivalent vehicle wells on that plate.
- In Screen, sum normal events across every contributing source plate before backfill. In Assay, each addition section is balanced independently; avoid overlapping same-picklist sections with native backfill unless section-local balancing is intentional.
- Keep native backfill operational: it appears in the Echo picklist and source-volume accounting, but not in treatment labels, treatment signatures, merge rows, or compound/concentration condition maps. It does not create its own reformatted-LIB row, although its volume contributes to daughter-well concentration calculations. With randomization enabled, its destination follows the randomized normal treatment.

The paired CLI override takes precedence over the workbook row. Its volume is the target total per treated well, not the amount of vehicle to add, and is normalized to the Echo 2.5 nL increment:

```bash
python -m PickliPy.Assay -f assay.xlsx --backfill_name vehicle --backfill_volume 100 --no-pause
python -m PickliPy.Screen -f screen.xlsx --backfill_name vehicle --backfill_volume 100 --no-pause
```

Both underscore and hyphen option spellings are accepted. Supply both name and volume or neither. For fixed-argument pipeline compatibility, Assay and Screen also accept both option names without values and treat that bare pair as absent; it does not override a workbook Backfill row. See `references/picklipy_schema_and_design.md` for the full contract and `references/troubleshooting.md` for failure modes.

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

## PickliPy.Bluetable workbook rules

The first worksheet (normally `Plate`) contains three five-column addition blocks in `A:E`, `G:K`, and `M:Q`. Rows beneath the header may be destination-well rows such as `A1`, component definitions marked `def`, or vehicle definitions marked `veh`.

The top treatment/definition area is expandable. Users may insert additional well, `def`, or `veh` rows before the control area. Preserve these column-A control labels because the generator locates all later volume settings relative to them:

- `Starting assay volume (ul)`
- `Addition paradigm:` immediately below it

Do not rely on fixed Excel row numbers. Relative to `Addition paradigm:`, dilution is in column I on the same row, maximum reservoir volumes for additions 1–3 are in columns C/E/G one row below, and preparation volumes are in columns C/E/G three rows below.

For combined treatment names such as `Glucose+FCCP`, every component except the final component must have a matching `def` row in that addition block. A `veh` row both defines the vehicle treatment and enables automatic vehicle balancing for inventory rows assigned that vehicle type.

## PickliPy.Visualize options

`PickliPy.Visualize.visualize_picklist` renders source- and destination-plate dispense paths. Its `plate_color` option accepts exactly `black` or `white` and defaults to `black`. The command-line equivalent is `--plate-color black|white`. Use the white palette for light-background reports or printing; use the default black palette for the standard dark plate presentation.

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

An optional `Layout:` map may follow the plate map. Use it only to define randomization boundaries. Identically labeled positions form one pool even when they are spatially disconnected; differently labeled positions never exchange treatments. A unique layout label on a single used well pins that treatment to its original position. The physical layout map itself does not move.

The `Labels` list maps each plate-map label to an inventory item label and final concentration. A typical library slot uses the same text in `Labels` and `Compounds`, e.g. `#1` → `#1`. Multiple plate-map labels may map to the same compound item label to create dose levels for the same library row, e.g. `#1_d1`, `#1_d2`, `#1_d3` all map to `#1` with different concentrations. Repeating the same label in multiple wells creates replicates. Final concentration may be a number or an `@ColumnLetter` reference into `LIB`.

Screen runtime randomization is opt-in:

```bash
python -m PickliPy.Screen -f design.xlsx --randomize-wells true --no-pause
```

Screen accepts `-f/--file` like Assay. For compatibility, the legacy positional
design-file argument remains valid when `-f/--file` is omitted. `-r` and
`--randomize_wells` are equivalent switch names. With randomization enabled:

- Each destination barcode is randomized independently with a deterministic seed derived from the design-file basename and destination barcode.
- Only wells with actual resolved dispenses are permuted, so empty slots on a partial final plate remain empty and the used-well set is unchanged.
- Replicate identity is the complete resolved treatment, including every co-dispensed identifier and final concentration, rather than the raw plate-map label.
- Different labels that resolve to the same treatment and concentration are treated as replicates, while different dose levels are distinct treatments.
- The algorithm first requires non-orthogonally-adjacent replicates and maximizes separation within each layout-label pool. If strict Screen randomization fails, it warns and retries allowing adjacent explicit DMSO-only controls. All other compound replicates, including DMSO/compound mixtures, remain strict; an infeasible retry is an error. Controls qualify by SRC Compound Name `DMSO` (case-insensitive) with no LIB `@` references. A LIB stock named DMSO is not exempt. The fallback preserves occupied wells, control counts, Layout boundaries and complete dose sets; no new flag is needed. Successful strict permutations are unchanged, and Assay does not automatically use this fallback.
- Controls with actual dispenses participate in randomization. Give a control well a unique `Layout:` label when it must remain fixed; if that well is blacklisted, preserving the control is impossible and the run fails.
- With blacklisting and randomization enabled, reserve healthy replacement wells for rejected controls within the original `Layout:` group before allocating complete library replicate/dose sets. Blank/absent layout labels form one pool. Remaining LIB rows continue onto subsequent destination plates, so provide enough barcodes. If no complete library set fits after reserving controls, the run fails.
- Picklist targets, `table merge.txt`, `plate map compounds.txt`, `plate map concentrations.txt`, and the reformatted LIB export all use the randomized coordinates.

In a helper JSON spec, a `slot_layout` may set `randomization_layout_mode` to `none`, `all`, `halves`, or `quadrants`. These presets split the full 16 × 24 canvas; use a custom map for 96-well half/quadrant boundaries. When a preset map is generated, `pin_controls_for_randomization` defaults to `true` and assigns every generated control well a unique layout label; set it to `false` only when controls should participate. A custom layout may instead provide `randomization_layout` (or `layout_map`) as a well-to-label mapping or 16 × 24 grid. Keep the separate `groups` field for `DST_Blacklist` movable positions.

The current Python Screen path removes zero-concentration label rows and does not use a final fill-label row during blacklist compaction. Do not rely on a `fill` label to populate freed wells.

### Screen `LIB` worksheet and shortlisting

Keep the first row as headers. Required fields depend on the `@` references used in `SRC`, but at minimum the library table needs source well, source plate barcode, and an identifier/name. Include stock or final concentration columns when the design references them.

Shortlist by filtering the `LIB` table before writing the workbook. Do not mark unwanted library rows inside the sheet unless the Screen generator is known to ignore that marker; the Screen generator dispenses the rows present in `LIB`. Always sort the final shortlisted table by source plate barcode. Preserve any row order within a barcode if that order has operational meaning.

Every successful Screen run also writes `<design stem>_reformatted_LIB.xlsx` beside the design. This workbook contains exactly one `LIB` sheet and expands each positive, rounded, LIB-linked dispense into one row. The columns referenced by `SRC` source-well and source-barcode `@` fields are replaced by the final destination well and barcode; their headers identify them as dispense destinations. `Dispense volume (nL)`, `Original source well`, and `Original source plate barcode` are appended, with the latter two preserving the resolved pre-dispense source coordinates. Replicates therefore duplicate a LIB row, and a combination produces one row per LIB-linked component. Explicit SRC controls and operational backfill do not receive synthetic rows because they have no originating LIB record. Use one consistent set of well and barcode `@` columns across LIB-linked SRC templates, plus one consistent stock mapping when backfill is active, so the export has an unambiguous schema.

When backfill is active and `SRC` stock concentration uses an `@` reference, that LIB column becomes `Concentration after dispense`. For each exported component calculate `source stock × component rounded nL / total positive rounded nL delivered to that destination well`. The denominator includes all normal co-dispenses and operational backfill. Without active backfill, preserve the original stock-concentration value and header.

### Screen source boundaries

Use `--flush_dst_after_src [true|false]` or `--flush-dst-after-src [true|false]` when each destination must receive LIB compounds from only one original source barcode. Bare enables it; default/false retains overlapping sources. Screen compares resolved SRC barcodes in LIB row order, leaves the first row of the next group unconsumed, and restarts the destination template on a fresh barcode. Unused slots remain empty. A source can span multiple destinations; nonconsecutive A/B/A groups start separate destinations without sorting. Supply enough destination barcodes for partial plates.

Fixed controls, backfill and intermediate sources do not trigger this boundary. It works with destination randomization and intermediate preparation; repeated-stock intermediate allocations and volume accounting persist across flushes. Existing blacklist behavior still applies. Combinations containing multiple LIB slots must fit wholly within one source group; a partial combination raises an error. Rerun generated intermediate designs with the same flag because it is a CLI/API setting, not a worksheet field. Python API: `generate_screening_picklist(..., flush_dst_after_src=True)`.

### Screen intermediate dilutions

Use `--intermediate_dilution [true|false]` or `--intermediate-dilution [true|false]` for native Screen preparation of diluted stocks; a bare option enables it and the default is disabled. In DST's Compounds column, `#1-dil33` requests a 33-fold dilution of the preceding `#1` base slot from the same original LIB row. Use original-stock LIB rows, without adding manual diluted pairs. Repeated original-stock rows are allowed and retain their destination requests: matching physical barcode/well, compound identifier, stock and factor reuse one intermediate well and one preparation across slots/plates. Exports track every contributing LIB worksheet row. Different dilution factors remain separate; conflicting identities or stocks for one physical well are errors. Multiple factors are independent preparations from the original stock, not a serial chain.

Supply an `Intermediate_dilutions` worksheet with inventory/process/picklist names, one intermediate barcode, and a fixed 16 × 24 solvent-preload map in µL. Preparation and downstream doses share source accounting; actual stock after 2.5 nL rounding determines cell doses. Capacity and retained-volume failures are errors. The map is allocated in row-major order and is not randomized with cell destinations. Shared output names merge across source, destination and intermediate files; use distinct process names when the PlateWorks method uses separate aliases.

Screen writes intermediate LIB and editable intermediate-design workbooks in addition to normal outputs, preserving the input workbook and original treatment identities. Rerun the generated design with the intermediate flag. See `references/picklipy_schema_and_design.md` for exact fields, defaults, dilution math and output semantics. The helper workbook builder does not substitute for running the production Screen implementation.

### Screen blacklisting

The optional `DST_Blacklist` worksheet supports quality-controlled dispensing after imaging/QC of destination plates.

It contains:

1. `Groups:` plus a 16 × 24 group map.
2. A `Well Blacklist:` marker row containing `Barcode_DST` and `Well` headers, followed by one row per destination barcode with a comma-separated list of rejected wells.

Group-map rules:

- Empty group-map cells are not replacement destinations when a non-empty group map is present. Controls may remain blank in this map: with randomization enabled their relocation is governed by `DST!Layout:`, not by membership in `Groups:`.
- A non-empty group-map cell marks a movable/available blacklist position. The current shift-based Python path does not enforce different group-label text as separate relocation pools; use `Layout:` for randomization boundaries and validate blacklist compaction separately when regional isolation matters.
- Label the library area and any intentionally available spare wells. Group labels on controls do not enable relocation when randomization is off; enable randomization and give controls a shared `Layout:` pool when replacement is wanted.
- Without randomization, blacklisted controls are skipped and legacy compaction can truncate library replicate/dose sets. With randomization, control counts are preserved and complete compound sets (including connected combinations) are allocated within their original `Layout:` pools; omitted slots consume no LIB rows.
- Blacklisted wells are marked as `rejected` in merge/map outputs.

Blacklist allocation is resolved before runtime randomization. Rejected physical coordinates stay rejected and are never remapped. With randomization enabled, both blacklist allocation and subsequent randomization respect `DST!Layout:` boundaries. Distinct `Groups:` text does not define relocation pools; without randomization, legacy compaction still needs separate validation for regional isolation.

Use the group map to mark positions available for blacklist compaction, leaving fixed controls blank. Encode left/right, cell-type, or other hard randomization boundaries in `DST!Layout:`; do not depend on distinct group-label text for that isolation in the current shift-based Python path.

## Layout advice by experimental goal

- **Uniform single-cell assay-ready plates:** use interleaved replicate placement and avoid putting all replicates in one row or column.
- **Dose response:** put concentrations along columns if the imaging/readout scans by rows; put concentrations along rows if users visually interpret by columns. Keep vehicle and positive controls on both sides of the dose series.
- **Combinatorial matrix:** use one axis per perturbation and ensure the `Labels` table contains all component labels. For multi-compound dose-response combinatorics, use multiple Assay sections with the same picklist name.
- **Library screen with controls:** reserve edge wells or regular sentinel wells for vehicle, positive, and negative controls. Thread library slots through the remaining wells in a deterministic order.
- **Randomized library screen:** use a `Layout:` map for cell-type, batch, or plate-region boundaries. Reuse the same layout label throughout a region; use a unique label to pin a well. Replicate spacing is optimized only among wells in the same label pool.
- **Two cell types on one plate:** split the map into halves or quadrants, duplicate each library slot once per region, and use matching DST `Layout:` labels. With randomization enabled, blacklist allocation preserves complete compound sets across these pools. Without randomization, validate or separate blacklist handling because distinct `Groups:` text does not enforce boundaries.
- **QC blacklisting expected:** include spare movable wells and validate replicate/dose cardinality in the generated merge and plate-map outputs before dispensing.
- **96-well destination plate:** use only the upper-left 8 × 12 coordinates. Do not accidentally populate rows I:P or columns 13:24.
- **High-volume transfers:** recommend increasing stock concentration where feasible, using PP source plates when appropriate, or confirming that the generator/method can split transfers above the LDV single-dispense limit.

## Recommended workflow for Codex

1. Interpret the user’s experimental goal and choose Assay, Screen, or hybrid.
2. Identify destination format, well volume, source plate format, minimum source volume, source stocks, destination barcodes, controls, number of replicates, whether operational backfill is wanted and its target/vehicle, whether runtime randomization is wanted, required `Layout:` boundaries or pinned wells, and whether QC blacklisting is needed.
3. Propose the plate layout in words and identify any tradeoffs.
4. Use `scripts/picklipy_design_builder.py` to create a workbook when the user wants an actual `.xlsx` design file.
5. Run the helper’s `validate` command on generated files.
6. In the final response, report which sheets and design sections were created, what assumptions were used, and any warnings such as backfill scope/rounding, concentration rounding, insufficient destination barcodes, source-volume risk, or library sorting constraints.

## Helper script commands

From the skill directory:

```bash
python scripts/picklipy_design_builder.py assay-demo --output assay_demo.xlsx
python scripts/picklipy_design_builder.py screen-demo --output screen_demo.xlsx --slots 77 --dst-barcodes Plate1,Plate2,Plate3
python scripts/picklipy_design_builder.py assay-from-spec examples/assay_dose_response_spec.json --output assay_from_spec.xlsx
python scripts/picklipy_design_builder.py screen-from-spec examples/screen_shortlist_spec.json --output screen_from_spec.xlsx
python scripts/picklipy_design_builder.py assay-from-spec examples/assay_backfill_spec.json --output assay_backfill.xlsx
python scripts/picklipy_design_builder.py screen-from-spec examples/screen_backfill_spec.json --output screen_backfill.xlsx
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
