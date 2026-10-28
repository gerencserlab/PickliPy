# PickliPy Assay, Screen, and Bluetable design reference

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
| 96-well blue-table with up to three timed additions | Bluetable | One workbook per destination barcode supports well, `def`, and `veh` rows plus manual additions. |

## Common geometry

- Bluetable uses direct 96-well names (`A1`–`H12`) in its addition blocks.
- Processed grid size: 16 rows × 24 columns.
- Rows: `A` through `P`; columns: `1` through `24`.
- Map data begins one cell below and one cell to the right of the map label.
- A `Plate Map:` label in cell `A14` means processed destination `A1` is in `B15`, destination `A2` is in `C15`, destination `B1` is in `B16`, etc.
- An optional `Layout:` uses the same geometry. It is a fixed physical map: runtime randomization may exchange only genuinely used wells with identical layout-label text. Blank layout cells form one shared group.
- For 96-well mode, use only the top-left 8 rows × 12 columns of the 384 grid.

Do not confuse `DST!Layout:` with `DST_Blacklist!Groups:`. The former constrains randomization; non-empty cells in the latter identify blacklist-compaction positions. They remain separate inputs, and the current shift-based Python path does not enforce different `Groups:` text values as regional pools.

## Bluetable workbook schema

The first worksheet, normally `Plate`, uses addition blocks `A:E`, `G:K`, and `M:Q`. Each block accepts destination-well rows and helper rows whose first cell is `def` or `veh`. The treatment area may grow by inserting rows above the volume-control area.

The generator finds that boundary from `Starting assay volume (ul)` and `Addition paradigm:` in column A, ignoring case and surrounding spaces. Keep the labels adjacent. Relative to `Addition paradigm:`, the dilution factor is in column I on the same row, maximum reservoir volumes are in C/E/G one row below, and preparation volumes are in C/E/G three rows below. Never design against fixed row numbers for these fields.

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

### Native operational backfill

Backfill uses a reserved, unused Labels row in either an Assay section or the single Screen Labels table:

| Labels | Compounds | Concentrations |
|---|---|---|
| `Backfill` | `vehicle` | mock final concentration |

The reserved-word comparison is whitespace-trimmed and case-insensitive. If any case variant of `backfill` occurs as a token in the corresponding `Plate Map:`, the row is an ordinary treatment label instead of configuration. Without a CLI override, allow at most one unused Backfill row.

The compound must exactly match an explicit `SRC` compound label. The mock final must be positive, finite, and in the same unit as the vehicle stock. The target total is:

`target_nL = round_to_2.5_nL(1000 × well_volume_uL × mock_final / vehicle_stock)`

PickliPy rounds every normal transfer independently, sums those rounded values at each final physical destination well, and adds only the positive remainder:

- normal sum below target → add the difference as vehicle;
- normal sum equal to or above target → add no automatic vehicle;
- no positive normal compound transfer → do not backfill;
- blank/missing dose, unresolved partial-Screen slot, or blacklisted coordinate → do not backfill;
- randomized treatment → calculate and dispense the remainder at its randomized coordinate.

Screen totals all source plates feeding a destination before adding one backfill event. It selects the first resolved source plate containing the explicit vehicle, then uses normal source-well switching among matching wells on that plate; place sufficient alternative vehicle wells there. Assay balances each addition section independently. A workbook row belongs to its section; a global CLI override applies to every nonempty section. Overlapping same-picklist sections can each top up the same well, so put native backfill only in the intended balancing section or keep sections spatially disjoint.

Automatic vehicle is operational rather than a treatment annotation. It is included in Echo picklists, source depletion, and minimum-volume checks; Assay also records it in `picklist_after_dispense.xlsx`. Screen tracks depletion during generation but does not export a remaining-source-volume workbook. Backfill is excluded from treatment signatures and condition metadata (`table merge`, condition lists, and compound/concentration maps). Screen's reformatted-LIB export does not add a synthetic vehicle row, but includes the vehicle volume when calculating daughter-well concentrations.

CLI/API overrides require an exact SRC compound and a target-total volume in nL. They ignore an unused workbook Backfill row and accept underscore or hyphen spelling:

```bash
python -m PickliPy.Assay -f assay.xlsx --backfill_name vehicle --backfill_volume 100 --no-pause
python -m PickliPy.Screen -f screen.xlsx --backfill-name vehicle --backfill-volume 100 --no-pause
```

Supply both arguments or neither. Assay and Screen additionally permit both option names to appear without values, for example `--backfill_name --backfill_volume`; this bare pair is treated exactly like absent CLI options and therefore does not override workbook backfill configuration. The explicit volume is normalized once to the Echo 2.5 nL increment; it is not a fixed top-up amount.

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

The DST sheet may also contain one optional `Layout:` map. Identical layout labels define randomization pools; different labels are hard boundaries. A layout label used by only one occupied well pins that well. Identical labels may be spatially disconnected and still form one pool.

Runtime randomization is disabled by default. Enable it with:

```bash
python -m PickliPy.Screen -f design.xlsx --randomize-wells true --no-pause
```

Screen also accepts `design.xlsx` as a legacy positional argument when
`-f/--file` is absent.

The `-r` and `--randomize_wells` aliases accept the same explicit boolean value. Randomization is deterministic for a design-file basename and destination barcode, is calculated independently per destination plate after blacklist handling and LIB resolution, and permutes only actual used wells. Replicates are detected from the full resolved treatment signature including all identifiers/co-dispenses and final concentrations. Echo targets, merge rows, and compound/concentration maps all receive the same permutation.

Screen first forbids orthogonally adjacent replicates. If this fails and the destination contains eligible DMSO-only controls, it prints a warning and retries with adjacency allowed for those controls. No extra flag is needed. Eligibility requires an explicit SRC Compound Name equal to `DMSO` (case-insensitive), with no LIB references, and no other compounds in the treatment. A LIB compound named DMSO, a DMSO/compound mixture, or another control is not exempt. Strict compound groups are placed before exempt control groups; DMSO controls still participate in spreading where possible. Occupied wells, `Layout:` boundaries, control counts, complete dose sets and retained-volume checks are preserved. Successful strict permutations are unchanged. If the fallback cannot satisfy the remaining constraints, generation still fails. Assay retains its original strict behavior.

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

Replicates can use repeated occurrences of the same label in the `Plate Map:`. With randomization enabled, blacklist allocation treats all instances/dose aliases of the same compound item, including connected combinations, as an indivisible set within the original `Layout:` pools. Non-randomized legacy compaction can truncate these sets.

### Sheet 3: `LIB`

Use one header row and one row per source well/library compound. Only rows present in the table are dispensed. To shortlist a library, filter/delete rows before the Screen generator is run.

Always sort by source plate barcode so PlateWorks does not need to return to an earlier source plate. If rows are custom-ordered within a source plate, preserve that order.

### Destination flush at a LIB source boundary

Screen normally permits multiple original library source plates to fill a destination. `--flush_dst_after_src [true|false]`, also spelled `--flush-dst-after-src`, optionally prevents that overlap. Bare means true; omitted or false preserves legacy behavior. The Python keyword is `flush_dst_after_src=True`.

Compare the original source barcode after resolving SRC references against the next LIB row, preserving input order. At a change, leave that row for a new destination and restart the template there; unused slots on the old destination stay empty. One barcode may fill several destinations. Nonconsecutive A/B/A input starts separate groups rather than revisiting a prior destination. More `Barcode_DST:` entries may be required.

Fixed SRC controls, native backfill and the synthetic intermediate source are exempt. Direct and intermediate doses of each assigned library slot remain on its allocated destination. Intermediate wells and remaining-volume balances persist across flushes. Destination randomization runs after allocation, while intermediate allocation remains fixed. Blacklist behavior is unchanged, including the non-randomized compaction limitations. A combined treatment requiring multiple LIB slots must be completely populated within one source group; a source boundary or end of LIB that would leave it partial raises a configuration error.

Outputs, including picklists, inventories, processes and destination annotations, follow the new allocation. The flag is not stored in the workbook; include it again when rerunning an exported intermediate design.

### Mandatory reformatted LIB output

Every successful Screen generation writes `<original design stem>_reformatted_LIB.xlsx` in the design folder. The workbook is newly created and contains only one worksheet named `LIB`; it does not copy `SRC`, `DST`, `DST_Blacklist`, or other input sheets.

For every positive rounded normal dispense that actually resolves to a LIB row, the generator copies that original row and:

- replaces every LIB column referenced by `SRC!Source Well` with the final destination well;
- replaces every LIB column referenced by `SRC!Plate barcode` with the destination plate barcode;
- annotates rewritten headers as dispense destination values;
- appends `Dispense volume (nL)`, `Original source well`, and `Original source plate barcode`, in that order;
- duplicates rows for replicated destination wells and writes one row per LIB-linked component of a combination;
- uses post-blacklist and post-randomization coordinates.

The appended originals are the resolved source well and barcode before replacement. With direct references such as `@A` and `@D`, these are the exact data-only values from those input LIB cells; embedded references also retain any SRC prefix or suffix. Explicit controls and native backfill have no originating LIB row and therefore do not create standalone export rows. LIB-linked SRC templates must use one consistent set of `@` columns for source well and barcode, plus a consistent stock mapping when backfill is active, so the single output header has an unambiguous meaning.

If backfill is active and `SRC!Stock concentration` references a LIB column, that column is renamed as concentration after dispense and updated per component:

`daughter concentration = source stock × component rounded volume / total positive rounded volume in the destination well`

The denominator includes all library components, explicit co-dispenses, and native backfill delivered to that well. Use actual rounded volumes rather than the configured target, because a normal combination may already exceed the target. Without active backfill, the original stock column and values remain unchanged.

### Optional intermediate dilution preparation

Screen can prepare diluted source stocks before cell dosing when explicitly enabled:

```bash
python -m PickliPy.Screen -f screen.xlsx --intermediate_dilution --no-pause
python -m PickliPy.Screen -f screen.xlsx --intermediate-dilution true --no-pause
python -m PickliPy.Screen -f screen.xlsx --intermediate_dilution false --no-pause
```

The underscore and hyphen spellings accept an optional boolean. Bare means `true`; omitted means `false`. Use the mode only with a matching `Intermediate_dilutions` worksheet. It prepares one intermediate plate; it does not edit or schedule a PlateWorks `.asy` workflow.

#### DST aliases and LIB row identity

Append `-dilN` to a library placeholder in the DST **Compounds** column, where `N` is a finite factor greater than one. For example:

| Labels | Compounds | Concentrations |
|---|---|---|
| `#1_high` | `#1` | `10` |
| `#1_mid` | `#1` | `1` |
| `#1_low` | `#1-dil33` | `0.3` |

The base placeholder must precede its diluted alias in the complete Labels table; its direct-dose label does not have to appear in the plate map. `#1-dil33` resolves to the same original LIB row as `#1`; it does not advance the library or require an extra pre-diluted LIB row. Several factors such as `#1-dil33` and `#1-dil100` may refer to that base. Each is made directly from the original source, not from another intermediate. Use the original compound identifier for annotations; the dilution suffix selects a stock and must not turn the compound into a new treatment identity.

Repeated original-stock LIB rows are supported. Rows with the same barcode, source well, resolved compound name/identifier and stock concentration reuse one intermediate well for each requested factor, whether assigned to different slots on one destination, later destinations, or nonconsecutive LIB rows. The allocator tracks the actual original well used for preparation, not a row number or placeholder. Different factors get separate preparations; conflicting identities or stock concentrations for one physical source well raise an error, including conflicts first encountered on later destination plates.

Preserve repeated LIB rows and their order: every row still supplies its intended destination request. Only preparation/allocation is deduplicated, and source/intermediate remaining volumes accumulate across all uses rather than resetting for each repeated row. Insufficient retained intermediate volume is an error rather than an instruction to allocate another copy. Destination randomization does not change this reuse. Do not create interleaved original/diluted LIB pairs for this mode.

#### `Intermediate_dilutions` worksheet

Use exact labels in column A, with scalar values in column B. The map follows the common geometry: one row below and one column to the right of `Plate Map:`, 16 rows × 24 columns.

| Field | Meaning / default |
|---|---|
| `Inventory_DIL:` | Output inventory filename for the intermediate plate. |
| `Inventory_DIL_Rack#:` | Intermediate plate starting rack number. |
| `Process_DIL:` | Output process-list filename for the intermediate plate. |
| `Picklist Name:` | Preparation picklist filename. |
| `Barcode_INT:` | One physical intermediate plate barcode, distinct from original-source and final-destination barcodes. |
| `Plate Map:` | Positive, finite solvent preload in µL in each available well; blank cells reserve unavailable positions. |
| `Minimum volume (ul):` | Optional retained intermediate volume in µL; defaults to the source minimum in `SRC!G1`. |
| `Maximum volume (ul):` | Optional prepared-well capacity in µL; defaults to `12` for LDV. Change it to the applicable capacity for a different plate. |
| `Maximum transfer (nl):` | Optional per-transfer limit in nL for preparation and intermediate-to-cell dispensing; defaults to `500`. Larger transfers are split without changing their total volume. |

Map numbers are starting solvent volumes, not compound labels, stock concentrations, or Echo transfer volumes. Solvent preloading remains a separate preparation step. Wells are allocated in row-major order, A1–A24, B1–B24, and so on. `DST!Layout:` and destination randomization do not randomize this physical intermediate plate. Unused map positions are not automatically filled.

For preload `V` µL and requested factor `F`, calculate:

`preparation_nL = round_to_2.5_nL(1000 × V / (F − 1))`

`actual_stock = original_stock × preparation_nL / (1000 × V + preparation_nL)`

Cell-transfer calculations use this actual stock after preparation rounding, with the normal 2.5 nL Echo increment. Keep stock and cell final concentrations in the same units. A 4 µL preload and factor 33 produce 125 nL preparation and 4.125 µL total; 10,000 µM original stock becomes 303.030303 µM. Splitting a large preparation preserves its total rounded volume while respecting the configured transfer limit.

Original source budgets include preparation, direct doses, controls and any applicable backfill. Intermediate budgets include all subsequent doses and must preserve the configured minimum. A requested preparation that exceeds well capacity, a depleted source, insufficient retained intermediate volume, or exhaustion of the fixed intermediate well pool is an error, not permission to silently change the factor, preload or plate.

#### Files, names and rerunning the design

Use output filenames without directories; the appropriate `.csv` or `.txt` extension may be included or omitted. Names are compared case-insensitively. Preparation uses the intermediate sheet's `Picklist Name:`; cell dosing uses DST's name. Different names produce separate CSVs. A shared name produces one preparation-first CSV. Cell doses are ordered by original source plates first, then the intermediate plate. Identically named inventory/process outputs are merged across SRC, DST and DIL rather than overwritten. This includes the common choice of sharing `Inventory_SRC:` and `Inventory_DIL:`. Shared inventory rows must not assign two different barcodes to the same rack/slot or conflicting addresses to one barcode. Using the same filename for incompatible output types also fails. Use separate filenames when a downstream PlateWorks assay expects separate plate-alias process lists or storage inventories, and verify the resulting barcode order.

Normal Screen metadata still describes the cell destinations and original compound identities. The existing reformatted daughter-LIB workbook remains a separate output. Intermediate mode also produces:

- `<design stem>_intermediate_LIB.xlsx`: original LIB rows copied into an intermediate LIB, with referenced source well/barcode/stock fields rewritten to allocated intermediate values and added preparation provenance.
- `<design stem>_intermediate_design.xlsx`: editable copy with resolved SRC/DST/LIB values, effective intermediate settings/defaults, and added `Intermediate_LIB` and `Intermediate_Allocation` sheets. Cached formula results are written as values because saving formulas without recalculation would lose the caches Screen reads. The original workbook's formulas are unchanged. Rerun with the intermediate flag enabled. Keep its original-source LIB as the source for preparation; the exported intermediate LIB is not a replacement input for that full two-stage run.

The accompanying `<design stem>_intermediate_report.json` records settings, allocated wells, source identities, requested/achieved factors, preparation chunks and remaining intermediate volumes. Each allocation's `lib_indices` lists all contributing data-row indices (starting at 1); `lib_worksheet_rows` gives actual Excel row numbers (starting at 2). The standalone intermediate-LIB workbook contains one sheet named `LIB` with one record per physical preparation. The original source coordinates/stock and first `Original LIB row` remain for compatibility; `Original LIB worksheet rows` lists every contributing Excel row. The generated design retains the repeated original-stock rows in its input `LIB` sheet and lists the deduplicated preparations in `Intermediate_LIB`/`Intermediate_Allocation`.

The input workbook is not modified. Formula errors in the original design are not repaired by enabling the feature. Inspect the resulting workbooks, CSV order and source accounting; native PlateWorks acceptance, mixing and physical execution remain separate checks.

### Sheet 4: `DST_Blacklist` (optional)

The blacklist sheet contains a group map and a destination-barcode/well list.

`Groups:` map:

- Non-empty cells identify eligible replacement/library positions; different group-label values do not define separate pools. If the group map is entirely empty, occupied plate-map positions are used.
- Controls may have blank group cells. With randomization enabled, rejected controls relocate to eligible healthy wells in the same `DST!Layout:` group, regardless of their original group-map membership. Blank/absent layout labels share one pool.
- Without randomization, controls stay at their original positions and blacklisted controls are skipped, even if their group cells are populated.

`Well Blacklist:` table (the marker is required by the production parser):

| Well Blacklist: | Barcode_DST | Well |
|---|---|---|
|  | Plate1 | A2,B5,C10 |
|  | Plate2 | D7,E8 |

Neither allocation path uses a final fill-label row. With randomization enabled, controls reserve healthy replacements first, then complete library replicate/dose sets are allocated within their original `Layout:` pools. Unassigned slots consume no LIB rows; remaining rows continue on subsequent destination plates, which may require additional barcodes. A pinned blacklisted control, insufficient capacity to preserve controls, or no capacity for any complete library set raises a configuration error. Controls co-dispensed in a library set cannot be silently dropped when that set does not fit.

Without randomization, legacy compaction is unchanged and can truncate a repeated library set; validate output cardinality.

Blacklist rejection/relocation is applied before optional runtime randomization. Rejected coordinates remain physical `rejected` positions in the merge and plate-map outputs, and are excluded from randomization targets.

## Operational checks before running PickliPy

- Does every label appearing in a `Plate Map:` have a row in the `Labels` table?
- Does every Assay `Labels` compound exist in `SRC` for the active `Barcode_SRC:`?
- Does every Screen wildcard/reference row resolve to valid `LIB` columns?
- If an unused Backfill row is present, is there exactly one, does its compound resolve exactly in `SRC`, and are its mock final, stock, and derived target positive?
- Is Assay backfill intentionally section-local, without overlapping same-picklist sections that top up the same well more than once?
- Is there enough vehicle volume after preserving the minimum source volume?
- Are `LIB` rows sorted by source plate barcode?
- Are there enough destination barcodes in Screen `Barcode_DST:`?
- Are stock/final concentration units consistent?
- Does the estimated transfer volume round to a valid Echo increment?
- Are source wells above minimum volume after all expected transfers?
- Is randomization enabled when blacklisted control counts must be preserved, and do their `Layout:` pools contain eligible healthy replacements?
- If runtime randomization is enabled, do all intended exchanges remain within identical `Layout:` labels, and are uniquely labeled controls intentionally pinned?
- Are repeated treatments feasible without orthogonal adjacency inside each layout-label pool?
- With randomized blacklisting, are there enough healthy wells per `Layout:` pool for controls and complete library sets, and enough destination barcodes for remaining LIB rows?
- Are output file names unique enough to avoid overwriting prior PlateWorks files?
