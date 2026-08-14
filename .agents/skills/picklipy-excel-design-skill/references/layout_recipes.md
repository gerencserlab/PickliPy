# PickliPy layout recipes

Use these patterns when advising a user or generating a workbook.

## Assay: one-compound dose response

Goal: one compound across a dilution series, with replicates.

Recommended design:

- Use `PickliPy.Assay`.
- One `Plate Map:` section.
- Put the compound label, e.g. `FCCP`, in every treatment well.
- Add `Concentrations Map:` with final concentrations at the same coordinates.
- Put vehicle controls in separate wells as a separate label with its own section or as `0`/blank concentration depending on the intended dispense.

Good placement:

- Columns encode dose; rows encode replicate.
- Interleave replicate rows when edge effects are expected.
- Put vehicle and positive controls at both left and right sides of the dose block.

## Assay: two-drug combinatorial matrix

Goal: all combinations of A and B.

Recommended design:

- Use `PickliPy.Assay` without `Concentrations Map:`.
- Use labels like `A0`, `A1`, `A2` for drug A and `B0`, `B1`, `B2` for drug B.
- In each well, write `A1+B2`, etc.
- Define every component label in the `Labels` list with its compound and concentration.

When dose-response plus combinatorics is needed, make one section for compound A with a concentration map and one section for compound B with a concentration map. Give both sections the same picklist name so they append to one picklist.

## Assay: serial additions

Goal: add one reagent at assay setup and another later.

Recommended design:

- Same destination barcodes throughout.
- Use separate sections.
- Use the same `Picklist Name:` when the additions are executed as one picklist.
- Use different picklist names such as `Plate_T0`, `Plate_T30`, `Plate_T60` when PlateWorks should run them separately as a time course.

## Screen: library with fixed controls

Goal: thread a library through a fixed destination layout with controls.

Recommended design:

- Use `PickliPy.Screen`.
- Put `#1`, `#2`, … in treatment positions.
- Put fixed controls like `veh`, `pos`, `neg`, `blank` in reserved wells.
- `Labels` rows for library slots map `#n` → `#n` with a fixed concentration or `@` concentration reference.
- Explicit controls go into `SRC`; library rows come from `LIB` via a wildcard row.

Good placement:

- Reserve outer columns/rows for vehicle/positive controls if edge effects are likely.
- Otherwise place sentinel controls every 6–8 columns or in both halves of the plate.
- Fill slots in a deterministic order; row-major or serpentine is easiest to audit.

## Screen: runtime-randomized library

Goal: randomize actual Screen treatments independently on every destination plate while retaining physical experimental boundaries.

Recommended design:

- Keep the `Plate Map:` as the deterministic auditable template.
- Add a same-size `Layout:` map. Use one repeated label per allowed movement region, such as `left`, `right`, `cell_A`, or `cell_B`.
- Use a unique layout label for a control well that must stay at its exact coordinate. Controls otherwise participate in randomization just like other resolved treatments.
- Run `python -m PickliPy.Screen design.xlsx --randomize-wells true --no-pause`.
- Expect the same used-well set as the resolved, post-blacklist plate: empty slots on a partial final plate remain empty.
- Audit `table merge.txt`, `plate map compounds.txt`, and `plate map concentrations.txt`; they contain the same randomized coordinates as the Echo picklist.

The seed is deterministic from the design-file basename and destination barcode. Copying the same-named workbook elsewhere reproduces the placement; renaming it changes the seed.

## Screen: library dose response

Goal: each library compound at multiple concentrations.

Recommended design:

- Use multiple labels for one compound item label.
- Example: `#1_d1`, `#1_d2`, `#1_d3` in the plate map all map to compound item `#1` in the `Labels` table.
- Concentrations can be fixed values or `@` references to `LIB` columns.

Blacklisting:

- Group all labels for one library row together by mapping them to the same compound item label.
- Use DST `Layout:` labels when multiple cell populations or plate regions must remain separate during runtime randomization; the current blacklist-compaction path does not enforce distinct `Groups:` text as regional boundaries.

Randomization:

- Replicate identity includes the resolved library identifier and final concentration. Same-dose repeats are spread as replicates; different doses are distinct treatments.
- Multiple plate-map labels that resolve to the same full treatment still count as replicates.
- Ensure each `Layout:` region has enough geometry to place repeats without shared-edge adjacency. Diagonal neighbors are allowed.

## Screen: duplicate cell populations on one plate

Goal: same library compounds applied to two cell cultures or two assay conditions.

Recommended design:

- Split the plate into left/right halves or quadrants.
- Place each slot once in each group, e.g. `#1` in left and `#1` again in right, or use paired labels that map to the same compound item.
- Use `left` and `right` labels in the DST `Layout:` map to keep runtime randomization regional.
- The current shift-based blacklist path does not enforce different `Groups:` text as separate regions. Validate blacklist-compacted outputs carefully or process the cell populations separately when blacklist relocation must remain regional.
- Leave controls ungrouped unless the user intentionally wants them movable.

## Screen: QC blacklist expected

Goal: some destination wells will be rejected after cell-count QC.

Recommended design:

- Include a `DST_Blacklist` sheet.
- Add a `Groups:` map for all movable library wells.
- Keep controls blank in the group map.
- Generate a blacklist CSV from Harmony or another analysis and paste it under the required `Well Blacklist:` marker and `Barcode_DST`/`Well` headers.
- Validate replicate/dose counts after blacklist compaction; the current Python shift path does not use a fill row and can truncate a repeated group at a capacity boundary.

## 96-well assay-ready plates

Goal: use the same PickliPy map format for 96-well plates.

Recommended design:

- Use only rows `A:H` and columns `1:12` of the map.
- Leave all cells outside the top-left 8 × 12 blank.
- Use barcodes that clearly encode the physical 96-well destination plate identity.

## Layout anti-patterns

- Do not use the same source barcode in non-contiguous blocks in a Screen `LIB` table.
- Do not put comma, semicolon, plus, or newline characters inside label names.
- Do not rely on cell color or formatting for logic.
- Do not put multiple labels in a dose-response `Plate Map:` section that has `Concentrations Map:`; split into multiple sections instead.
- Do not group blacklisted control wells unless controls can be moved safely.
- Do not use `DST_Blacklist!Groups:` as a substitute for `DST!Layout:`; they control different stages.
- Do not assume a Screen control stays fixed during runtime randomization; pin it with a unique `Layout:` label.
- Do not create a layout-label pool too small or narrow to separate its replicate counts; strict non-adjacent placement will fail rather than cross the boundary.
- Do not leave too few destination barcodes for a shortlisted Screen library.
