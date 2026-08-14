# PickliPy workbook troubleshooting

## “Missing label definition”

Cause:

- A text token in `Plate Map:` is absent from the `Labels` table.
- A label contains a separator character (`+`, `,`, `;`, newline) and is being split unexpectedly.
- The `Labels` table is not below the corresponding map section or has a blank row too early.

Fix:

- Enumerate all labels from the map and compare to the `Labels` rows.
- Replace label separators inside names with `_`.
- Ensure each map section has its own `Labels` table.

## “Missing definitions on SRC” or source barcode does not exist

Cause:

- A compound in an Assay `Labels` row is absent from `SRC` for the active `Barcode_SRC:`.
- `Barcode_SRC:` has a typo.
- The compound is listed under a different stock-concentration suffix.

Fix:

- Check exact case-sensitive compound names.
- Add the source well to `SRC` or correct `Barcode_SRC:`.
- If a compound has multiple stocks, use distinct compound names and update the `Labels` rows.

## “Out of Volume”

Cause:

- All candidate source wells for a compound would drop below `SRC!G1` after planned transfers.
- Survey XML volumes overwrote planned volumes and found lower volumes.

Fix:

- Increase source volume, add another source well with the same compound and stock, reduce destination well volume, increase stock concentration, or split the experiment.
- Re-run after correcting actual pipetted volumes.

## Zero-volume or severe rounding warning

Cause:

- Desired transfer is below the Echo 2.5 nL increment.
- Stock is too concentrated for the requested final concentration and well volume.

Fix:

- Lower stock concentration, increase well volume, choose a higher final concentration, or accept/quantify rounding only when scientifically justified.

## LDV transfer above 500 nL

Cause:

- Requested final concentration is high relative to stock concentration.

Fix:

- Use a more concentrated stock, a different source plate type/method, or a generator/method version that explicitly splits large dispenses.

## Backfill is missing or applied to unexpected wells

Cause:

- `Backfill` appears as a case-insensitive token in the `Plate Map:`, so it is an ordinary treatment label rather than reserved configuration.
- The well has no positive normal transfer, or its rounded normal-transfer total already equals/exceeds the target.
- A Screen slot is unresolved on the partial final plate, or the physical well is blacklisted.
- An Assay design uses overlapping same-picklist sections; backfill is calculated independently in each section.

Fix:

- Keep exactly one `Backfill | vehicle | mock final` row in `Labels` and keep that label out of the map.
- Remember that `--backfill_volume` is the target total in nL, not a fixed vehicle amount.
- Sum the rounded normal transfers at the generated destination coordinates when auditing the remainder.
- Put native Assay backfill only in the intended balancing section or keep backfilled sections spatially disjoint.

## Backfill configuration error or out of volume

Cause:

- More than one unused Backfill row is present.
- The vehicle name does not exactly match a compound in the active `SRC` inventory.
- Mock final or stock is missing, nonpositive, or nonfinite; the derived target rounds to 0 nL.
- Only one of `--backfill_name` and `--backfill_volume` was supplied.
- The vehicle source would fall below the minimum source volume after top-ups.

Fix:

- Keep one reserved row, use consistent concentration units, and validate the derived target with `estimate-volume`.
- Supply both CLI arguments and use the compound label from SRC column B.
- Increase vehicle source volume or add equivalent source wells with the same stock concentration.
- Treat native backfill as operational: it is visible in the Echo picklist and source accounting, but intentionally absent from treatment merge/condition metadata.

## Screen library dispenses wrong plate order

Cause:

- `LIB` rows are not sorted by source plate barcode.
- A source plate barcode appears, disappears, then appears again.

Fix:

- Sort the shortlisted `LIB` by source plate barcode while preserving within-plate order if needed.

## Screen blacklisted controls disappear

Cause:

- Controls were left fixed/unmovable or the current generator does not relocate controls.

Fix:

- Prefer leaving controls fixed and accept rejected controls in the merge output.
- Add redundant controls in multiple locations.
- Only group controls if the method version and design intentionally support moving them.

## Blacklist relocation behaves unexpectedly

Cause:

- The `Groups:` map does not label all intended movable library wells.
- Controls were included in groups.
- A repeated replicate/dose group was cut by the surviving well capacity.

Fix:

- Rebuild the group map with only library wells in groups.
- Validate output replicate/dose cardinality after compaction.
- Do not rely on distinct `Groups:` text values for regional isolation in the current shift-based Python path; inspect the compacted output or separate the designs.

## Screen wells are not randomized

Cause:

- Runtime randomization defaults to off.
- The CLI boolean was omitted or false.
- A layout-label pool has only one genuinely used well, so its mapping is necessarily unchanged.

Fix:

- Run `python -m PickliPy.Screen -f design.xlsx --randomize-wells true --no-pause` (or use `-r true` / `--randomize_wells true`). The legacy positional design path remains valid only when `-f/--file` is omitted.
- Remember that only actual post-blacklist, LIB-resolved dispenses are candidates; empty final-plate slots do not become used.
- The seed is deterministic from the workbook basename and destination barcode, so rerunning the same design intentionally reproduces the same placement.

## “Well randomization failed” or replicates cannot be separated

Cause:

- A `Layout:` label pool is too small or too narrow for the replicate count.
- Too many identical resolved treatment signatures must fit without orthogonal adjacency.

Fix:

- Add more used positions carrying the same layout label, reduce the replicate count, or redesign the region geometry.
- Do not merge layout groups merely to suppress the error unless movement across that biological boundary is valid.

## Randomized wells cross an expected boundary

Cause:

- The boundary exists only in `DST_Blacklist!Groups:` and not in the DST `Layout:` map.
- Layout text differs by whitespace/case or is blank where a distinct region was intended.
- Identical layout labels were reused in disconnected regions; identical text always forms one pool.

Fix:

- Add or correct `DST!Layout:` labels for every intended randomization region.
- Treat `Layout:` and blacklist `Groups:` as separate maps and validate both when both features are active.

## Randomized metadata disagrees with the expected design positions

Cause:

- The original Excel `Plate Map:` is the pre-randomization template and is not rewritten.
- A downstream analysis used template coordinates instead of generated runtime metadata.

Fix:

- Use the destination wells in the generated Echo picklist, `table merge.txt`, `plate map compounds.txt`, and `plate map concentrations.txt`; these artifacts all carry the randomized coordinates.
- Keep blacklist `rejected` positions and the `Layout:` grid at their original physical coordinates; they are intentionally not remapped.

## Excel turns `@A` references into something unexpected

Cause:

- Manual Excel entry or formatting changed the `@ColumnLetter` string.

Fix:

- Format those cells as text before entry, or generate the workbook with the helper script.
- The actual cell value should be `@A`, not a formula and not a string with a literal leading apostrophe.

## Plate map shifted by one row/column

Cause:

- Rows or columns were inserted/cut near `Plate Map:` or `Concentrations Map:`.

Fix:

- Ensure the intended destination `A1` is one row below and one column right of the map label.
- Use the helper script to regenerate the section if in doubt.
