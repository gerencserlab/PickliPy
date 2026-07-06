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
- The final `Labels` row is not a suitable fill control.

Fix:

- Rebuild the group map with only library wells in groups.
- Use separate group labels for separate cell populations/regions.
- Make the last `Labels` row a benign fill treatment.

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
