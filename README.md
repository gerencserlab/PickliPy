# PicklyPy

PicklyPy generates Echo-style picklists (well-to-well transfers + volumes) from human-readable 384‑well plate maps stored in an Excel design file. It also produces Plate:Works inventory/process files plus metadata tables for downstream analysis.

This package is a Python translation of the original Wolfram Language scripts:

- `assaypicklist.wls` → `PicklyPy.Assay`
- `screeningpicklist.wls` → `PicklyPy.Screen`

## Modules

- **PicklyPy.Assay**: tool library / combinatorics / dose-response / multi-additions using multiple plate maps.
- **PicklyPy.Screen**: screening + reformatting workflows where a multi‑plate library is automatically assigned to assay (destination) plates from a single generic layout; supports optional well blacklisting.

## Inputs

All workflows use an Excel design file with worksheets:

- `SRC`: source inventory (compound identifiers, stock concentrations, volumes, plate barcodes)
- `DST`: destination layout + label lists
- `LIB` (Screen only): library table used by `@<column>` placeholders
- optional blacklist sheet (Screen only): `Groups Map:` and `Well Blacklist:` sections

If one or more `*_platesurvey.xml` files exist in the same folder as the design file, they are used to update source well volumes.

## Usage

```bash
# Assay workflow
python -m PicklyPy.Assay /path/to/design.xlsx

# Screening workflow
python -m PicklyPy.Screen /path/to/design.xlsx

# Disable the end-of-run "press Enter" pause
python -m PicklyPy.Assay /path/to/design.xlsx --no-pause
```

Outputs are written next to the input Excel file.

## Dependency

- `openpyxl`

