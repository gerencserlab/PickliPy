# PickliPy

PickliPy is a small set of command-line tools that generate **Echo 650** picklists (CSV) and the companion **Plate:Works** helper files (inventory + process lists) from Excel-based experimental designs.

It supports three workflows:

- **Assay** – flexible, user-defined plate maps for tool compounds, combinatorics, dose response, and serial additions.
- **Screen** – screening / reformatting where a **library table (LIB)** is threaded into a single destination plate template across as many destination plates as needed.
- **Bluetable** – a “blue-table” workflow where **each destination plate is defined by its own XLSX file**; all destination files in a folder are merged into a single picklist.

> These tools are written for internal Echo + Plate:Works workflows. They perform extensive input validation (missing labels, missing barcodes, out-of-volume checks, etc.) and will stop with a clear error if something is inconsistent.

---

## Installation

### 1) Requirements

- Python **3.10+** recommended
- Microsoft Excel files as input (`.xlsx`)

### 2) Create a virtual environment (recommended)

**Windows (PowerShell):**

```powershell
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

**macOS / Linux (bash/zsh):**

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
```

---

## Quick usage

Run the operation that matches your workbook/workflow:

### Assay

```bash
python -m PickliPy.Assay /path/to/picklist-assay.xlsx
```

### Screen

```bash
python -m PickliPy.Screen /path/to/picklist-screening.xlsx
```

### Bluetable

```bash
python -m PickliPy.Bluetable PICKLIST_NAME /path/to/inventory.xlsx /path/to/folderdst
```

Where:

- `PICKLIST_NAME` is the base name used for output files.
- `inventory.xlsx` is a SRC-style inventory workbook (same structure as Assay SRC).
- `folderdst` is a folder containing one or more destination “blue-table” `.xlsx` files.

---

## What each operation does

## 1) Assay workflow

### Purpose

Use **Assay** when you want *full control* over what goes into each destination well:

- one-to-many dispensing from a tool plate into one or more destination plates
- combinatoric additions (multiple labels per well)
- optional **dose-response mode** using a Concentrations Map
- serial additions (repeat plate maps with the same picklist name to append dispenses)

### Inputs

A single Excel workbook (typically a template like `picklist-assay.xlsx`) with **two worksheets**:

- **SRC**: source inventory (what is in the Echo source plate)
- **DST**: destination design and run configuration

#### SRC sheet (inventory)

Header (row 1) includes a **minimum allowed source well volume** (typically in cell `G1`).

Columns are typically:

1. Source Well (e.g. `A1`)
2. Compound Name
3. Stock concentration
4. Unit (informational; should match DST units)
5. Volume in source plate (µL)
6. Plate barcode

Rules (high level):

- `Plate barcode + Source Well` must be unique per row.
- The same compound may appear in multiple wells (alternative wells), enabling automatic well switching when a well reaches the minimum volume.
- Source well volumes are tracked so the tool will not dispense below the minimum volume.

#### DST sheet (design + configuration)

The top of the sheet contains label/value configuration (inventory/process filenames, rack numbers, well volume, etc.).

Then one or more **dispense blocks** follow. Each block contains:

- `Plate Map:` (a 24×16 grid in 384-well layout; the top-left region can be used for 96-well plates)
- Optional `Concentrations Map:` (same size grid) to activate dose-response mode for that block
- `Picklist Name:` (base filename for the Echo picklist)
- `Barcode_SRC:` (one source plate barcode for that block)
- `Barcode_DST:` (one or more destination barcodes, or `*` to target all destination barcodes declared in the workbook)
- `Labels` table (mapping plate-map labels to SRC compound names and final concentrations)

### Outputs

Assay produces, in the same folder as the input workbook (unless your module is configured otherwise):

- `PicklistName.csv` (Echo picklist)
- One or two inventory files (depending on whether SRC and DST inventory filenames are the same)
- `Process_SRC.txt`, `Process_DST.txt`
- A “after dispense” workbook updating remaining source volumes (name varies by module version)
- Optional merge/conditions tables for downstream analysis (labels/compounds)

---

## 2) Screen workflow

### Purpose

Use **Screen** when you want to apply a **single destination plate layout template** across an entire compound library:

- compounds from a multi-plate library are automatically assigned to “slots” (#1, #2, …) in the template
- generates as many destination plates as needed to cover the entire library
- supports controls, replicates, and per-compound concentrations
- optional quality-control *blacklist* workflow to skip or relocate dispensing away from bad wells

### Inputs

A single Excel workbook (typically a template like `picklist-screening-*.xlsx`) with:

- **SRC** – controls + a row defining how to pull compound name / well / barcode from **LIB**
- **DST** – the destination template plate map and run configuration
- **LIB** – the library table (one row per compound entry)
- **DST_Blacklist** *(optional)* – per-plate well exclusions + relocation groups

Important constraints:

- The **LIB** table must be ordered by **source plate barcode** (Plate:Works limitation: once a source plate is completed, it cannot be revisited later in the run).
- The entire library table is dispensed; to dispense a subset, remove rows from LIB.

### Outputs

Screen produces, next to the input workbook:

- `PicklistName.csv` (Echo picklist)
- Inventory and process files (`Inventory_*.csv`, `Process_*.txt`)
- Loop / parameter files used by the Plate:Works screening protocol (module-dependent naming)
- “merge table” and destination plate maps for compounds and concentrations (for downstream analysis)
- Summary counts printed to the console (destination plate count, library plate count, total dispense events)

---

## 3) Bluetable workflow

### Purpose

Use **Bluetable** when destination plate layouts are stored as **separate XLSX files**, one per destination plate, and you want to merge them into a single Echo picklist.

Key characteristics:

- **Single source plate** (inventory must contain exactly one unique source barcode)
- **Many destination plates** (one per `.xlsx` in `folderdst`)
- Destination barcodes are taken from the **destination filenames** (the filename stem)
- Allows duplicate compound labels in the SRC inventory at **different stock concentrations**; the selected source well must match the concentration requested by the blue-table destination file.

### Inputs

1) `inventory.xlsx`

- A SRC-style inventory sheet (same columns as Assay SRC)
- Must contain **exactly one** distinct plate barcode (Bluetable uses one source plate)

2) `folderdst/`

- Contains one or more destination blue-table `.xlsx` files
- Each `.xlsx` defines a single destination plate in its **first worksheet**
- `folderdst` may also contain one or more `*_platesurvey.xml` files; if present, these may be used to update the tracked source volumes before dispensing

### Outputs

Bluetable writes all outputs into:

```
<folderdst>/outs/
```

Typical outputs:

- `PICKLIST_NAME_Picklist.csv`
- `PICKLIST_NAME_Inventory.csv`
- `PICKLIST_NAME_Process_SRC.txt`
- `PICKLIST_NAME_Process_DST.txt`
- `after_dispense_*.xlsx` (updated inventory / remaining volumes)
- Per-destination `*_notebook_tables.xlsx` (human-readable Echo/manual preparation tables)

---

## Plate:Works integration notes

The generated files need to be placed where Plate:Works expects them. Typical paths in existing workflows:

- Picklist CSV: `C:\Work` (referenced by the Echo:PICK step)
- Inventory CSV: `C:\Plateworks_6_30\Provider\Run\Inventory`
- Process TXT: `C:\Plateworks_6_30\Provider\Run\Threads\process List`
- Screening parameter CSV (if used by your protocol): `C:\Plateworks_6_30\GUI`

Close any open CSV/TXT files before running or simulating an assay in Plate:Works—open files can cause confusing “device errors”.

---

## Troubleshooting

- **“Missing label definition”**: a label used in the plate map is not defined in the Labels table.
- **“Out of Volume”**: dispensing would deplete a source well below the minimum volume; refill the source plate or add an alternative well in SRC.
- **Rounding warnings**: Echo volumes are rounded to 2.5 nL; very small requested dispenses may be inaccurate.
- **Excel lock files**: if you see `~$...xlsx` files, close Excel and re-run.

---

## License / attribution

Internal use. If you need to redistribute this package or publish derived work, add an explicit license and attribution policy here.
