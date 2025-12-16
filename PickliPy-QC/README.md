# PickliPy QC

Post-execution validation tool for Echo acoustic dispenser. Parses print and survey XML files and matches to compound library data.

## Features

- **Print XML Mode**: Validates post-dispense transfers, identifies skipped wells, matches compounds to library
- **Survey XML Mode**: Checks well volumes against working ranges, flags out-of-range wells
- **Rerun Picklist Generation**: Automatically creates a picklist CSV containing only skipped transfers for easy re-dispensing
- **Auto-discovery**: Scans Echo log directory for recent XML files (configurable time window)
- **Barcode display**: Shows source/destination barcodes extracted from XML files
- **Persistent settings**: Remembers library file path, log directory, and time window between sessions

## Requirements

- Python 3.9+
- pandas
- xlsxwriter
- tkinter (included with standard Python installation)

## Installation

### Option 1: Standalone Executable (Recommended)

Download `picklipy_qc.exe` from the [Releases](https://github.com/vkattunga/picklipy-qc/releases) page. No Python installation required.

### Option 2: Python Script

1. Install dependencies:

```bash
   pip install pandas xlsxwriter
```

2. Run the script:

```bash
   python picklipy_qc.py
```

## Usage

1. **Select Mode**: Choose "Print XML" for post-dispense validation or "Survey XML" for volume checks

2. **Set Log Directory**: Point to your Echo log folder (default: `C:\Labcyte\Echo\Client\Logs`)

3. **Adjust Time Window**: Set how many hours back to search for XML files (default: 12 hours)

4. **Select Library File**: Browse to your compound library Excel file
   - Must have sheets named by plate barcode
   - Sheets must contain 'RACKPOS' (or 'Well') and 'Compound' columns

5. **Select Files**: Click on XML files in the list (Ctrl+click for multiple, or use "Select All")

6. **Process**: Click "Process Selected Files" to generate QC reports

## Library File Format

The compound library must be an Excel file (`.xlsx`) with the following structure:

### Sheet Naming
Each sheet must be named **exactly** matching the plate barcode. For example:
- Plate barcode `M1001` → Sheet name `M1001`
- Plate barcode `M2001` → Sheet name `M2001`

### Required Columns

| Column | Description | Example |
|--------|-------------|---------|
| `RACKPOS` | Well position (A1 format) | A1, B3, P24 |
| `Compound` | Compound name or identifier | Rotenone, DMSO, CMPD-001234 |

### Optional Columns

| Column | Description |
|--------|-------------|
| `Plate Barcode` | Used for validation (confirms sheet matches barcode) |
| Any other columns | Ignored by PickliPy QC but preserved for your reference |

### Example Sheet

A sheet named `M1001` might look like:

| Plate Barcode | RACKPOS | Compound | Concentration | Notes |
|---------------|---------|----------|---------------|-------|
| M1001 | B2 | Oligomycin | 10mM | ATP synthase inhibitor |
| ... | ... | ... | ... | ... |

### Multiple Plates

For libraries spanning multiple plates, include one sheet per plate:
```
MyLibrary.xlsx
├── Sheet: M1001 (compounds 1-384)
├── Sheet: M1002 (compounds 385-768)
├── Sheet: M1003 (compounds 769-1152)
└── ...
```

### Common Issues

- **Sheet not found**: The barcode in the XML must exactly match a sheet name (case-sensitive)
- **Column not found**: Check for typos — column must be `RACKPOS` not `RackPos` or `Rack Position`
- **Empty compounds**: Wells without entries in the library file will show `NA` in the output

## Output

### Print XML Output

Creates an Excel file with:
- Metadata (source/destination barcodes, plate types, transfer date)
- Error summary (if any wells were skipped)
- Transfer table with:
  - Source well and compound name
  - Destination well
  - Pre/post transfer source volumes
  - Target vs actual volumes dispensed
  - Skip flags and reasons
- Conditional formatting: red highlighting for skipped transfers, yellow for volume discrepancies

**Rerun Picklist**: If skipped transfers are detected, a separate CSV file is generated containing only the failed transfers in Echo picklist format, ready for re-dispensing.

### Survey XML Output

Creates an Excel file with:
- Metadata (plate barcode, type, working volume range)
- Well table with:
  - Well position and compound name
  - Measured volume
  - Volume flags (below working range, exceeds range, missing)

## Adding New Plate Types

Edit the `WORKING_RANGE_LOOKUP` dictionary in the script:

```python
WORKING_RANGE_LOOKUP = {
    '384LDV_DMSO':  {'dead_volume': 2.5,  'max': 12.0},
    '384PP_DMSO':   {'dead_volume': 15.0, 'max': 65.0},
    'YOUR_PLATE':   {'dead_volume': X.X,  'max': XX.X},
    # Add more as needed
}
```

## Configuration

Settings are saved to `~/.picklipy_qc_config.json` and persist between sessions:
- Log directory path
- Library file path  
- Output directory path
- Time window (hours)

## Troubleshooting

| Error | Solution |
|-------|----------|
| "No sheet found for source barcode" | Library Excel file needs a sheet named exactly matching the plate barcode in the XML file |
| "Missing expected columns" | Library sheets must contain 'RACKPOS' (or 'Well') and 'Compound' columns |
| No files showing | Check that the log directory is correct and XML files exist within the time window |


## Author

Varunya Kattunga
