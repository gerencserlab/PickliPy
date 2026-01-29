# PickliPy QC

Post-execution validation tool for Echo acoustic dispenser. Parses print and survey XML files, matches compounds to library data, and generates smart rerun picklists for failed transfers.

## Features

- **Print XML Validation**: Analyze post-dispense results, identify skipped transfers, volume discrepancies
- **Survey XML Analysis**: Check source plate volumes against working ranges
- **Compound Lookup**: Automatically match wells to compounds from your picklist design file
- **Smart Rerun Picklists**: Generate recovery picklists with intelligent well-switching for failed transfers
- **Batch Processing**: Process multiple XML files at once with combined Excel output

## Installation

### Option 1: Standalone Executable (Windows)

Download `picklipy_qc.exe` from the [Releases](https://github.com/vkattunga/PickliPy-QC/releases) page. No installation required.

### Option 2: Run from Source

Requires Python 3.10+

```bash
# Clone the repository
git clone https://github.com/vkattunga/PickliPy-QC.git
cd PickliPy-QC

# Install dependencies
pip install -r requirements.txt

# Run
python picklipy_qc.py
```

## Usage

### 1. Configure Settings

| Setting | Description |
|---------|-------------|
| **XML Mode** | `Print XML` for post-dispense validation, `Survey XML` for volume checks |
| **Picklist Mode** | `Assay` or `Screening` - determines how compounds are looked up |
| **Log Directory** | Folder containing Echo XML log files |
| **Time Window** | Only show files modified within this many hours |
| **Picklist Design File** | The Excel file used to generate the original picklist (for compound lookup) |
| **Output Directory** | Where to save QC reports and rerun picklists |

### 2. Select Files

- Check the boxes next to the XML files you want to process
- Use **Select All** / **Deselect All** for batch operations
- Click anywhere on a file row to toggle selection

### 3. Process

- (Optional) Click **Load Compound Lookup** to verify compound matching before processing
- Click **▶ Process Selected Files** to run validation
- Monitor progress in the Status Log

### 4. Review Output

- **Excel Report**: `PickliPy_QC_print_YYYYMMDD_HHMMSS.xlsx` with one sheet per source→destination pair
- **Rerun Picklist**: `rerun_picklist_YYYYMMDD_HHMMSS.csv` (only generated if skipped transfers exist)

## Compound Lookup

The tool can match source wells to compound names using your picklist design file.

### Screening Mode

Reads from two sheets:
- **LIB sheet**: `Plate`, `RACKPOS`, `Compound` columns for library compounds
- **SRC sheet**: `Plate barcode`, `Source well`, `Compound` columns for controls

The `*` wildcard in SRC's `Plate barcode` column expands to all plates found in LIB.

### Assay Mode

Reads from:
- **SRC sheet**: `Plate barcode`, `Source well`, `Compound` columns
- **DST sheet**: Scans for `Barcode_SRC:` entries to expand `*` wildcards

## Smart Rerun Picklists

When transfers are skipped, the tool generates a rerun picklist with intelligent well-switching:

1. **Identifies failed transfers** from print XML
2. **Finds alternative wells** for each compound (from well_alternatives mapping)
3. **Checks volumes**:
   - Uses post-dispense volume from print XML for wells that were used
   - Uses pre-dispense volume from matched survey XML for unused alternative wells
4. **Selects best available well** with sufficient volume
5. **Logs all decisions** (switches, exclusions) in the Status Log

### Survey Matching

Survey files are automatically matched to print files by:
- **Barcode**: Must match the source plate barcode
- **Timestamp**: Survey must be taken BEFORE the print (most recent survey wins)

## Output Format

### Excel Report Sheets

Each sheet contains:
- **Metadata**: Source/destination barcodes, plate types, transfer date
- **Error Summary**: Count of each Echo error code (print mode only)
- **Transfer Table**: All transfers with volumes, flags for skipped/volume differences

### Rerun Picklist Format

Standard Echo picklist CSV format:
```
SRC_Barcode,SRC_Well,DST_Barcode,Volume_nL,DST_Well,Volume_nL,DST_Well,...
```

## Supported Plate Types

| Plate Type | Dead Volume (µL) | Max Volume (µL) |
|------------|------------------|-----------------|
| 384LDV_DMSO | 2.5 | 12.0 |
| 384LDV_DMSO2 | 2.5 | 12.0 |
| 384LDV_AQ_B2 | 3.0 | 12.0 |
| 384LDV_AQ_P2 | 6.0 | 16.0 |
| 384PP_DMSO | 15.0 | 65.0 |
| 384PP_DMSO2 | 15.0 | 65.0 |
| 384PP_AQ_SP2 | 15.0 | 65.0 |
| 384PP_AQ_CP | 20.0 | 50.0 |

To add more plate types, edit the `WORKING_RANGE_LOOKUP` dictionary in the source code.

## Configuration

Settings are automatically saved to `~/.picklipy_qc_config.json` and restored on next launch.

## Requirements

- Python 3.10+ (if running from source)
- pandas
- openpyxl
- xlsxwriter

## Author

Varunya Kattunga
