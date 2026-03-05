"""
PickliPy QC - Post-execution validation tool for Echo acoustic dispenser
Parses print and survey XML files and matches to compound library data.

Version 2.1 - Fixed duplicate sheet name error for multi-dispense workflows
              Added Run Summary sheet for auditability

Author: Varunya Kattunga
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xml.etree.ElementTree as ET
import pandas as pd
import os
import json
from datetime import datetime, timedelta
from pathlib import Path
from collections import Counter, defaultdict

# =============================================================================
# CONFIGURATION
# =============================================================================

CONFIG_FILE = Path.home() / '.picklipy_qc_config.json'
DEFAULT_LOG_DIR = r"C:\Users\vkattunga\Workspace\Echo_QC_Scripts\Echo_XML_Script\v3_QC_script\v3_testfiles"
DEFAULT_TIME_WINDOW = 12  # hours

# Error code descriptions for print XML parsing
ERROR_CODE_MAP = {
    '0201001': "Invalid well survey data. Unable to determine fluid properties.",
    '0201002': "Meniscus not found",
    '0201003': "Excessive fluid height change",
    '0201004': "Inconsistent fluid level",
    '0202001': "Empty well",
    '0202002': "Overfilled well",
    '0202003': "Excessive fluid height change",
    '0202004': "Inconsistent fluid level",
    '0202005': "Problem calculating fluid ejection parameters",
    '0202006': "Problem calculating fluid ejection parameters",
    '0202007': "Problem calculating fluid ejection parameters",
    '0202008': "Problem calculating fluid ejection parameters",
    '0203001': "Unsuccessful Realtime power adjustment. Check compatibility of fluid with labware and calibration.",
}

# Working volume ranges by plate type
WORKING_RANGE_LOOKUP = {
    '384LDV_DMSO':    {'dead_volume': 2.5,  'max': 12.0},
    '384LDV_Dest':    {'dead_volume': 2.5,  'max': 12.0},
    '384LDV_DMSO2':   {'dead_volume': 2.5,  'max': 12.0},
    '384LDV_AQ_P2':   {'dead_volume': 6.0,  'max': 16.0},
    '384LDV_AQ_B2':   {'dead_volume': 3.0,  'max': 12.0},
    '384PP_DMSO':     {'dead_volume': 15.0, 'max': 65.0},
    '384PP_DMSO2':    {'dead_volume': 15.0, 'max': 65.0},
    '384PP_AQ_SP2':   {'dead_volume': 15.0, 'max': 65.0},
    '384PP_AQ_CP':    {'dead_volume': 20.0, 'max': 50.0},
}


# =============================================================================
# CONFIG MANAGEMENT
# =============================================================================

def load_config():
    """Load saved configuration from JSON file."""
    default_config = {
        'log_directory': DEFAULT_LOG_DIR,
        'picklist_file': '',
        'output_directory': '',
        'time_window_hours': DEFAULT_TIME_WINDOW,
        'picklist_mode': 'screening',
    }

    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, 'r') as f:
                saved = json.load(f)
                default_config.update(saved)
        except (json.JSONDecodeError, IOError):
            pass

    return default_config


def save_config(config):
    """Save configuration to JSON file."""
    try:
        with open(CONFIG_FILE, 'w') as f:
            json.dump(config, f, indent=2)
    except IOError as e:
        print(f"Warning: Could not save config: {e}")


# =============================================================================
# COMPOUND LOOKUP BUILDER
# =============================================================================

def get_column_case_insensitive(df, target_name):
    """Find column name case-insensitively."""
    for col in df.columns:
        if not isinstance(col, str):
            continue
        if col.lower() == target_name.lower():
            return col
    return None


def build_compound_lookup(picklist_file, picklist_mode, log_func=None):
    """
    Build compound lookup dictionary from picklist design file.

    Args:
        picklist_file: Path to the XLSX file used to generate the picklist
        picklist_mode: 'assay' or 'screening'
        log_func: Optional function to log messages

    Returns:
        tuple: (
            compound_lookup: dict of (barcode, well) -> compound,
            well_alternatives: dict of (barcode, compound) -> [list of wells],
            dead_volume_ul: float
        )
    """
    def log(msg):
        if log_func:
            log_func(msg)
        else:
            print(msg)

    if not picklist_file or not os.path.exists(picklist_file):
        log(" No picklist design file provided - compound lookup disabled")
        return {}, {}, 2.5

    log(f"Loading picklist design file: {os.path.basename(picklist_file)}")
    log(f"  Mode: {picklist_mode}")

    try:
        xls = pd.ExcelFile(picklist_file)
        log(f"  Available sheets: {xls.sheet_names}")
    except Exception as e:
        log(f" Could not open picklist file: {e}")
        return {}, {}, 2.5

    compound_lookup = {}
    well_alternatives = defaultdict(list)
    dead_volume_ul = 2.5
    all_barcodes_for_expansion = set()

    # =================================================================
    # STEP 1: For screening mode, load LIB sheet FIRST to get barcodes
    # =================================================================
    if picklist_mode == 'screening':
        lib_sheet_name = None
        for name in xls.sheet_names:
            if name.lower() == 'lib':
                lib_sheet_name = name
                break

        if lib_sheet_name:
            lib_df = pd.read_excel(xls, sheet_name=lib_sheet_name)
            log(f"  Found LIB sheet with {len(lib_df)} rows")
            log(f"  LIB columns found: {[str(c) for c in lib_df.columns if isinstance(c, str)]}")

            lib_plate_col = get_column_case_insensitive(lib_df, 'src_barcode')
            lib_well_col = get_column_case_insensitive(lib_df, 'rackpos')
            lib_compound_col = get_column_case_insensitive(lib_df, 'compound')

            if lib_plate_col:
                all_barcodes_for_expansion = set(lib_df[lib_plate_col].dropna().unique())
                log(f"  Found {len(all_barcodes_for_expansion)} unique source plates in LIB")

            if all([lib_plate_col, lib_well_col, lib_compound_col]):
                log(f"  LIB mapping: plate='{lib_plate_col}', well='{lib_well_col}', compound='{lib_compound_col}'")
                lib_count = 0
                for _, row in lib_df.iterrows():
                    barcode = str(row[lib_plate_col]).strip() if pd.notna(row[lib_plate_col]) else ''
                    well = str(row[lib_well_col]).strip() if pd.notna(row[lib_well_col]) else ''
                    compound = str(row[lib_compound_col]).strip() if pd.notna(row[lib_compound_col]) else ''

                    if not barcode or not well or not compound:
                        continue

                    compound_lookup[(barcode, well)] = compound
                    if well not in well_alternatives[(barcode, compound)]:
                        well_alternatives[(barcode, compound)].append(well)
                    lib_count += 1

                log(f"  Loaded {lib_count} compound entries from LIB sheet")
            else:
                missing = []
                if not lib_plate_col: missing.append("'SRC_Barcode'")
                if not lib_well_col: missing.append("'RACKPOS'")
                if not lib_compound_col: missing.append("'Compound'")
                log(f"  LIB sheet missing columns: {', '.join(missing)}")
        else:
            log("  No LIB sheet found in picklist file")

    # =================================================================
    # STEP 2: For assay mode, get barcodes from DST sheet
    # =================================================================
    if picklist_mode == 'assay':
        dst_sheet_name = None
        for name in xls.sheet_names:
            if name.lower() == 'dst':
                dst_sheet_name = name
                break

        if dst_sheet_name:
            dst_df = pd.read_excel(xls, sheet_name=dst_sheet_name, header=None)
            for idx, row in dst_df.iterrows():
                if len(row) >= 2:
                    cell_a = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
                    if cell_a.lower() == 'barcode_src:':
                        cell_b = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ''
                        if cell_b and cell_b != '*':
                            all_barcodes_for_expansion.add(cell_b)
            log(f"  Found {len(all_barcodes_for_expansion)} source barcodes in DST sheet")

    # =================================================================
    # STEP 3: Load SRC sheet (controls and explicit compounds)
    # =================================================================
    src_sheet_name = None
    for name in xls.sheet_names:
        if name.lower() == 'src':
            src_sheet_name = name
            break

    if src_sheet_name:
        src_df = pd.read_excel(xls, sheet_name=src_sheet_name)
        log(f"  Found SRC sheet with {len(src_df)} rows")
        log(f"  SRC columns found: {[str(c) for c in src_df.columns if isinstance(c, str)]}")

        barcode_col = get_column_case_insensitive(src_df, 'plate barcode')
        well_col = get_column_case_insensitive(src_df, 'source well')
        compound_col = get_column_case_insensitive(src_df, 'compound name')

        if all([barcode_col, well_col, compound_col]):
            log(f"  SRC mapping: barcode='{barcode_col}', well='{well_col}', compound='{compound_col}'")
            src_count = 0
            src_wildcard_count = 0

            for _, row in src_df.iterrows():
                barcode = str(row[barcode_col]).strip() if pd.notna(row[barcode_col]) else ''
                well = str(row[well_col]).strip() if pd.notna(row[well_col]) else ''
                compound = str(row[compound_col]).strip() if pd.notna(row[compound_col]) else ''

                if not well or not compound:
                    continue

                if barcode == '*':
                    barcodes_to_add = all_barcodes_for_expansion
                    src_wildcard_count += 1
                else:
                    barcodes_to_add = {barcode} if barcode else set()

                for bc in barcodes_to_add:
                    compound_lookup[(bc, well)] = compound
                    if well not in well_alternatives[(bc, compound)]:
                        well_alternatives[(bc, compound)].append(well)
                    src_count += 1

            log(f" Loaded {src_count} compound entries from SRC sheet ({src_wildcard_count} with * wildcard)")
        else:
            missing = []
            if not barcode_col: missing.append("'Plate barcode'")
            if not well_col: missing.append("'Source well'")
            if not compound_col: missing.append("'Compound Name'")
            log(f"  SRC sheet missing columns: {', '.join(missing)}")
    else:
        log("  No SRC sheet found in picklist file")

    log(f" Compound lookup ready: {len(compound_lookup)} total entries")
    return dict(compound_lookup), dict(well_alternatives), dead_volume_ul


# =============================================================================
# SURVEY FILE MATCHING
# =============================================================================

def get_survey_timestamp(xml_file):
    """Extract timestamp from survey XML file."""
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()
        date_str = root.attrib.get('date', '')
        if date_str:
            for fmt in ['%m/%d/%Y %H:%M:%S', '%Y-%m-%d %H:%M:%S', '%m/%d/%Y']:
                try:
                    return datetime.strptime(date_str, fmt)
                except ValueError:
                    continue
        return datetime.fromtimestamp(Path(xml_file).stat().st_mtime)
    except Exception:
        return datetime.fromtimestamp(Path(xml_file).stat().st_mtime)


def get_print_timestamp(xml_file):
    """Extract timestamp from print XML file."""
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()
        date_str = root.attrib.get('date', '')
        if date_str:
            for fmt in ['%m/%d/%Y %H:%M:%S', '%Y-%m-%d %H:%M:%S', '%m/%d/%Y']:
                try:
                    return datetime.strptime(date_str, fmt)
                except ValueError:
                    continue
        return datetime.fromtimestamp(Path(xml_file).stat().st_mtime)
    except Exception:
        return datetime.fromtimestamp(Path(xml_file).stat().st_mtime)


def get_survey_barcode(xml_file):
    """Extract barcode from survey XML file."""
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()
        return root.attrib.get('barcode', '')
    except Exception:
        return ''


def find_matching_surveys(log_dir, print_files, log_func=None):
    """
    Find matching survey files for each print file based on barcode and timestamp.

    Args:
        log_dir: Directory containing XML files
        print_files: List of print file paths
        log_func: Optional logging function

    Returns:
        dict: print_file_path -> survey_file_path (or None if no match)
    """
    def log(msg):
        if log_func:
            log_func(msg)
        else:
            print(msg)

    log_path = Path(log_dir)

    surveys = []
    for survey_file in log_path.glob('*_platesurvey.xml'):
        barcode = get_survey_barcode(survey_file)
        timestamp = get_survey_timestamp(survey_file)
        surveys.append({
            'path': str(survey_file),
            'barcode': barcode,
            'timestamp': timestamp
        })

    log(f"Found {len(surveys)} survey files in directory")

    matches = {}
    for print_file in print_files:
        try:
            tree = ET.parse(print_file)
            root = tree.getroot()
            plate_info = root.find('plateInfo')
            src_plate = plate_info.find('plate[@type="source"]')
            src_barcode = src_plate.attrib.get('barcode', '') if src_plate is not None else ''
        except Exception:
            src_barcode = ''

        print_timestamp = get_print_timestamp(print_file)

        candidates = [
            s for s in surveys
            if s['barcode'] == src_barcode and s['timestamp'] <= print_timestamp
        ]

        if candidates:
            candidates.sort(key=lambda x: x['timestamp'], reverse=True)
            best_match = candidates[0]
            matches[print_file] = best_match['path']
            log(f"  {Path(print_file).name} → {Path(best_match['path']).name}")
        else:
            matches[print_file] = None
            log(f"  {Path(print_file).name} → No matching survey found")

    return matches


def load_survey_volumes(survey_file):
    """
    Load volumes from a survey XML file.

    Returns:
        dict: well -> volume_ul
    """
    if not survey_file or not os.path.exists(survey_file):
        return {}

    volumes = {}
    try:
        tree = ET.parse(survey_file)
        root = tree.getroot()
        for w in root.findall('w'):
            well = w.get('n')
            try:
                volume = float(w.get('vl'))
                volumes[well] = volume
            except (TypeError, ValueError):
                pass
    except Exception:
        pass

    return volumes


# =============================================================================
# XML FILE DISCOVERY
# =============================================================================

def get_xml_files(log_dir, mode, hours):
    """
    Find XML files in log directory within time window.

    Args:
        log_dir: Path to Echo log directory
        mode: 'print' or 'survey'
        hours: Number of hours to look back

    Returns:
        List of dicts with file info: {path, filename, modified, barcodes}
    """
    files = []
    cutoff_time = datetime.now() - timedelta(hours=hours)

    suffix = '_print.xml' if mode == 'print' else '_platesurvey.xml'

    log_path = Path(log_dir)
    if not log_path.exists():
        return files

    for filepath in log_path.glob(f'*{suffix}'):
        try:
            mtime = datetime.fromtimestamp(filepath.stat().st_mtime)
            if mtime < cutoff_time:
                continue

            barcodes = extract_barcodes(filepath, mode)

            files.append({
                'path': str(filepath),
                'filename': filepath.name,
                'modified': mtime,
                'barcodes': barcodes,
            })
        except Exception as e:
            print(f"Warning: Could not process {filepath}: {e}")

    files.sort(key=lambda x: x['modified'], reverse=True)
    return files


def extract_barcodes(filepath, mode):
    """Extract plate barcodes from XML file."""
    try:
        tree = ET.parse(filepath)
        root = tree.getroot()

        if mode == 'print':
            plate_info = root.find('plateInfo')
            if plate_info is not None:
                src = plate_info.find('plate[@type="source"]')
                dst = plate_info.find('plate[@type="destination"]')
                src_bc = src.get('barcode', '?') if src is not None else '?'
                dst_bc = dst.get('barcode', '?') if dst is not None else '?'
                return f"{src_bc} → {dst_bc}"
        else:
            barcode = root.attrib.get('barcode', '?')
            return barcode

    except Exception:
        return '?'

    return '?'


# =============================================================================
# PRINT XML PARSER
# =============================================================================

def parse_printmap(xml_file, compound_lookup=None):
    """
    Parse print XML file and match to compound library.

    Args:
        xml_file: Path to _print.xml file
        compound_lookup: Dict of (barcode, well) -> compound

    Returns:
        dict with summary info
    """
    tree = ET.parse(xml_file)
    root = tree.getroot()

    transfer_date = root.attrib.get('date', 'NA')

    plate_info = root.find('plateInfo')
    src_plate = plate_info.find('plate[@type="source"]')
    dst_plate = plate_info.find('plate[@type="destination"]')
    src_plate_type = src_plate.get('name')
    dst_plate_type = dst_plate.get('name')

    src_barcode = src_plate.attrib.get('barcode', 'NA')
    dst_barcode = dst_plate.attrib.get('barcode', 'NA')

    working_range = WORKING_RANGE_LOOKUP.get(src_plate_type, {})
    src_dead_vol = working_range.get('dead_volume', None)
    src_max_vol = working_range.get('max', None)

    # Parse successful transfers
    transfer_rows = []
    for w in root.find('printmap').findall('w'):
        pre_src_vol = float(w.get('vl', 0))
        post_src_vol = float(w.get('cvl', 0))
        target_vol = float(w.get('vt', 0)) / 1000
        actual_vol = float(w.get('avt', 0)) / 1000

        transfer_rows.append({
            'Well': w.get('n'),
            'SRC_Vol_Before': pre_src_vol,
            'SRC_Vol_After': post_src_vol,
            'Destination_Well': w.get('dn'),
            'Vol_Target': target_vol,
            'Vol_Actual': actual_vol,
            'Reason_Skipped': None,
            'Skipped': False,
            'ViolatesDeadVol': False,
        })

    # Parse skipped transfers
    skipped_rows = []
    error_summary = Counter()

    skipped_block = root.find('skippedwells')
    if skipped_block is not None:
        for w in skipped_block.findall('w'):
            reason_str = w.get('reason', '')
            if reason_str.startswith('MM') and ':' in reason_str:
                code = reason_str.split(':')[0].replace('MM', '')
                explanation = ERROR_CODE_MAP.get(code, "Unknown error")
                full_label = f"{code}: {explanation}"
                error_summary[full_label] += 1

            pre_src_vol = float(w.get('vl', 0))
            post_src_vol = float(w.get('cvl', 0))
            target_vol = float(w.get('vt', 0)) / 1000
            actual_vol = float(w.get('avt', 0)) / 1000

            violates_dead_vol = (
                src_dead_vol is not None and
                ((pre_src_vol < src_dead_vol) or (pre_src_vol - target_vol < src_dead_vol))
            )

            skipped_rows.append({
                'Well': w.get('n'),
                'SRC_Vol_Before': pre_src_vol,
                'SRC_Vol_After': post_src_vol,
                'Destination_Well': w.get('dn'),
                'Vol_Target': target_vol,
                'Vol_Actual': actual_vol,
                'Reason_Skipped': w.get('reason'),
                'Skipped': True,
                'ViolatesDeadVol': violates_dead_vol,
            })

    # Combine all rows
    all_rows = transfer_rows + skipped_rows
    df = pd.DataFrame(all_rows)

    # Match compounds using lookup
    if compound_lookup:
        df['Compound'] = df['Well'].apply(
            lambda w: compound_lookup.get((src_barcode, w), 'NA')
        )

    # Add barcode columns
    df['SRC_Barcode'] = src_barcode
    df['DST_Barcode'] = dst_barcode

    # Calculate Vol_Difference
    df['Vol_Difference'] = df['Vol_Actual'] - df['Vol_Target']

    # Conditional display: blank unless meaningful
    df['Skipped'] = df['Skipped'].apply(lambda x: 'Yes' if x else '')
    df['Reason_Skipped'] = df['Reason_Skipped'].fillna('')
    df['ViolatesDeadVol'] = df['ViolatesDeadVol'].apply(lambda x: 'Yes' if x else '')
    df['Vol_Difference'] = df['Vol_Difference'].apply(lambda x: x if x != 0 else '')

    # Build column list
    final_columns = ['SRC_Barcode', 'Well']
    if 'Compound' in df.columns:
        final_columns.append('Compound')
    final_columns.extend([
        'DST_Barcode', 'Destination_Well',
        'Vol_Target', 'Vol_Actual', 'Vol_Difference',
        'SRC_Vol_Before', 'SRC_Vol_After',
        'Skipped', 'Reason_Skipped', 'ViolatesDeadVol'
    ])
    df_final = df[final_columns]

    metadata_dict = {
        'SRC_Barcode': src_barcode,
        'SRC_PlateType': src_plate_type,
        'SRC_DeadVolume': src_dead_vol,
        'SRC_MaxVolume': src_max_vol,
        'DST_Barcode': dst_barcode,
        'DST_PlateType': dst_plate_type,
        'Transfer_Date': transfer_date
    }

    return {
        'df': df_final,
        'metadata': metadata_dict,
        'error_summary': error_summary,
        'src_barcode': src_barcode,
        'summary': {
            'src_barcode': src_barcode,
            'dst_barcode': dst_barcode,
            'total_transfers': len(transfer_rows),
            'skipped_transfers': len(skipped_rows),
            'unmatched_compounds': (df_final['Compound'] == 'NA').sum() if 'Compound' in df_final.columns else 0,
        }
    }


# =============================================================================
# SURVEY XML PARSER
# =============================================================================

def parse_platesurvey(xml_file, compound_lookup=None):
    """
    Parse survey XML file and match to compound library.

    Args:
        xml_file: Path to _platesurvey.xml file
        compound_lookup: Dict of (barcode, well) -> compound

    Returns:
        dict with summary info
    """
    tree = ET.parse(xml_file)
    root = tree.getroot()

    plate_barcode = root.attrib.get('barcode', 'NA')
    plate_date = root.attrib.get('date', 'NA')
    plate_type = root.attrib.get('name', 'NA')

    range_info = WORKING_RANGE_LOOKUP.get(plate_type)
    if not range_info:
        raise ValueError(f"Unknown plate type '{plate_type}'. Add it to WORKING_RANGE_LOOKUP.")

    min_vol = range_info['dead_volume']
    max_vol = range_info['max']

    data = []
    for w in root.findall('w'):
        well_id = w.get('n')
        try:
            volume = float(w.get('vl'))
        except (TypeError, ValueError):
            volume = None

        if volume is None:
            flag = "Missing"
        elif volume < min_vol:
            flag = "Below working range"
        elif volume > max_vol:
            flag = "Exceeds working range"
        else:
            flag = None

        data.append({'Well': well_id, 'Volume': volume, 'Volume_Flag': flag})

    df = pd.DataFrame(data)

    if compound_lookup:
        df['Compound'] = df['Well'].apply(
            lambda w: compound_lookup.get((plate_barcode, w), 'NA')
        )
        df['Plate_Barcode'] = plate_barcode
        df = df[['Plate_Barcode', 'Well', 'Compound', 'Volume', 'Volume_Flag']]
    else:
        df['Plate_Barcode'] = plate_barcode
        df = df[['Plate_Barcode', 'Well', 'Volume', 'Volume_Flag']]

    metadata_dict = {
        'Plate_Barcode': plate_barcode,
        'Plate_Type': plate_type,
        'DeadVolume': min_vol,
        'MaxVolume': max_vol,
        'Survey_Date': plate_date
    }

    flagged_count = df['Volume_Flag'].notna().sum()
    return {
        'df': df,
        'metadata': metadata_dict,
        'error_summary': None,
        'summary': {
            'plate_barcode': plate_barcode,
            'total_wells': len(df),
            'flagged_wells': flagged_count,
            'unmatched_compounds': (df['Compound'] == 'NA').sum() if 'Compound' in df.columns else 0,
        }
    }


# =============================================================================
# GUI APPLICATION
# =============================================================================

class PickliPyQC(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("PickliPy QC v2.1")
        self.geometry("900x750")
        self.minsize(800, 650)

        self.config = load_config()

        # Variables
        self.xml_mode_var = tk.StringVar(value='print')
        self.picklist_mode_var = tk.StringVar(value=self.config.get('picklist_mode', 'screening'))
        self.time_window_var = tk.StringVar(value=str(self.config['time_window_hours']))
        self.log_dir_var = tk.StringVar(value=self.config['log_directory'])
        self.picklist_file_var = tk.StringVar(value=self.config.get('picklist_file', ''))
        self.output_dir_var = tk.StringVar(value=self.config.get('output_directory', ''))

        # File list with checkboxes
        self.xml_files = []
        self.file_check_vars = []

        # Compound lookup cache
        self.compound_lookup = {}
        self.well_alternatives = {}
        self.dead_volume_ul = 2.5

        self._create_widgets()
        self._refresh_file_list()

    def _create_widgets(self):
        """Build the GUI layout."""

        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # === SETTINGS SECTION ===
        settings_frame = ttk.LabelFrame(main_frame, text="Settings", padding="5")
        settings_frame.pack(fill=tk.X, pady=(0, 10))

        # Row 1: XML Mode selection
        mode_frame = ttk.Frame(settings_frame)
        mode_frame.pack(fill=tk.X, pady=2)

        ttk.Label(mode_frame, text="XML Mode:").pack(side=tk.LEFT)
        ttk.Radiobutton(mode_frame, text="Print XML (post-dispense)",
                        variable=self.xml_mode_var, value='print',
                        command=self._refresh_file_list).pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(mode_frame, text="Survey XML (volume check)",
                        variable=self.xml_mode_var, value='survey',
                        command=self._refresh_file_list).pack(side=tk.LEFT, padx=10)

        # Row 2: Picklist Mode selection
        picklist_mode_frame = ttk.Frame(settings_frame)
        picklist_mode_frame.pack(fill=tk.X, pady=2)

        ttk.Label(picklist_mode_frame, text="Picklist Mode:").pack(side=tk.LEFT)
        ttk.Radiobutton(picklist_mode_frame, text="Assay",
                        variable=self.picklist_mode_var, value='assay').pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(picklist_mode_frame, text="Screening",
                        variable=self.picklist_mode_var, value='screening').pack(side=tk.LEFT, padx=10)

        # Row 3: Log directory
        log_frame = ttk.Frame(settings_frame)
        log_frame.pack(fill=tk.X, pady=2)

        ttk.Label(log_frame, text="Log Directory:").pack(side=tk.LEFT)
        ttk.Entry(log_frame, textvariable=self.log_dir_var, width=60).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(log_frame, text="Browse...", command=self._browse_log_dir).pack(side=tk.LEFT)

        # Row 4: Time window
        time_frame = ttk.Frame(settings_frame)
        time_frame.pack(fill=tk.X, pady=2)

        ttk.Label(time_frame, text="Time Window:").pack(side=tk.LEFT)
        time_entry = ttk.Entry(time_frame, textvariable=self.time_window_var, width=5)
        time_entry.pack(side=tk.LEFT, padx=5)
        ttk.Label(time_frame, text="hours").pack(side=tk.LEFT)
        ttk.Button(time_frame, text="Refresh", command=self._refresh_file_list).pack(side=tk.LEFT, padx=20)

        # Row 5: Picklist design file
        picklist_frame = ttk.Frame(settings_frame)
        picklist_frame.pack(fill=tk.X, pady=2)

        ttk.Label(picklist_frame, text="Picklist Design File:").pack(side=tk.LEFT)
        ttk.Entry(picklist_frame, textvariable=self.picklist_file_var, width=55).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(picklist_frame, text="Browse...", command=self._browse_picklist_file).pack(side=tk.LEFT)

        # Row 6: Output directory
        out_frame = ttk.Frame(settings_frame)
        out_frame.pack(fill=tk.X, pady=2)

        ttk.Label(out_frame, text="Output Directory:").pack(side=tk.LEFT)
        ttk.Entry(out_frame, textvariable=self.output_dir_var, width=55).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(out_frame, text="Browse...", command=self._browse_output_dir).pack(side=tk.LEFT)

        # === FILE LIST SECTION ===
        files_frame = ttk.LabelFrame(main_frame, text="XML Files", padding="5")
        files_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        list_container = ttk.Frame(files_frame)
        list_container.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(list_container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_container, orient=tk.VERTICAL, command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        btn_frame = ttk.Frame(files_frame)
        btn_frame.pack(fill=tk.X, pady=(5, 0))

        ttk.Button(btn_frame, text="Select All", command=self._select_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Deselect All", command=self._deselect_all).pack(side=tk.LEFT, padx=5)

        self.file_count_label = ttk.Label(btn_frame, text="0 files found")
        self.file_count_label.pack(side=tk.RIGHT, padx=5)

        # === ACTION BUTTONS ===
        action_frame = ttk.LabelFrame(main_frame, text="Actions", padding="10")
        action_frame.pack(fill=tk.X, pady=(0, 10))

        btn_row = ttk.Frame(action_frame)
        btn_row.pack(fill=tk.X)

        self.process_btn = ttk.Button(btn_row, text="▶ Process Selected Files",
                                      command=self._process_files)
        self.process_btn.pack(side=tk.LEFT, padx=5, pady=5)

        ttk.Button(btn_row, text="Load Compound Lookup",
                   command=self._load_compound_lookup).pack(side=tk.LEFT, padx=5, pady=5)

        ttk.Button(btn_row, text="Clear Log", command=self._clear_log).pack(side=tk.RIGHT, padx=5, pady=5)

        # === LOG SECTION ===
        log_frame = ttk.LabelFrame(main_frame, text="Status Log", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))

        log_container = ttk.Frame(log_frame)
        log_container.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(log_container, height=8, wrap=tk.WORD, state=tk.DISABLED)
        log_scrollbar = ttk.Scrollbar(log_container, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)

        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def _on_mousewheel(self, event):
        """Handle mousewheel scrolling."""
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def _log(self, message):
        """Add message to log box."""
        self.log_text.configure(state=tk.NORMAL)
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)
        self.update()

    def _clear_log(self):
        """Clear the log box."""
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _browse_log_dir(self):
        """Open dialog to select log directory."""
        path = filedialog.askdirectory(
            initialdir=self.log_dir_var.get(),
            title="Select Echo Log Directory"
        )
        if path:
            self.log_dir_var.set(path)
            self._save_config()
            self._refresh_file_list()

    def _browse_picklist_file(self):
        """Open dialog to select picklist design file."""
        path = filedialog.askopenfilename(
            initialdir=os.path.dirname(self.picklist_file_var.get()) or None,
            title="Select Picklist Design Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if path:
            self.picklist_file_var.set(path)
            if not self.output_dir_var.get():
                self.output_dir_var.set(os.path.dirname(path))
            self._save_config()
            self._log(f"Selected picklist file: {os.path.basename(path)}")

    def _browse_output_dir(self):
        """Open dialog to select output directory."""
        path = filedialog.askdirectory(
            initialdir=self.output_dir_var.get() or None,
            title="Select Output Directory"
        )
        if path:
            self.output_dir_var.set(path)
            self._save_config()

    def _load_compound_lookup(self):
        """Load compound lookup from picklist design file."""
        picklist_file = self.picklist_file_var.get()
        picklist_mode = self.picklist_mode_var.get()

        if not picklist_file:
            self._log("⚠ No picklist design file selected")
            return

        self.compound_lookup, self.well_alternatives, self.dead_volume_ul = build_compound_lookup(
            picklist_file, picklist_mode, log_func=self._log
        )

    def _refresh_file_list(self):
        """Scan log directory and update file list with checkboxes."""
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.file_check_vars = []

        try:
            hours = float(self.time_window_var.get())
        except ValueError:
            hours = DEFAULT_TIME_WINDOW

        mode = self.xml_mode_var.get()
        self.xml_files = get_xml_files(self.log_dir_var.get(), mode, hours)

        for i, f in enumerate(self.xml_files):
            var = tk.BooleanVar(value=False)
            self.file_check_vars.append(var)

            frame = ttk.Frame(self.scrollable_frame)
            frame.pack(fill=tk.X, pady=1)

            cb = ttk.Checkbutton(frame, variable=var)
            cb.pack(side=tk.LEFT)

            modified_str = f['modified'].strftime('%Y-%m-%d %H:%M')
            label_text = f"{f['filename']}  |  {modified_str}  |  {f['barcodes']}"
            label = ttk.Label(frame, text=label_text)
            label.pack(side=tk.LEFT, padx=5)

            label.bind("<Button-1>", lambda e, v=var: v.set(not v.get()))

        self.file_count_label.config(text=f"{len(self.xml_files)} files found")
        self._save_config()

    def _select_all(self):
        for var in self.file_check_vars:
            var.set(True)

    def _deselect_all(self):
        for var in self.file_check_vars:
            var.set(False)

    def _save_config(self):
        """Save current settings to config file."""
        try:
            hours = float(self.time_window_var.get())
        except ValueError:
            hours = DEFAULT_TIME_WINDOW

        self.config.update({
            'log_directory': self.log_dir_var.get(),
            'picklist_file': self.picklist_file_var.get(),
            'output_directory': self.output_dir_var.get(),
            'time_window_hours': hours,
            'picklist_mode': self.picklist_mode_var.get(),
        })
        save_config(self.config)

    def _process_files(self):
        """Process selected XML files."""
        selected_indices = [i for i, var in enumerate(self.file_check_vars) if var.get()]

        if not selected_indices:
            self._log(" No files selected. Please select at least one file.")
            return

        self._log(f"Processing {len(selected_indices)} file(s)...")

        if not self.compound_lookup:
            self._load_compound_lookup()

        output_dir = self.output_dir_var.get()
        if not output_dir:
            output_dir = self.log_dir_var.get()
            self.output_dir_var.set(output_dir)

        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
                self._log(f"Created output directory: {output_dir}")
            except OSError as e:
                self._log(f" Could not create output directory: {e}")
                return

        mode = self.xml_mode_var.get()
        parsed_results = []
        errors = []

        if mode == 'print':
            selected_print_files = [self.xml_files[i]['path'] for i in selected_indices]
            self._log("Matching survey files to print files...")
            survey_matches = find_matching_surveys(
                self.log_dir_var.get(), selected_print_files, log_func=self._log
            )

        for idx in selected_indices:
            file_info = self.xml_files[idx]
            xml_path = file_info['path']

            try:
                self._log(f"Processing: {file_info['filename']}")

                if mode == 'print':
                    result = parse_printmap(xml_path, self.compound_lookup or None)
                else:
                    result = parse_platesurvey(xml_path, self.compound_lookup or None)

                result['filename'] = file_info['filename']
                result['filepath'] = xml_path

                if mode == 'print':
                    result['matched_survey'] = survey_matches.get(xml_path)

                parsed_results.append(result)

                summary = result['summary']
                if 'src_barcode' in summary:
                    self._log(f"  ✓ {summary['total_transfers']} transfers, {summary['skipped_transfers']} skipped")
                else:
                    self._log(f"  ✓ {summary['total_wells']} wells, {summary['flagged_wells']} flagged")

            except Exception as e:
                self._log(f"  ✗ Error: {str(e)}")
                errors.append(f"{file_info['filename']}: {str(e)}")

        if parsed_results:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_file = os.path.join(output_dir, f"PickliPy_QC_{mode}_{timestamp}.xlsx")
            self._write_combined_output(parsed_results, output_file, mode)
            self._log(f"✓ Output saved to: {output_file}")

            if mode == 'print':
                rerun_path = self._generate_smart_rerun_picklist(parsed_results, output_dir)
                if rerun_path:
                    self._log(f" Rerun picklist saved to: {rerun_path}")

        self._log("=" * 50)
        self._log(f"Processing complete: {len(parsed_results)} succeeded, {len(errors)} failed")
        if errors:
            self._log("Errors:")
            for e in errors:
                self._log(f"  • {e}")

    def _write_combined_output(self, parsed_results, output_file, mode):
        """Write all parsed results to a single Excel file with multiple sheets."""

        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            workbook = writer.book

            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#D9E1F2',
                'border': 1
            })
            skipped_format = workbook.add_format({
                'bg_color': '#FFCCCC'
            })
            vol_diff_format = workbook.add_format({
                'bg_color': '#FFFFCC'
            })
            metadata_key_format = workbook.add_format({
                'bold': True
            })

            # Build grouped_results:
            # Print mode: merge XMLs sharing the same dst barcode (handles power assay)
            # Survey mode: each XML gets its own sheet in chronological order
            if mode == 'print':
                grouped_results = defaultdict(list)
                for result in parsed_results:
                    key = result['metadata']['DST_Barcode']
                    grouped_results[key].append(result)
            else:
                grouped_results = {}
                barcode_counters = Counter()
                for result in parsed_results:
                    barcode = result['metadata']['Plate_Barcode']
                    barcode_counters[barcode] += 1
                    key = f"{barcode}_{barcode_counters[barcode]}"
                    grouped_results[key] = [result]

            for dst_key, results in grouped_results.items():
                merged_df = pd.concat([r['df']
                                      for r in results], ignore_index=True)

                metadata = results[0]['metadata']
                error_summary = Counter()
                for r in results:
                    if r.get('error_summary'):
                        error_summary.update(r['error_summary'])

                if mode == 'print':
                    sheet_name = metadata['DST_Barcode'][:31]
                else:
                    sheet_name = dst_key[:31]

                worksheet = workbook.add_worksheet(sheet_name)
                writer.sheets[sheet_name] = worksheet

                df = merged_df
                current_row = 0

                for key, value in metadata.items():
                    worksheet.write(current_row, 0, key, metadata_key_format)
                    worksheet.write(current_row, 1, value)
                    current_row += 1

                current_row += 1

                if error_summary:
                    worksheet.write(
                        current_row, 0, "Echo Error Summary:", metadata_key_format)
                    current_row += 1
                    worksheet.write(current_row, 0, "Error Code")
                    worksheet.write(current_row, 1, "Description")
                    worksheet.write(current_row, 2, "Count")
                    current_row += 1

                    for error_entry, count in error_summary.items():
                        if ':' in error_entry:
                            code, explanation = error_entry.split(':', 1)
                            worksheet.write(current_row, 0, code.strip())
                            worksheet.write(
                                current_row, 1, explanation.strip())
                            worksheet.write(current_row, 2, count)
                            current_row += 1

                    current_row += 1

                header_row = current_row
                for col, column_name in enumerate(df.columns):
                    worksheet.write(header_row, col,
                                    column_name, header_format)
                current_row += 1

                for row_idx, (_, row_data) in enumerate(df.iterrows()):
                    data_row = current_row + row_idx

                    is_skipped = row_data.get('Skipped') == 'Yes'
                    has_vol_diff = row_data.get('Vol_Difference') != '' and pd.notna(
                        row_data.get('Vol_Difference'))

                    for col, value in enumerate(row_data):
                        if is_skipped:
                            worksheet.write(
                                data_row, col, value, skipped_format)
                        elif has_vol_diff:
                            worksheet.write(
                                data_row, col, value, vol_diff_format)
                        else:
                            worksheet.write(data_row, col, value)

                worksheet.freeze_panes(header_row + 1, 0)

                for col, column_name in enumerate(df.columns):
                    max_len = max(
                        len(str(column_name)),
                        df[column_name].astype(
                            str).str.len().max() if len(df) > 0 else 0
                    )
                    worksheet.set_column(col, col, min(max_len + 2, 50))

            # ---- Run Summary sheet (always last) ----
            summary_sheet = workbook.add_worksheet('Run Summary')
            writer.sheets['Run Summary'] = summary_sheet

            if mode == 'print':
                summary_headers = ['Source_XML', 'SRC_Barcode', 'DST_Barcode', 'Transfer_Date',
                                'Total_Transfers', 'Skipped_Transfers', 'Errors']


            else:
                summary_headers = ['Source_XML', 'Plate_Barcode', 'Plate_Type', 'Survey_Date',
                                'Total_Wells', 'Flagged_Wells']
            for col, h in enumerate(summary_headers):
                summary_sheet.write(0, col, h, header_format)

            for row_idx, result in enumerate(parsed_results, start=1):
                meta = result['metadata']
                summary_sheet.write(row_idx, 0, result['filename'])

                if mode == 'print':
                    summary_sheet.write(
                        row_idx, 1, meta.get('SRC_Barcode', ''))
                    summary_sheet.write(
                        row_idx, 2, meta.get('DST_Barcode', ''))
                    summary_sheet.write(
                        row_idx, 3, meta.get('Transfer_Date', ''))
                    summary_sheet.write(
                        row_idx, 4, result['summary'].get('total_transfers', ''))
                    summary_sheet.write(
                        row_idx, 5, result['summary'].get('skipped_transfers', ''))
                    error_count = sum(result['error_summary'].values()) if result.get(
                        'error_summary') else 0
                    summary_sheet.write(row_idx, 6, error_count)
                else:
                    summary_sheet.write(row_idx, 1, meta.get('Plate_Barcode', ''))
                    summary_sheet.write(row_idx, 2, meta.get('Plate_Type', ''))
                    summary_sheet.write(row_idx, 3, meta.get('Survey_Date', ''))
                    summary_sheet.write(row_idx, 4, result['summary'].get('total_wells', ''))
                    summary_sheet.write(row_idx, 5, result['summary'].get('flagged_wells', ''))

            for col, h in enumerate(summary_headers):
                summary_sheet.set_column(col, col, max(len(h) + 2, 20))

    def _generate_smart_rerun_picklist(self, parsed_results, output_dir):
        """
        Generate a smart re-run picklist with volume-aware well switching.

        Uses:
        - Print XML post-dispense volumes for used wells
        - Survey XML pre-dispense volumes for unused alternative wells
        """
        skipped_transfers = []
        for result in parsed_results:
            df = result['df']
            src_barcode = result.get('src_barcode', result['metadata'].get('SRC_Barcode'))
            matched_survey = result.get('matched_survey')

            survey_volumes = {}
            if matched_survey:
                survey_volumes = load_survey_volumes(matched_survey)

            print_volumes = {}
            for _, row in df.iterrows():
                well = row['Well']
                if pd.notna(row.get('SRC_Vol_After')):
                    print_volumes[well] = row['SRC_Vol_After']

            skipped_df = df[df['Skipped'] == 'Yes']
            for _, row in skipped_df.iterrows():
                skipped_transfers.append({
                    'src_barcode': src_barcode,
                    'src_well': row['Well'],
                    'dst_barcode': row['DST_Barcode'],
                    'dst_well': row['Destination_Well'],
                    'volume_ul': row['Vol_Target'],
                    'compound': row.get('Compound', 'NA'),
                    'print_volumes': print_volumes,
                    'survey_volumes': survey_volumes,
                })

        if not skipped_transfers:
            self._log("No skipped transfers found - no rerun picklist needed")
            return None

        self._log(f"Generating smart rerun picklist for {len(skipped_transfers)} skipped transfers...")

        dead_volume_ul = self.dead_volume_ul
        rerun_lines = []
        excluded_count = 0
        switched_count = 0

        for t in skipped_transfers:
            src_barcode = t['src_barcode']
            original_well = t['src_well']
            dst_barcode = t['dst_barcode']
            dst_well = t['dst_well']
            volume_nl = t['volume_ul'] * 1000.0
            compound = t['compound']

            chosen_well = original_well
            needed_ul = dead_volume_ul + t['volume_ul']

            if compound and compound != 'NA' and self.well_alternatives:
                alternatives = self.well_alternatives.get((src_barcode, compound), [original_well])

                for candidate in alternatives:
                    if candidate in t['print_volumes']:
                        available = t['print_volumes'][candidate]
                    elif candidate in t['survey_volumes']:
                        available = t['survey_volumes'][candidate]
                    else:
                        continue

                    if available >= needed_ul:
                        if candidate != original_well:
                            self._log(f"  Switching {compound}: {original_well} → {candidate} ({available:.1f} µL available)")
                            switched_count += 1
                        chosen_well = candidate
                        break
                else:
                    self._log(f"  No well with sufficient volume for {compound}, excluding from rerun")
                    excluded_count += 1
                    continue

            rerun_lines.append({
                'src_barcode': src_barcode,
                'src_well': chosen_well,
                'dst_barcode': dst_barcode,
                'dst_well': dst_well,
                'volume_nl': volume_nl
            })

        if not rerun_lines:
            self._log("All skipped transfers excluded due to insufficient volume")
            return None

        grouped = defaultdict(list)
        for line in rerun_lines:
            key = (line['src_barcode'], line['src_well'], line['dst_barcode'])
            grouped[key].append({
                'dst_well': line['dst_well'],
                'volume_nl': line['volume_nl']
            })

        csv_lines = []
        for (src_barcode, src_well, dst_barcode), transfers in grouped.items():
            line_parts = [src_barcode, src_well, dst_barcode]
            for t in transfers:
                line_parts.append(f"{t['volume_nl']:.1f}")
                line_parts.append(t['dst_well'])
            csv_lines.append(','.join(line_parts))

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(output_dir, f"rerun_picklist_{timestamp}.csv")

        with open(output_path, 'w') as f:
            f.write('\n'.join(csv_lines))

        self._log(f"  {len(rerun_lines)} transfers in rerun ({switched_count} well switches, {excluded_count} excluded)")
        return output_path


# =============================================================================
# MAIN ENTRY POINT
# =============================================================================

def main():
    app = PickliPyQC()
    app.mainloop()


if __name__ == '__main__':
    main()
