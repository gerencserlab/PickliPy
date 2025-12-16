"""
PickliPy QC - Post-execution validation tool for Echo acoustic dispenser
Parses print and survey XML files and matches to compound library data.

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
from collections import Counter

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
    # Add more plate types as needed
}


# =============================================================================
# CONFIG MANAGEMENT
# =============================================================================

def load_config():
    """Load saved configuration from JSON file."""
    default_config = {
        'log_directory': DEFAULT_LOG_DIR,
        'library_file': '',
        'output_directory': '',
        'time_window_hours': DEFAULT_TIME_WINDOW,
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
    
    # Determine file suffix based on mode
    suffix = '_print.xml' if mode == 'print' else '_platesurvey.xml'
    
    log_path = Path(log_dir)
    if not log_path.exists():
        return files
    
    for filepath in log_path.glob(f'*{suffix}'):
        try:
            # Check modification time
            mtime = datetime.fromtimestamp(filepath.stat().st_mtime)
            if mtime < cutoff_time:
                continue
            
            # Extract barcodes from XML
            barcodes = extract_barcodes(filepath, mode)
            
            files.append({
                'path': str(filepath),
                'filename': filepath.name,
                'modified': mtime,
                'barcodes': barcodes,
            })
        except Exception as e:
            print(f"Warning: Could not process {filepath}: {e}")
    
    # Sort by modification time, newest first
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
        else:  # survey
            barcode = root.attrib.get('barcode', '?')
            return barcode
            
    except Exception:
        return '?'
    
    return '?'


# =============================================================================
# PRINT XML PARSER
# =============================================================================

def parse_printmap(xml_file, lib_file):
    """
    Parse print XML file and match to compound library.
    
    Args:
        xml_file: Path to _print.xml file
        lib_file: Path to Excel library file
        
    
    Returns:
        dict with summary info
    """
    # Parse XML
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

    # Working range from lookup
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

            # Flags
            exceeds_range = (
                src_max_vol is not None and
                (pre_src_vol > src_max_vol)
            )
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

    # Load library file and match compounds
    xls = pd.ExcelFile(lib_file)

    if src_barcode not in xls.sheet_names:
        raise ValueError(f"No sheet found for source barcode: {src_barcode}")

    lib_df = pd.read_excel(xls, sheet_name=src_barcode)

    # Validate barcode in sheet
    if 'Plate Barcode' in lib_df.columns:
        src_barcodes_in_sheet = lib_df['Plate Barcode'].dropna().unique()
        if src_barcode not in src_barcodes_in_sheet:
            raise ValueError(f"Sheet '{src_barcode}' does not contain matching 'Plate Barcode' values.")

    # Rename RACKPOS -> Well
    lib_df = lib_df.rename(columns={'RACKPOS': 'Well'})

    # Validate required columns
    expected_cols = {'Well', 'Compound'}
    if not expected_cols.issubset(lib_df.columns):
        raise ValueError(f"Missing expected columns in sheet {src_barcode}: {expected_cols - set(lib_df.columns)}")

    # Merge
    merged = pd.merge(df, lib_df[['Well', 'Compound']], on='Well', how='left')
    merged['Compound'] = merged['Compound'].fillna('NA')

   # Add barcode columns
    merged['SRC_Barcode'] = src_barcode
    merged['DST_Barcode'] = dst_barcode
    
    # Calculate Vol_Difference
    merged['Vol_Difference'] = merged['Vol_Actual'] - merged['Vol_Target']
    
    # Conditional display: blank unless meaningful
    # Skipped: blank unless True
    merged['Skipped'] = merged['Skipped'].apply(lambda x: 'Yes' if x else '')
    # Reason_Skipped: already None/blank for successful transfers
    merged['Reason_Skipped'] = merged['Reason_Skipped'].fillna('')
    # ViolatesDeadVol: blank unless True
    merged['ViolatesDeadVol'] = merged['ViolatesDeadVol'].apply(lambda x: 'Yes' if x else '')
    # Vol_Difference: blank unless non-zero
    merged['Vol_Difference'] = merged['Vol_Difference'].apply(lambda x: x if x != 0 else '')

    # Reorder columns
    final_columns = [
        'SRC_Barcode', 'Well', 'Compound', 
        'DST_Barcode', 'Destination_Well',
        'Vol_Target', 'Vol_Actual', 'Vol_Difference',
        'SRC_Vol_Before', 'SRC_Vol_After',
        'Skipped', 'Reason_Skipped', 'ViolatesDeadVol'
    ]
    df_final = merged[final_columns]
    
    # Metadata
    metadata_dict = {
        'SRC_Barcode': src_barcode,
        'SRC_PlateType': src_plate_type,
        'SRC_DeadVolume': src_dead_vol,
        'SRC_MaxVolume': src_max_vol,
        'DST_Barcode': dst_barcode,
        'DST_PlateType': dst_plate_type,
        'Transfer_Date': transfer_date
    }

    # Return data for combined output (don't write file here)
    return {
        'df': df_final,
        'metadata': metadata_dict,
        'error_summary': error_summary,
        'summary': {
            'src_barcode': src_barcode,
            'dst_barcode': dst_barcode,
            'total_transfers': len(transfer_rows),
            'skipped_transfers': len(skipped_rows),
            'unmatched_compounds': (df_final['Compound'] == 'NA').sum(),
        }
    }


# =============================================================================
# SURVEY XML PARSER
# =============================================================================

def parse_platesurvey(xml_file, lib_file):
    """
    Parse survey XML file and match to compound library.
    
    Args:
        xml_file: Path to _platesurvey.xml file
        lib_file: Path to Excel library file
       
    
    Returns:
        dict with summary info
    """
    # Parse XML
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

    # Extract well data
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

    # Load library file
    xls = pd.ExcelFile(lib_file)

    if plate_barcode not in xls.sheet_names:
        raise ValueError(f"No sheet found for source barcode: {plate_barcode}")

    lib_df = pd.read_excel(xls, sheet_name=plate_barcode)

    # Validate barcode
    if 'Plate Barcode' in lib_df.columns:
        plate_barcodes_in_sheet = lib_df['Plate Barcode'].dropna().unique()
        if plate_barcode not in plate_barcodes_in_sheet:
            raise ValueError(f"Sheet '{plate_barcode}' does not contain matching 'Plate Barcode' values.")

    # Rename RACKPOS -> Well
    lib_df = lib_df.rename(columns={'RACKPOS': 'Well'})

    # Validate required columns
    expected_cols = {'Well', 'Compound'}
    if not expected_cols.issubset(lib_df.columns):
        raise ValueError(f"Missing expected columns in sheet {plate_barcode}: {expected_cols - set(lib_df.columns)}")

    # Merge
    merged = pd.merge(df, lib_df[['Well', 'Compound']], on='Well', how='left')
    merged['Compound'] = merged['Compound'].fillna('NA')

    # Add barcode column
    merged['Plate_Barcode'] = plate_barcode

    # Reorder
    merged = merged[['Plate_Barcode', 'Well', 'Compound', 'Volume', 'Volume_Flag']]

    # Metadata
    metadata_dict = {
        'Plate_Barcode': plate_barcode,
        'Plate_Type': plate_type,
        'DeadVolume': min_vol,
        'MaxVolume': max_vol,
        'Survey_Date': plate_date
    }

    # Return data for combined output (don't write file here)
    flagged_count = merged['Volume_Flag'].notna().sum()
    return {
        'df': merged,
        'metadata': metadata_dict,
        'error_summary': None,
        'summary': {
            'plate_barcode': plate_barcode,
            'total_wells': len(merged),
            'flagged_wells': flagged_count,
            'unmatched_compounds': (merged['Compound'] == 'NA').sum(),
        }
    }

# =============================================================================
# GUI APPLICATION
# =============================================================================

class PickliPyQC(tk.Tk):
    def __init__(self):
        super().__init__()
        
        self.title("PickliPy QC")
        self.geometry("800x600")
        self.minsize(700, 500)
        
        # Load saved config
        self.config = load_config()
        
        # Variables
        self.mode_var = tk.StringVar(value='print')
        self.time_window_var = tk.StringVar(value=str(self.config['time_window_hours']))
        self.log_dir_var = tk.StringVar(value=self.config['log_directory'])
        self.lib_file_var = tk.StringVar(value=self.config['library_file'])
        self.output_dir_var = tk.StringVar(value=self.config['output_directory'])
        
        # File list data
        self.xml_files = []
        
        # Build UI
        self._create_widgets()
        
        # Initial file scan
        self._refresh_file_list()
    
    def _create_widgets(self):
        """Build the GUI layout."""
        
        # Main container with padding
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # === TOP SECTION: Mode and Settings ===
        settings_frame = ttk.LabelFrame(main_frame, text="Settings", padding="5")
        settings_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Row 1: Mode selection
        mode_frame = ttk.Frame(settings_frame)
        mode_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(mode_frame, text="Mode:").pack(side=tk.LEFT)
        ttk.Radiobutton(mode_frame, text="Print XML (post-dispense)", 
                       variable=self.mode_var, value='print',
                       command=self._refresh_file_list).pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(mode_frame, text="Survey XML (volume check)", 
                       variable=self.mode_var, value='survey',
                       command=self._refresh_file_list).pack(side=tk.LEFT, padx=10)
        
        # Row 2: Log directory
        log_frame = ttk.Frame(settings_frame)
        log_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(log_frame, text="Log Directory:").pack(side=tk.LEFT)
        ttk.Entry(log_frame, textvariable=self.log_dir_var, width=60).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(log_frame, text="Browse...", command=self._browse_log_dir).pack(side=tk.LEFT)
        
        # Row 3: Time window
        time_frame = ttk.Frame(settings_frame)
        time_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(time_frame, text="Time Window:").pack(side=tk.LEFT)
        time_entry = ttk.Entry(time_frame, textvariable=self.time_window_var, width=5)
        time_entry.pack(side=tk.LEFT, padx=5)
        ttk.Label(time_frame, text="hours").pack(side=tk.LEFT)
        ttk.Button(time_frame, text="Refresh", command=self._refresh_file_list).pack(side=tk.LEFT, padx=20)
        
        # Row 4: Library file
        lib_frame = ttk.Frame(settings_frame)
        lib_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(lib_frame, text="Library File:").pack(side=tk.LEFT)
        ttk.Entry(lib_frame, textvariable=self.lib_file_var, width=60).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(lib_frame, text="Browse...", command=self._browse_lib_file).pack(side=tk.LEFT)
        
        # Row 5: Output directory
        out_frame = ttk.Frame(settings_frame)
        out_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(out_frame, text="Output Directory:").pack(side=tk.LEFT)
        ttk.Entry(out_frame, textvariable=self.output_dir_var, width=60).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(out_frame, text="Browse...", command=self._browse_output_dir).pack(side=tk.LEFT)
        
        # === MIDDLE SECTION: File List ===
        files_frame = ttk.LabelFrame(main_frame, text="XML Files", padding="5")
        files_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Treeview with scrollbar
        tree_frame = ttk.Frame(files_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        self.tree = ttk.Treeview(tree_frame, columns=('filename', 'modified', 'barcodes'), 
                                  show='headings', selectmode='extended')
        self.tree.heading('filename', text='Filename')
        self.tree.heading('modified', text='Modified')
        self.tree.heading('barcodes', text='Barcodes')
        
        self.tree.column('filename', width=350)
        self.tree.column('modified', width=150)
        self.tree.column('barcodes', width=200)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Selection buttons
        btn_frame = ttk.Frame(files_frame)
        btn_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Button(btn_frame, text="Select All", command=self._select_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Clear Selection", command=self._clear_selection).pack(side=tk.LEFT, padx=5)
        
        self.file_count_label = ttk.Label(btn_frame, text="0 files found")
        self.file_count_label.pack(side=tk.RIGHT, padx=5)
        
        # === BOTTOM SECTION: Actions ===
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X)
        
        self.process_btn = ttk.Button(action_frame, text="Process Selected Files", 
                                       command=self._process_files)
        self.process_btn.pack(side=tk.RIGHT, padx=5)
        
        self.status_label = ttk.Label(action_frame, text="Ready")
        self.status_label.pack(side=tk.LEFT, padx=5)
    
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
    
    def _browse_lib_file(self):
        """Open dialog to select library file."""
        path = filedialog.askopenfilename(
            initialdir=os.path.dirname(self.lib_file_var.get()) or None,
            title="Select Library Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if path:
            self.lib_file_var.set(path)
            # Set output directory to same as library file if not set
            if not self.output_dir_var.get():
                self.output_dir_var.set(os.path.dirname(path))
            self._save_config()
    
    def _browse_output_dir(self):
        """Open dialog to select output directory."""
        path = filedialog.askdirectory(
            initialdir=self.output_dir_var.get() or None,
            title="Select Output Directory"
        )
        if path:
            self.output_dir_var.set(path)
            self._save_config()
    
    def _refresh_file_list(self):
        """Scan log directory and update file list."""
        # Clear current list
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Get time window
        try:
            hours = float(self.time_window_var.get())
        except ValueError:
            hours = DEFAULT_TIME_WINDOW
        
        # Scan for files
        mode = self.mode_var.get()
        self.xml_files = get_xml_files(self.log_dir_var.get(), mode, hours)
        
        # Populate tree
        for f in self.xml_files:
            modified_str = f['modified'].strftime('%Y-%m-%d %H:%M:%S')
            self.tree.insert('', tk.END, values=(f['filename'], modified_str, f['barcodes']))
        
        # Update count label
        self.file_count_label.config(text=f"{len(self.xml_files)} files found")
        
        # Save config
        self._save_config()
    
    def _select_all(self):
        """Select all items in the tree."""
        for item in self.tree.get_children():
            self.tree.selection_add(item)
    
    def _clear_selection(self):
        """Clear all selections."""
        self.tree.selection_remove(*self.tree.selection())
    
    def _save_config(self):
        """Save current settings to config file."""
        try:
            hours = float(self.time_window_var.get())
        except ValueError:
            hours = DEFAULT_TIME_WINDOW
            
        self.config.update({
            'log_directory': self.log_dir_var.get(),
            'library_file': self.lib_file_var.get(),
            'output_directory': self.output_dir_var.get(),
            'time_window_hours': hours,
        })
        save_config(self.config)
    
    def _process_files(self):
        """Process selected XML files and write to single output file."""
        # Get selected items
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Please select at least one file to process.")
            return
        
        # Validate library file
        lib_file = self.lib_file_var.get()
        if not lib_file or not os.path.exists(lib_file):
            messagebox.showerror("Error", "Please select a valid library Excel file.")
            return
        
        # Validate output directory
        output_dir = self.output_dir_var.get()
        if not output_dir:
            output_dir = os.path.dirname(lib_file)
            self.output_dir_var.set(output_dir)
        
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except OSError as e:
                messagebox.showerror("Error", f"Could not create output directory: {e}")
                return
        
        # Process each selected file
        mode = self.mode_var.get()
        parse_func = parse_printmap if mode == 'print' else parse_platesurvey
        
        parsed_results = []
        errors = []
        
        self.status_label.config(text="Processing...")
        self.update()
        
        for item in selected:
            idx = self.tree.index(item)
            file_info = self.xml_files[idx]
            xml_path = file_info['path']
            
            try:
                result = parse_func(xml_path, lib_file)
                result['filename'] = file_info['filename']
                parsed_results.append(result)
            except Exception as e:
                errors.append(f"{file_info['filename']}: {str(e)}")
        
       # Write all results to single Excel file
        if parsed_results:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_file = os.path.join(output_dir, f"PickliPy_QC_{mode}_{timestamp}.xlsx")
            self._write_combined_output(parsed_results, output_file, mode)
            
            # Generate rerun picklist if there are skipped transfers (print mode only)
            rerun_path = None
            if mode == 'print':
                rerun_path = self._generate_rerun_picklist(parsed_results, output_dir)
            
            # Show results summary
            self._show_results(parsed_results, errors, output_file, rerun_path)
        else:
            self._show_results(parsed_results, errors, None, None)
        
        self.destroy()

    def _write_combined_output(self, parsed_results, output_file, mode):
        """Write all parsed results to a single Excel file with multiple sheets."""
        
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Define formats
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#D9E1F2',
                'border': 1
            })
            skipped_format = workbook.add_format({
                'bg_color': '#FFCCCC'  # Light red
            })
            vol_diff_format = workbook.add_format({
                'bg_color': '#FFFFCC'  # Light yellow
            })
            metadata_key_format = workbook.add_format({
                'bold': True
            })
            
            for result in parsed_results:
                df = result['df']
                metadata = result['metadata']
                error_summary = result.get('error_summary')
                
                # Generate sheet name from barcodes (max 31 chars for Excel)
                if mode == 'print':
                    sheet_name = f"{metadata['SRC_Barcode']}_{metadata['DST_Barcode']}"
                else:
                    sheet_name = metadata['Plate_Barcode']
                
                # Truncate if needed and ensure uniqueness
                sheet_name = sheet_name[:31]
                
                # Create worksheet
                worksheet = workbook.add_worksheet(sheet_name)
                writer.sheets[sheet_name] = worksheet
                
                current_row = 0
                
                # Write metadata
                for key, value in metadata.items():
                    worksheet.write(current_row, 0, key, metadata_key_format)
                    worksheet.write(current_row, 1, value)
                    current_row += 1
                
                current_row += 1  # Blank row
                
                # Write error summary if present (print mode only)
                if error_summary:
                    worksheet.write(current_row, 0, "Echo Error Summary:", metadata_key_format)
                    current_row += 1
                    worksheet.write(current_row, 0, "Error Code")
                    worksheet.write(current_row, 1, "Description")
                    worksheet.write(current_row, 2, "Count")
                    current_row += 1
                    
                    for error_entry, count in error_summary.items():
                        if ':' in error_entry:
                            code, explanation = error_entry.split(':', 1)
                            worksheet.write(current_row, 0, code.strip())
                            worksheet.write(current_row, 1, explanation.strip())
                            worksheet.write(current_row, 2, count)
                            current_row += 1
                    
                    current_row += 1  # Blank row
                
                # Write header row
                header_row = current_row
                for col, column_name in enumerate(df.columns):
                    worksheet.write(header_row, col, column_name, header_format)
                current_row += 1
                
                # Write data rows with conditional formatting
                for row_idx, (_, row_data) in enumerate(df.iterrows()):
                    data_row = current_row + row_idx
                    
                    # Check if this row should be highlighted
                    is_skipped = row_data.get('Skipped') == 'Yes'
                    has_vol_diff = row_data.get('Vol_Difference') != '' and pd.notna(row_data.get('Vol_Difference'))
                    
                    for col, value in enumerate(row_data):
                        # Choose format based on row status
                        if is_skipped:
                            worksheet.write(data_row, col, value, skipped_format)
                        elif has_vol_diff:
                            worksheet.write(data_row, col, value, vol_diff_format)
                        else:
                            worksheet.write(data_row, col, value)
                
                # Freeze panes at header row (row after header, column 0)
                worksheet.freeze_panes(header_row + 1, 0)
                
                # Auto-fit column widths (approximate)
                for col, column_name in enumerate(df.columns):
                    max_len = max(
                        len(str(column_name)),
                        df[column_name].astype(str).str.len().max() if len(df) > 0 else 0
                    )
                    worksheet.set_column(col, col, min(max_len + 2, 50))

    def _generate_rerun_picklist(self, parsed_results, output_dir):
        """
        Generate a re-run picklist containing only skipped transfers.
        
        Args:
            parsed_results: List of parsed result dicts from processing
            output_dir: Directory to save the rerun picklist
        
        Returns:
            Path to generated file, or None if no skipped transfers
        """
        from collections import defaultdict
        
        # Collect all skipped transfers across all processed files
        skipped_transfers = []
        for result in parsed_results:
            df = result['df']
            # Filter to skipped rows only
            skipped_df = df[df['Skipped'] == 'Yes']
            for _, row in skipped_df.iterrows():
                skipped_transfers.append(row.to_dict())
        
        if not skipped_transfers:
            return None
        
        # Group by (src_barcode, src_well, dst_barcode) to combine into single lines
        grouped = defaultdict(list)
        for t in skipped_transfers:
            key = (t['SRC_Barcode'], t['Well'], t['DST_Barcode'])
            grouped[key].append({
                'dst_well': t['Destination_Well'],
                'volume': t['Vol_Target']
            })
        
        # Generate picklist lines
        lines = []
        for (src_barcode, src_well, dst_barcode), transfers in grouped.items():
            # Build line: src_barcode, src_well, dst_barcode, vol1, dst_well1, vol2, dst_well2, ...
            line_parts = [src_barcode, src_well, dst_barcode]
            for t in transfers:
                line_parts.append(str(t['volume']))
                line_parts.append(t['dst_well'])
            lines.append(','.join(line_parts))
        
        # Generate output filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(output_dir, f"rerun_picklist_{timestamp}.csv")
        
        with open(output_path, 'w') as f:
            f.write('\n'.join(lines))
        
        return output_path       
    
    def _show_results(self, parsed_results, errors, output_file, rerun_path=None):
        """Display processing results in a dialog."""
        msg = f"Processed {len(parsed_results)} file(s) successfully.\n\n"
        
        if parsed_results:
            msg += "Summary:\n"
            for r in parsed_results:
                summary = r['summary']
                if 'src_barcode' in summary:  # Print mode
                    msg += f"  • {summary['src_barcode']} → {summary['dst_barcode']}: "
                    msg += f"{summary['total_transfers']} transfers, {summary['skipped_transfers']} skipped\n"
                else:  # Survey mode
                    msg += f"  • {summary['plate_barcode']}: "
                    msg += f"{summary['total_wells']} wells, {summary['flagged_wells']} flagged\n"
        
        if errors:
            msg += f"\n{len(errors)} error(s):\n"
            for e in errors[:5]:
                msg += f"  • {e}\n"
            if len(errors) > 5:
                msg += f"  ... and {len(errors) - 5} more\n"
        
        if output_file:
            msg += f"\nOutput saved to:\n{output_file}"
        
        if rerun_path:
            msg += f"\n\n Skipped transfers detected!\nRerun picklist saved to:\n{rerun_path}"
        
        if errors:
            messagebox.showwarning("Processing Complete", msg)
        else:
            messagebox.showinfo("Processing Complete", msg)

# =============================================================================
# MAIN ENTRY POINT
# =============================================================================

def main():
    app = PickliPyQC()
    app.mainloop()


if __name__ == '__main__':
    main()
