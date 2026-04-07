#!/usr/bin/env python3
"""
RVTools Cluster Summary Generator
==================================
Processes RVTools Excel exports to generate formatted cluster summary reports.

Output: Excel workbook with per-cluster summary tables including:
  - VM counts, vCPU totals, memory allocation
  - Host hardware details (model, CPU, memory per host)
  - Storage utilization across clusters

Compatibility: Designed for RVTools 4.x Excel exports (.xlsx)
Required sheets: vHost, vInfo, vCPU, vMemory, vDisk, vPartition

Install dependencies: pip install pandas openpyxl
"""

import pandas as pd
import os
from pathlib import Path
from datetime import datetime

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    print("WARNING: openpyxl not installed. Run: pip install openpyxl")
    print("Falling back to CSV output.")
    OPENPYXL_AVAILABLE = False

# ============================================================================
# CONFIGURATION
# ============================================================================
# Use the directory where the script is located
SCRIPT_DIR = Path(__file__).parent.resolve()
INPUT_FILES_DIR = SCRIPT_DIR / "input_files"  # Place collector Excel files here
RESULTS_DIR = SCRIPT_DIR / "results"
DATE_STR = datetime.now().strftime("%Y-%m-%d")

# Create folders if they don't exist
INPUT_FILES_DIR.mkdir(exist_ok=True)
RESULTS_DIR.mkdir(exist_ok=True)

# ============================================================================
# SHEET NAME ALIASES
# ============================================================================
# Maps expected sheet name -> list of common variations (case-insensitive matching)
SHEET_NAME_ALIASES = {
    'vCluster': ['vCluster', 'vclusters', 'Cluster', 'Clusters', 'vClusters'],
    'vHosts': ['vHosts', 'vHost', 'Hosts', 'Host', 'ESXi', 'ESXiHosts'],
    'vInfo': ['vInfo', 'vinfo', 'VMInfo', 'VM Info', 'VMs', 'VM'],
    'vCPU': ['vCPU', 'vcpu', 'CPU', 'VMcpu', 'VM CPU'],
    'vMemory': ['vMemory', 'vmemory', 'Memory', 'VMmemory', 'VM Memory', 'vMem'],
    'vDisk': ['vDisk', 'vdisk', 'Disk', 'VMdisk', 'VM Disk', 'Disks'],
    'vPartition': ['vPartition', 'vpartition', 'Partition', 'Partitions', 'VMpartition'],
}

# ============================================================================
# EXCEL EXTRACTION
# ============================================================================

def find_sheet_match(sheet_names, expected_name):
    """
    Find a matching sheet name from the available sheets using aliases.
    
    Args:
        sheet_names: List of actual sheet names in the Excel file
        expected_name: The canonical/expected sheet name
    
    Returns:
        Tuple of (actual_sheet_name, canonical_name) or (None, None) if not found
    """
    aliases = SHEET_NAME_ALIASES.get(expected_name, [expected_name])
    
    # First, try exact match (case-sensitive)
    for alias in aliases:
        if alias in sheet_names:
            return alias, expected_name
    
    # Second, try case-insensitive match
    sheet_names_lower = {s.lower(): s for s in sheet_names}
    for alias in aliases:
        if alias.lower() in sheet_names_lower:
            return sheet_names_lower[alias.lower()], expected_name
    
    return None, None


def extract_excel_sheets(excel_path, output_dir):
    """
    Extract all sheets from a collector Excel file to individual CSVs.
    Uses sheet name aliasing to handle different collector versions.
    
    Returns:
        Tuple of (success: bool, sheet_mapping: dict)
        sheet_mapping maps canonical names to actual extracted names
    """
    print(f"  Extracting sheets from {excel_path.name}...")
    
    try:
        xl = pd.ExcelFile(excel_path)
        sheet_names = xl.sheet_names
        print(f"    Found {len(sheet_names)} sheets: {', '.join(sheet_names[:5])}{'...' if len(sheet_names) > 5 else ''}")
        
        # Required sheets for processing (canonical names)
        required = ['vCluster', 'vHosts', 'vInfo', 'vCPU', 'vMemory', 'vDisk', 'vPartition']
        
        # Build mapping of canonical name -> actual sheet name
        sheet_mapping = {}
        for req in required:
            actual_name, canonical_name = find_sheet_match(sheet_names, req)
            if actual_name:
                sheet_mapping[canonical_name] = actual_name
        
        # Report any aliases used
        aliases_used = [(canon, actual) for canon, actual in sheet_mapping.items() if canon != actual]
        if aliases_used:
            print(f"    Sheet name aliases applied:")
            for canon, actual in aliases_used:
                print(f"      '{actual}' -> '{canon}'")
        
        # Check for missing required sheets
        missing = [s for s in required if s not in sheet_mapping]
        if missing:
            print(f"    ERROR: Missing required sheets: {', '.join(missing)}")
            print(f"    Available sheets: {', '.join(sheet_names)}")
            print(f"    Tip: Ensure your collector file has sheets for: Clusters, Hosts, VM Info, CPU, Memory, Disk, Partition")
            return False, {}
        
        # Extract all sheets - use canonical names for required sheets
        extracted = []
        for sheet in sheet_names:
            df = pd.read_excel(xl, sheet_name=sheet)
            
            # Check if this sheet should be saved with a canonical name
            canonical_name = None
            for canon, actual in sheet_mapping.items():
                if actual == sheet:
                    canonical_name = canon
                    break
            
            # Save with canonical name if it's a required sheet, otherwise use original name
            save_name = canonical_name if canonical_name else sheet
            csv_path = output_dir / f"{save_name}.csv"
            df.to_csv(csv_path, index=False)
            extracted.append(save_name)
        
        print(f"    Extracted {len(extracted)} sheets to CSVs")
        return True, sheet_mapping
        
    except Exception as e:
        print(f"    ERROR extracting Excel: {e}")
        import traceback
        traceback.print_exc()
        return False, {}

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def mib_to_tib(mib):
    """Convert MiB to TiB (binary)"""
    if pd.isna(mib):
        return 0
    return mib / (1024 * 1024)

def mib_to_gib(mib):
    """Convert MiB to GiB (binary)"""
    if pd.isna(mib):
        return 0
    return mib / 1024

def mib_to_mb(mib):
    """Convert MiB to MB (decimal, 1 MB = 1000^2 bytes)"""
    if pd.isna(mib):
        return 0
    return (mib * 1024 * 1024) / (1000 ** 2)

def mib_to_tb(mib):
    """Convert MiB to TB (decimal)"""
    if pd.isna(mib):
        return 0
    return (mib * 1024 * 1024) / (1000 ** 4)

def gb_to_gib(gb):
    """Convert GB (decimal) to GiB (binary)"""
    if pd.isna(gb):
        return 0
    return gb * (1000 ** 3) / (1024 ** 3)

def load_csv(filepath, name):
    """Load CSV with error handling and report row count"""
    try:
        if not Path(filepath).exists():
            print(f"ERROR loading {name}: File not found at {filepath}")
            return pd.DataFrame()
        df = pd.read_csv(filepath)
        df = df.replace('', pd.NA)
        return df
    except Exception as e:
        print(f"ERROR loading {name}: {e}")
        return pd.DataFrame()

def safe_round(val, decimals=2):
    """Safely round a value, returning 0 if not numeric"""
    try:
        return round(float(val), decimals)
    except:
        return 0

def normalize_column_names(df, column_mappings):
    """
    Normalize column names by renaming variations to expected names.
    
    Args:
        df: DataFrame to normalize
        column_mappings: Dict mapping expected_name -> list of possible variations
    
    Returns:
        DataFrame with renamed columns
    """
    rename_map = {}
    for expected_name, variations in column_mappings.items():
        if expected_name not in df.columns:
            for variation in variations:
                if variation in df.columns:
                    rename_map[variation] = expected_name
                    break
    if rename_map:
        df = df.rename(columns=rename_map)
    return df

def get_column_safe(df, column_name, default=0):
    """
    Safely get a column from DataFrame, returning default if not present.
    
    Args:
        df: DataFrame to access
        column_name: Name of column to retrieve
        default: Default value if column doesn't exist
    
    Returns:
        Column Series or Series of default values
    """
    if column_name in df.columns:
        return df[column_name]
    return pd.Series([default] * len(df), index=df.index)


# ============================================================================
# VALIDATION FUNCTIONS
# ============================================================================

class DataValidationError(Exception):
    """Raised when required data is missing or invalid."""
    pass


def validate_dataframe(df, name, required_columns, optional_columns=None):
    """
    Validate a DataFrame has required columns and report status.
    
    Args:
        df: DataFrame to validate
        name: Name of the data source (for error messages)
        required_columns: List of column names that must exist
        optional_columns: List of columns that are nice to have
    
    Returns:
        Tuple of (is_valid: bool, missing_required: list, missing_optional: list)
    """
    if df is None or df.empty:
        return False, required_columns, optional_columns or []
    
    actual_columns = set(df.columns)
    
    missing_required = [col for col in required_columns if col not in actual_columns]
    missing_optional = [col for col in (optional_columns or []) if col not in actual_columns]
    
    is_valid = len(missing_required) == 0
    
    return is_valid, missing_required, missing_optional


def diagnose_column_mismatch(df, name, expected_columns):
    """
    Provide diagnostic information when columns don't match expectations.
    
    Args:
        df: DataFrame with unexpected columns
        name: Name of the data source
        expected_columns: List of expected column names
    
    Returns:
        String with diagnostic information
    """
    actual = set(df.columns) if not df.empty else set()
    expected = set(expected_columns)
    
    missing = expected - actual
    extra = actual - expected
    
    lines = [f"\n  Column mismatch in {name}:"]
    lines.append(f"    Expected: {', '.join(sorted(expected))}")
    lines.append(f"    Found: {', '.join(sorted(actual)) if actual else '(empty DataFrame)'}")
    
    if missing:
        lines.append(f"    Missing: {', '.join(sorted(missing))}")
    
    # Suggest possible matches for missing columns
    if missing and actual:
        lines.append(f"    Possible matches:")
        for miss in missing:
            # Find columns that might be similar
            miss_lower = miss.lower().replace(' ', '').replace('_', '')
            for act in actual:
                act_lower = act.lower().replace(' ', '').replace('_', '')
                if miss_lower in act_lower or act_lower in miss_lower:
                    lines.append(f"      '{miss}' might be '{act}'")
    
    return '\n'.join(lines)


def load_and_validate_csv(filepath, name, column_mappings, required_columns):
    """
    Load a CSV file, normalize column names, and validate required columns exist.
    
    Args:
        filepath: Path to CSV file
        name: Name for error messages
        column_mappings: Dict of column name mappings to apply
        required_columns: List of required columns after normalization
    
    Returns:
        Tuple of (DataFrame, validation_messages: list)
    """
    messages = []
    
    # Load the CSV
    try:
        if not filepath.exists():
            messages.append(f"ERROR: File not found: {filepath}")
            return pd.DataFrame(), messages
        
        df = pd.read_csv(filepath)
        df = df.replace('', pd.NA)
        
        if df.empty:
            messages.append(f"WARNING: {name} is empty (0 rows)")
            return df, messages
            
    except Exception as e:
        messages.append(f"ERROR loading {name}: {e}")
        return pd.DataFrame(), messages
    
    # Store original columns for diagnostics
    original_columns = list(df.columns)
    
    # Normalize column names
    if column_mappings:
        df = normalize_column_names(df, column_mappings)
        
        # Report any normalizations applied
        renamed = []
        for new_name in df.columns:
            if new_name not in original_columns and new_name in column_mappings:
                for old_name in column_mappings[new_name]:
                    if old_name in original_columns:
                        renamed.append(f"'{old_name}' -> '{new_name}'")
                        break
        if renamed:
            messages.append(f"  Column names normalized in {name}: {', '.join(renamed)}")
    
    # Validate required columns
    is_valid, missing_req, _ = validate_dataframe(df, name, required_columns)
    
    if not is_valid:
        messages.append(f"ERROR: {name} missing required columns: {', '.join(missing_req)}")
        messages.append(diagnose_column_mismatch(df, name, required_columns))
    
    return df, messages

# Column name mappings - maps expected name to possible variations
# Applied to normalize column names across different collector versions
VHOSTS_COLUMN_MAPPINGS = {
    'MOID': ['MOID', 'Host MOID', 'HostMOID', 'Host_MOID', 'MoRef', 'MoRefId', 'Object ID', 'ObjectID'],
    'Host': ['Host', 'Hostname', 'Host Name', 'ESXi Host', 'Server', 'Name'],
    'Cluster': ['Cluster', 'Cluster Name', 'ClusterName', 'Parent Cluster', 'ParentCluster'],
    'Datacenter': ['Datacenter', 'Datacenter MOID', 'DatacenterMOID', 'DC', 'DC MOID'],
    'Free Space (MiB)': ['Freespace (MiB)', 'FreeSpace (MiB)', 'Free Space(MiB)', 'Freespace(MiB)', 'Free Space MiB', 'FreespaceMiB'],
    'Service Tag': ['ServiceTag', 'Service_Tag', 'Service tag', 'Serial Number', 'Serial', 'SerialNumber', 'Asset Tag', 'Serial number'],
    'Maintenance Mode': ['MaintenanceMode', 'Maintenance_Mode', 'Maint Mode', 'MaintMode', 'In Maintenance', 'in Maintenance Mode'],
    'CPUs': ['CPUs', 'CPU Sockets', 'Sockets', 'NumCPU', 'Num CPUs', 'CPU Count', '# CPU', '#CPU'],
    'CPU Cores': ['CPU Cores', 'Cores', 'TotalCores', 'Total Cores', 'NumCores', 'Num Cores', '# Cores', '#Cores', 'NumCpuCores'],
    'Cores per CPU': ['Cores per CPU', 'CoresPerCPU', 'Cores/CPU', 'CoresPerSocket'],
    'Memory Size': ['Memory Size', 'MemorySize', 'Memory (GB)', 'RAM', 'RAM (GB)', 'Memory GB', 'Mem Size', '# Memory', '#Memory', 'vRAM'],
    'VMs': ['VMs', 'VM Count', 'VMCount', 'NumVMs', 'Num VMs', 'Virtual Machines', '# VMs', '#VMs', '# VMs total'],
    'NICs': ['NICs', 'NIC Count', 'NICCount', 'Network Adapters', 'NumNICs', '# NICs', '#NICs'],
    'CPU Speed': ['CPU Speed', 'CPUSpeed', 'CPU MHz', 'CPU Frequency', 'ProcessorSpeed', 'Speed'],
    'Capacity (MiB)': ['Capacity (MiB)', 'Capacity MiB', 'CapacityMiB', 'Storage Capacity', 'Datastore Capacity'],
    'Consumed (MiB)': ['Consumed (MiB)', 'Consumed MiB', 'ConsumedMiB', 'Storage Used', 'Used Space'],
    'GPU Count': ['GPU Count', 'GPUCount', 'GPUs', 'NumGPUs', 'GPU'],
    'GPU Memory Size (MiB)': ['GPU Memory Size (MiB)', 'GPU Memory', 'GPUMemory', 'vGPU Memory'],
    'Model': ['Model', 'Hardware Model', 'Server Model', 'HW Model'],
    'Vendor': ['Vendor', 'Manufacturer', 'Make', 'HW Vendor'],
    'CPU Model': ['CPU Model', 'Processor', 'Processor Model', 'CPUModel', 'CPU Type'],
    'Hypervisor': ['Hypervisor', 'ESXi Version', 'Version', 'Host Version', 'ESX Version'],
    'BIOS': ['BIOS', 'BIOS Version', 'BIOSVersion', 'Firmware'],
}

VINFO_COLUMN_MAPPINGS = {
    'UUID': ['UUID', 'VM UUID', 'VMUUID', 'VM Id', 'VMId', 'Object ID', 'ObjectID'],
    'Host MOID': ['Host MOID', 'HostMOID', 'Host_MOID', 'Host MoRef', 'HostId'],
    'Host': ['Host', 'ESXi Host', 'ESX Host', 'Host Name', 'Hostname', 'Server'],
}

VCPU_COLUMN_MAPPINGS = {
    'UUID': ['UUID', 'VM UUID', 'VMUUID', 'Id', 'VM Id', 'VMId', 'Object ID', 'ObjectID'],
    'vCPUs': ['vCPUs', 'CPUs', 'NumCPU', 'Num vCPU', 'CPU Count', 'vCPU', '# vCPUs', 'Sockets'],
}

VMEMORY_COLUMN_MAPPINGS = {
    'UUID': ['UUID', 'VM UUID', 'VMUUID', 'Id', 'VM Id', 'VMId', 'Object ID', 'ObjectID'],
    'Size (MiB)': ['Size (MiB)', 'Size MiB', 'SizeMiB', 'Memory', 'Memory (MiB)', 'RAM MiB', 'Memory MiB'],
}

VDISK_COLUMN_MAPPINGS = {
    'UUID': ['UUID', 'VM UUID', 'VMUUID', 'Id', 'VM Id', 'VMId', 'Object ID', 'ObjectID'],
    'Capacity (MiB)': ['Capacity (MiB)', 'Capacity MiB', 'CapacityMiB', 'Size', 'Disk Size', 'Size MiB'],
}

VPARTITION_COLUMN_MAPPINGS = {
    'UUID': ['UUID', 'VM UUID', 'VMUUID', 'Id', 'VM Id', 'VMId', 'Object ID', 'ObjectID'],
    'Consumed (MiB)': ['Consumed (MiB)', 'Consumed MiB', 'ConsumedMiB', 'Used', 'Used Space'],
    'Capacity (MiB)': ['Capacity (MiB)', 'Capacity MiB', 'CapacityMiB', 'Size', 'Partition Size'],
}

VCLUSTER_COLUMN_MAPPINGS = {
    'Datacenter': ['Datacenter', 'Datacenter MOID', 'DatacenterMOID', 'DC', 'DC MOID', 'VI SDK Server'],
    'MOID': ['MOID', 'Cluster MOID', 'ClusterMOID', 'Cluster_MOID', 'MoRef', 'Object ID', 'ObjectID'],
    'Cluster Name': ['Cluster Name', 'ClusterName', 'Name', 'Cluster', 'ClusterTitle'],
}

# ============================================================================
# FILE SAVING WITH ERROR HANDLING
# ============================================================================

def save_file_with_retry(save_func, filepath, max_retries=3):
    """
    Attempt to save a file with retry logic for permission errors.
    If file is locked, tries alternative filenames with timestamps.
    
    Args:
        save_func: Callable that takes filepath and saves the file
        filepath: Path object or string for the target file
        max_retries: Number of alternative filenames to try
    
    Returns:
        Path: The actual path where the file was saved
    
    Raises:
        PermissionError: If all retries fail
    """
    from datetime import datetime
    filepath = Path(filepath)
    
    # First attempt: try the original filename
    try:
        save_func(filepath)
        return filepath
    except PermissionError as e:
        print(f"  WARNING: Cannot save to {filepath.name} - file may be open in another application.")
    
    # Retry with alternative filenames
    for retry in range(1, max_retries + 1):
        timestamp = datetime.now().strftime("%H%M%S")
        alt_name = f"{filepath.stem}_{timestamp}{filepath.suffix}"
        alt_path = filepath.parent / alt_name
        
        try:
            print(f"  Trying alternative filename: {alt_name}")
            save_func(alt_path)
            print(f"  Successfully saved to: {alt_path}")
            return alt_path
        except PermissionError:
            if retry < max_retries:
                import time
                time.sleep(0.5)  # Brief pause before next retry
                continue
    
    # All retries failed
    raise PermissionError(
        f"Cannot save file - all attempts failed.\n"
        f"  Target: {filepath}\n"
        f"  Please close the file if it's open in Excel or another application,\n"
        f"  then run the script again."
    )

def save_text_file_safe(filepath, content):
    """
    Save text content to a file with error handling.
    
    Args:
        filepath: Path to save to
        content: String content to write
    
    Returns:
        Path: The actual path where the file was saved, or None if failed
    """
    def do_save(path):
        with open(path, 'w', encoding='utf-8') as f:
            f.write(content)
    
    try:
        return save_file_with_retry(do_save, filepath)
    except PermissionError as e:
        print(f"  WARNING: Could not save validation report: {e}")
        return None

# ============================================================================
# EXCEL STYLING
# ============================================================================

def get_styles():
    """Return style definitions for Excel formatting"""
    if not OPENPYXL_AVAILABLE:
        return {}
    
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    return {
        'cluster_title': Font(bold=True, size=14, color='FFFFFF'),
        'cluster_title_fill': PatternFill(start_color='2F5496', end_color='2F5496', fill_type='solid'),
        'summary_label': Font(bold=True, size=10),
        'summary_value': Font(size=10),
        'summary_fill': PatternFill(start_color='D6DCE4', end_color='D6DCE4', fill_type='solid'),
        'host_header': Font(bold=True, size=10, color='FFFFFF'),
        'host_header_fill': PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid'),
        'host_data': Font(size=9),
        'host_alt_fill': PatternFill(start_color='E8EBF0', end_color='E8EBF0', fill_type='solid'),
        'border': thin_border,
        'center': Alignment(horizontal='center', vertical='center'),
        'left': Alignment(horizontal='left', vertical='center'),
        'right': Alignment(horizontal='right', vertical='center'),
    }

# ============================================================================
# MAIN PROCESSING
# ============================================================================

def process_collector(input_dir, output_prefix, output_excel, output_csv, validation_report):
    """
    Process a single collector's CSV files and generate outputs.
    
    Args:
        input_dir: Path to directory containing CSV files
        output_prefix: Prefix for output identification
        output_excel: Path TO output Excel file
        output_csv: Path to output CSV file (fallback)
        validation_report: Path to validation report
    """
    validation_lines = []
    validation_lines.append("=" * 80)
    validation_lines.append("VMware Cluster Summary - Validation Report")
    validation_lines.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    validation_lines.append("=" * 80)
    validation_lines.append("")
    
    # -------------------------------------------------------------------------
    # 1. LOAD ALL CSVs WITH VALIDATION
    # -------------------------------------------------------------------------
    print("Loading CSV files...")
    all_messages = []
    has_critical_errors = False
    
    # Load and validate vCluster
    vcluster, msgs = load_and_validate_csv(
        input_dir / "vCluster.csv", "vCluster",
        VCLUSTER_COLUMN_MAPPINGS,
        required_columns=['Datacenter', 'MOID', 'Cluster Name']
    )
    all_messages.extend(msgs)
    if vcluster.empty:
        has_critical_errors = True
    
    # Load and validate vHosts
    vhosts, msgs = load_and_validate_csv(
        input_dir / "vHosts.csv", "vHosts",
        VHOSTS_COLUMN_MAPPINGS,
        required_columns=['MOID', 'Cluster']
    )
    all_messages.extend(msgs)
    if vhosts.empty or 'MOID' not in vhosts.columns or 'Cluster' not in vhosts.columns:
        has_critical_errors = True
        print("  CRITICAL: vHosts missing required MOID or Cluster columns")
    
    # Load and validate vInfo - Host column can be either 'Host MOID' or 'Host' (hostname)
    vinfo, msgs = load_and_validate_csv(
        input_dir / "vInfo.csv", "vInfo",
        VINFO_COLUMN_MAPPINGS,
        required_columns=['UUID']  # Host/Host MOID checked separately, either works
    )
    # Filter out the "missing Host MOID" error if we have a Host column
    if 'Host' in vinfo.columns or 'Host MOID' in vinfo.columns:
        msgs = [m for m in msgs if 'Host MOID' not in m]
    all_messages.extend(msgs)
    if vinfo.empty or 'UUID' not in vinfo.columns:
        has_critical_errors = True
    elif 'Host' not in vinfo.columns and 'Host MOID' not in vinfo.columns:
        has_critical_errors = True
        print("  CRITICAL: vInfo missing both 'Host' and 'Host MOID' columns - cannot map VMs to hosts")
    
    # Load and validate vCPU
    vcpu, msgs = load_and_validate_csv(
        input_dir / "vCPU.csv", "vCPU",
        VCPU_COLUMN_MAPPINGS,
        required_columns=['UUID', 'vCPUs']
    )
    all_messages.extend(msgs)
    
    # Load and validate vMemory
    vmemory, msgs = load_and_validate_csv(
        input_dir / "vMemory.csv", "vMemory",
        VMEMORY_COLUMN_MAPPINGS,
        required_columns=['UUID', 'Size (MiB)']
    )
    all_messages.extend(msgs)
    
    # Load and validate vDisk
    vdisk, msgs = load_and_validate_csv(
        input_dir / "vDisk.csv", "vDisk",
        VDISK_COLUMN_MAPPINGS,
        required_columns=['UUID', 'Capacity (MiB)']
    )
    all_messages.extend(msgs)
    
    # Load and validate vPartition
    vpartition, msgs = load_and_validate_csv(
        input_dir / "vPartition.csv", "vPartition",
        VPARTITION_COLUMN_MAPPINGS,
        required_columns=['UUID']
    )
    all_messages.extend(msgs)
    
    # Print all validation messages
    for msg in all_messages:
        print(msg)
        validation_lines.append(msg)
    
    # Check for critical errors before proceeding
    if has_critical_errors:
        error_msg = (
            "\n" + "=" * 70 + "\n"
            "CRITICAL ERROR: Cannot process collector due to missing data.\n"
            "Please check the validation messages above.\n"
            "\n"
            "Common causes:\n"
            "  1. Sheet names don't match expected patterns (vHost vs vHosts, etc.)\n"
            "  2. Column names differ from expected (check column normalization)\n"
            "  3. Empty or corrupt sheets in the collector file\n"
            "\n"
            "The validation report has been saved with diagnostic details.\n"
            + "=" * 70
        )
        print(error_msg)
        validation_lines.append(error_msg)
        
        # Save partial validation report
        validation_content = '\n'.join(validation_lines)
        save_text_file_safe(validation_report, validation_content)
        raise DataValidationError("Missing required data - see validation report")
    
    validation_lines.append("")
    validation_lines.append("CSV Row Counts:")
    validation_lines.append("-" * 50)
    validation_lines.append(f"  vCluster.csv:    {len(vcluster):>6} rows")
    validation_lines.append(f"  vHosts.csv:      {len(vhosts):>6} rows")
    validation_lines.append(f"  vInfo.csv:       {len(vinfo):>6} rows")
    validation_lines.append(f"  vCPU.csv:        {len(vcpu):>6} rows")
    validation_lines.append(f"  vMemory.csv:     {len(vmemory):>6} rows")
    validation_lines.append(f"  vDisk.csv:       {len(vdisk):>6} rows")
    validation_lines.append(f"  vPartition.csv:  {len(vpartition):>6} rows")
    validation_lines.append("")
    
    # -------------------------------------------------------------------------
    # 2. BUILD HOST -> CLUSTER MAPPING
    # -------------------------------------------------------------------------
    print("Building host -> cluster mapping...")
    
    # Determine the best join key for host -> cluster mapping
    # Option 1: Cluster column contains MOID (starts with 'domain-' or similar)
    # Option 2: Cluster column contains cluster NAME (matches vCluster 'Cluster Name')
    
    # Check what type of value is in the Cluster column
    cluster_sample = vhosts['Cluster'].dropna().iloc[0] if not vhosts['Cluster'].dropna().empty else ''
    cluster_is_moid = str(cluster_sample).startswith('domain-') or str(cluster_sample).startswith('ClusterComputeResource')
    
    if cluster_is_moid:
        print("  Using Cluster MOID for host->cluster mapping")
        cluster_key = 'MOID'
    else:
        print("  Using Cluster NAME for host->cluster mapping (joining to vCluster.Cluster Name)")
        cluster_key = 'Cluster Name'
    
    # Build the cluster lookup from vCluster
    # Maps cluster identifier (MOID or Name) to the cluster MOID for final output
    if cluster_key == 'MOID':
        cluster_lookup = vcluster[['MOID', 'Cluster Name', 'Datacenter']].copy()
        cluster_lookup.columns = ['Cluster_MOID', 'Cluster_Name', 'Datacenter_MOID']
    else:
        # Join by cluster name - create mapping from name to MOID
        cluster_lookup = vcluster[['Cluster Name', 'MOID', 'Datacenter']].copy()
        cluster_lookup.columns = ['Cluster_Name', 'Cluster_MOID', 'Datacenter_MOID']
    
    # Build host -> cluster mapping using appropriate key
    if 'Host' in vhosts.columns and 'MOID' in vhosts.columns:
        # We have both hostname and MOID - keep both for flexible joining later
        host_to_cluster = vhosts[['Host', 'MOID', 'Cluster']].copy()
        host_to_cluster.columns = ['Hostname', 'Host_MOID', 'Host_Cluster_Key']
    elif 'Host' in vhosts.columns:
        host_to_cluster = vhosts[['Host', 'Cluster']].copy()
        host_to_cluster.columns = ['Hostname', 'Host_Cluster_Key']
        host_to_cluster['Host_MOID'] = host_to_cluster['Hostname']  # Use hostname as ID
    else:
        host_to_cluster = vhosts[['MOID', 'Cluster']].copy()
        host_to_cluster.columns = ['Host_MOID', 'Host_Cluster_Key']
        host_to_cluster['Hostname'] = host_to_cluster['Host_MOID']
    
    host_to_cluster = host_to_cluster.drop_duplicates(subset=['Hostname'])
    
    # Now map host's cluster key to the actual cluster MOID
    if cluster_key == 'MOID':
        host_to_cluster = host_to_cluster.merge(
            cluster_lookup[['Cluster_MOID', 'Cluster_Name']],
            left_on='Host_Cluster_Key',
            right_on='Cluster_MOID',
            how='left'
        )
    else:
        # Join by cluster name
        host_to_cluster = host_to_cluster.merge(
            cluster_lookup[['Cluster_Name', 'Cluster_MOID']],
            left_on='Host_Cluster_Key',
            right_on='Cluster_Name',
            how='left'
        )
    
    print(f"  Mapped {host_to_cluster['Cluster_MOID'].notna().sum()}/{len(host_to_cluster)} hosts to clusters")
    
    # -------------------------------------------------------------------------
    # 3. BUILD VM -> CLUSTER MAPPING
    # -------------------------------------------------------------------------
    print("Building VM -> cluster mapping...")
    
    # Determine best join key for VM -> Host mapping
    # Option 1: vInfo has 'Host MOID' that matches vHosts 'MOID'
    # Option 2: vInfo has 'Host' (hostname) that matches vHosts 'Host' (hostname)
    
    if 'Host MOID' in vinfo.columns and 'MOID' in vhosts.columns:
        # Check if Host MOID looks like actual MOID values
        host_moid_sample = vinfo['Host MOID'].dropna().iloc[0] if not vinfo['Host MOID'].dropna().empty else ''
        if str(host_moid_sample).startswith('host-'):
            print("  Using Host MOID for VM->host mapping")
            vm_to_host = vinfo[['UUID', 'Host MOID']].copy()
            vm_to_host.columns = ['UUID', 'Host_Key']
            join_on = 'Host_MOID'
        else:
            # Host MOID column exists but contains hostnames
            print("  Using Host (hostname) for VM->host mapping (Host MOID contains hostnames)")
            vm_to_host = vinfo[['UUID', 'Host MOID']].copy()
            vm_to_host.columns = ['UUID', 'Host_Key']
            join_on = 'Hostname'
    elif 'Host' in vinfo.columns:
        print("  Using Host (hostname) for VM->host mapping")
        vm_to_host = vinfo[['UUID', 'Host']].copy()
        vm_to_host.columns = ['UUID', 'Host_Key']
        join_on = 'Hostname'
    else:
        print("  ERROR: No host identifier found in vInfo")
        print(f"  Available columns: {list(vinfo.columns)}")
        raise DataValidationError("Cannot find host identifier in vInfo - need 'Host' or 'Host MOID' column")
    
    vm_to_host = vm_to_host.drop_duplicates(subset=['UUID'])
    
    # Join VM to host to get cluster
    vm_to_cluster = vm_to_host.merge(
        host_to_cluster[['Hostname', 'Host_MOID', 'Cluster_MOID', 'Cluster_Name']],
        left_on='Host_Key',
        right_on=join_on,
        how='left'
    )
    
    vms_missing_cluster = vm_to_cluster['Cluster_MOID'].isna().sum()
    vms_matched = len(vm_to_cluster) - vms_missing_cluster
    print(f"  Mapped {vms_matched}/{len(vm_to_cluster)} VMs to clusters")
    
    validation_lines.append(f"VMs in vInfo: {len(vm_to_cluster)}")
    validation_lines.append(f"VMs successfully mapped to clusters: {vms_matched}")
    validation_lines.append(f"VMs missing cluster mapping: {vms_missing_cluster}")
    if vms_missing_cluster > 0 and vms_missing_cluster < len(vm_to_cluster):
        validation_lines.append(f"  (Some VMs may be on standalone hosts not in a cluster)")
    elif vms_missing_cluster == len(vm_to_cluster):
        validation_lines.append(f"  WARNING: No VMs could be mapped! Check join key compatibility.")
        # Debug info
        validation_lines.append(f"  Join method: {join_on}")
        vm_hosts = set(vm_to_host['Host_Key'].dropna().unique())
        available_hosts = set(host_to_cluster[join_on].dropna().unique()) if join_on in host_to_cluster.columns else set()
        validation_lines.append(f"  Sample VM hosts: {list(vm_hosts)[:3]}")
        validation_lines.append(f"  Sample available hosts: {list(available_hosts)[:3]}")
    validation_lines.append("")
    
    # -------------------------------------------------------------------------
    # 4. MAP VMs TO CLUSTERS
    # -------------------------------------------------------------------------
    print("Mapping VMs to clusters...")
    
    # Safe merges - handle missing UUID columns gracefully
    if 'UUID' in vcpu.columns:
        vcpu_with_cluster = vcpu.merge(vm_to_cluster[['UUID', 'Cluster_MOID']], on='UUID', how='left')
    else:
        print("  WARNING: vCPU missing UUID column, using empty mapping")
        vcpu_with_cluster = vcpu.copy()
        vcpu_with_cluster['Cluster_MOID'] = pd.NA
    
    if 'UUID' in vmemory.columns:
        vmem_with_cluster = vmemory.merge(vm_to_cluster[['UUID', 'Cluster_MOID']], on='UUID', how='left')
    else:
        print("  WARNING: vMemory missing UUID column, using empty mapping")
        vmem_with_cluster = vmemory.copy()
        vmem_with_cluster['Cluster_MOID'] = pd.NA
    
    if 'UUID' in vdisk.columns:
        vdisk_with_cluster = vdisk.merge(vm_to_cluster[['UUID', 'Cluster_MOID']], on='UUID', how='left')
    else:
        print("  WARNING: vDisk missing UUID column, using empty mapping")
        vdisk_with_cluster = vdisk.copy()
        vdisk_with_cluster['Cluster_MOID'] = pd.NA
    
    if 'UUID' in vpartition.columns:
        vpart_with_cluster = vpartition.merge(vm_to_cluster[['UUID', 'Cluster_MOID']], on='UUID', how='left')
    else:
        print("  WARNING: vPartition missing UUID column, using empty mapping")
        vpart_with_cluster = vpartition.copy()
        vpart_with_cluster['Cluster_MOID'] = pd.NA
    
    # -------------------------------------------------------------------------
    # 5. AGGREGATE BY CLUSTER
    # -------------------------------------------------------------------------
    print("Aggregating metrics by cluster...")
    
    try:
        clusters = vcluster[['Datacenter', 'MOID', 'Cluster Name']].copy()
        clusters.columns = ['Datacenter_MOID', 'Cluster_MOID', 'Cluster_Name']
    except KeyError as e:
        print(f"  ERROR: Cannot extract cluster info - {e}")
        print(f"  Available columns in vCluster: {list(vcluster.columns)}")
        raise DataValidationError(f"Missing required columns in vCluster: {e}")
    
    # VM counts
    vm_counts = vm_to_cluster.groupby('Cluster_MOID').size().reset_index(name='Total_VMs')
    clusters = clusters.merge(vm_counts, on='Cluster_MOID', how='left')
    clusters['Total_VMs'] = clusters['Total_VMs'].fillna(0).astype(int)
    
    # vCPUs
    vcpu_agg = vcpu_with_cluster[vcpu_with_cluster['Cluster_MOID'].notna()].copy()
    vcpu_agg['vCPUs'] = pd.to_numeric(vcpu_agg['vCPUs'], errors='coerce').fillna(0)
    vcpu_totals = vcpu_agg.groupby('Cluster_MOID')['vCPUs'].sum().reset_index(name='Total_vCPU')
    clusters = clusters.merge(vcpu_totals, on='Cluster_MOID', how='left')
    clusters['Total_vCPU'] = clusters['Total_vCPU'].fillna(0).astype(int)
    
    # Max vCPU per VM
    vcpu_per_vm = vcpu_agg.groupby(['Cluster_MOID', 'UUID'])['vCPUs'].max().reset_index()
    max_vcpu = vcpu_per_vm.groupby('Cluster_MOID')['vCPUs'].max().reset_index(name='Max_VM_vCPU')
    clusters = clusters.merge(max_vcpu, on='Cluster_MOID', how='left')
    clusters['Max_VM_vCPU'] = clusters['Max_VM_vCPU'].fillna(0).astype(int)
    
    # RAM
    vmem_agg = vmem_with_cluster[vmem_with_cluster['Cluster_MOID'].notna()].copy()
    vmem_agg['Size (MiB)'] = pd.to_numeric(vmem_agg['Size (MiB)'], errors='coerce').fillna(0)
    vmem_per_vm = vmem_agg.groupby(['Cluster_MOID', 'UUID'])['Size (MiB)'].max().reset_index()
    ram_totals = vmem_per_vm.groupby('Cluster_MOID')['Size (MiB)'].sum().reset_index(name='Total_RAM_MiB')
    clusters = clusters.merge(ram_totals, on='Cluster_MOID', how='left')
    clusters['Total_RAM_MiB'] = clusters['Total_RAM_MiB'].fillna(0)
    clusters['Total_RAM_GiB'] = clusters['Total_RAM_MiB'].apply(mib_to_gib)
    clusters['Total_RAM_TiB'] = clusters['Total_RAM_MiB'].apply(mib_to_tib)
    
    # Max RAM per VM
    max_ram = vmem_per_vm.groupby('Cluster_MOID')['Size (MiB)'].max().reset_index(name='Max_VM_RAM_MiB')
    clusters = clusters.merge(max_ram, on='Cluster_MOID', how='left')
    clusters['Max_VM_RAM_MiB'] = clusters['Max_VM_RAM_MiB'].fillna(0)
    clusters['Max_VM_RAM_GiB'] = clusters['Max_VM_RAM_MiB'].apply(mib_to_gib)
    
    # vDisk capacity
    vdisk_agg = vdisk_with_cluster[vdisk_with_cluster['Cluster_MOID'].notna()].copy()
    vdisk_agg['Capacity (MiB)'] = pd.to_numeric(vdisk_agg['Capacity (MiB)'], errors='coerce').fillna(0)
    vdisk_per_vm = vdisk_agg.groupby(['Cluster_MOID', 'UUID'])['Capacity (MiB)'].sum().reset_index(name='VM_Capacity_MiB')
    capacity_totals = vdisk_per_vm.groupby('Cluster_MOID')['VM_Capacity_MiB'].sum().reset_index(name='Capacity_Storage_MiB')
    clusters = clusters.merge(capacity_totals, on='Cluster_MOID', how='left')
    clusters['Capacity_Storage_MiB'] = clusters['Capacity_Storage_MiB'].fillna(0)
    clusters['Capacity_Storage_GiB'] = clusters['Capacity_Storage_MiB'].apply(mib_to_gib)
    clusters['Capacity_Storage_TiB'] = clusters['Capacity_Storage_MiB'].apply(mib_to_tib)
    
    # Max vDisk per VM
    max_cap = vdisk_per_vm.groupby('Cluster_MOID')['VM_Capacity_MiB'].max().reset_index(name='Max_VM_Capacity_MiB')
    clusters = clusters.merge(max_cap, on='Cluster_MOID', how='left')
    clusters['Max_VM_Capacity_MiB'] = clusters['Max_VM_Capacity_MiB'].fillna(0)
    clusters['Max_VM_Capacity_GiB'] = clusters['Max_VM_Capacity_MiB'].apply(mib_to_gib)
    clusters['Max_VM_Capacity_TiB'] = clusters['Max_VM_Capacity_MiB'].apply(mib_to_tib)
    
    # vPartition provisioned/consumed
    vpart_agg = vpart_with_cluster[vpart_with_cluster['Cluster_MOID'].notna()].copy()
    vpart_agg['Consumed (MiB)'] = pd.to_numeric(vpart_agg['Consumed (MiB)'], errors='coerce').fillna(0)
    vpart_agg['Capacity (MiB)'] = pd.to_numeric(vpart_agg['Capacity (MiB)'], errors='coerce').fillna(0)
    vpart_per_vm = vpart_agg.groupby(['Cluster_MOID', 'UUID']).agg({
        'Consumed (MiB)': 'sum', 'Capacity (MiB)': 'sum'
    }).reset_index()
    vpart_per_vm.columns = ['Cluster_MOID', 'UUID', 'VM_Consumed_MiB', 'VM_Provisioned_MiB']
    vpart_totals = vpart_per_vm.groupby('Cluster_MOID').agg({
        'VM_Consumed_MiB': 'sum', 'VM_Provisioned_MiB': 'sum'
    }).reset_index()
    vpart_totals.columns = ['Cluster_MOID', 'Consumed_Storage_MiB', 'Provisioned_Storage_MiB']
    clusters = clusters.merge(vpart_totals, on='Cluster_MOID', how='left')
    clusters['Consumed_Storage_MiB'] = clusters['Consumed_Storage_MiB'].fillna(0)
    clusters['Provisioned_Storage_MiB'] = clusters['Provisioned_Storage_MiB'].fillna(0)
    clusters['Consumed_Storage_GiB'] = clusters['Consumed_Storage_MiB'].apply(mib_to_gib)
    clusters['Consumed_Storage_TiB'] = clusters['Consumed_Storage_MiB'].apply(mib_to_tib)
    clusters['Provisioned_Storage_GiB'] = clusters['Provisioned_Storage_MiB'].apply(mib_to_gib)
    clusters['Provisioned_Storage_TiB'] = clusters['Provisioned_Storage_MiB'].apply(mib_to_tib)
    
    # Max consumed/provisioned per VM
    max_vpart = vpart_per_vm.groupby('Cluster_MOID').agg({
        'VM_Consumed_MiB': 'max', 'VM_Provisioned_MiB': 'max'
    }).reset_index()
    max_vpart.columns = ['Cluster_MOID', 'Max_VM_Consumed_MiB', 'Max_VM_Provisioned_MiB']
    clusters = clusters.merge(max_vpart, on='Cluster_MOID', how='left')
    clusters['Max_VM_Consumed_MiB'] = clusters['Max_VM_Consumed_MiB'].fillna(0)
    clusters['Max_VM_Provisioned_MiB'] = clusters['Max_VM_Provisioned_MiB'].fillna(0)
    clusters['Max_VM_Consumed_GiB'] = clusters['Max_VM_Consumed_MiB'].apply(mib_to_gib)
    clusters['Max_VM_Consumed_TiB'] = clusters['Max_VM_Consumed_MiB'].apply(mib_to_tib)
    clusters['Max_VM_Provisioned_GiB'] = clusters['Max_VM_Provisioned_MiB'].apply(mib_to_gib)
    clusters['Max_VM_Provisioned_TiB'] = clusters['Max_VM_Provisioned_MiB'].apply(mib_to_tib)
    
    # -------------------------------------------------------------------------
    # 6. HOST AGGREGATES FOR CLUSTER SUMMARY
    # -------------------------------------------------------------------------
    print("Calculating host aggregates...")
    
    # Prepare host data with numeric conversions
    host_df = vhosts.copy()
    
    # Normalize column names to handle variations across collector versions
    host_df = normalize_column_names(host_df, VHOSTS_COLUMN_MAPPINGS)
    
    # Safe numeric conversions using get_column_safe for optional columns
    host_df['CPUs'] = pd.to_numeric(get_column_safe(host_df, 'CPUs', 0), errors='coerce').fillna(0).astype(int)
    host_df['CPU Cores'] = pd.to_numeric(get_column_safe(host_df, 'CPU Cores', 0), errors='coerce').fillna(0).astype(int)
    host_df['Cores per CPU'] = pd.to_numeric(get_column_safe(host_df, 'Cores per CPU', 0), errors='coerce').fillna(0).astype(int)
    host_df['Memory Size'] = pd.to_numeric(get_column_safe(host_df, 'Memory Size', 0), errors='coerce').fillna(0)
    host_df['VMs'] = pd.to_numeric(get_column_safe(host_df, 'VMs', 0), errors='coerce').fillna(0).astype(int)
    host_df['NICs'] = pd.to_numeric(get_column_safe(host_df, 'NICs', 0), errors='coerce').fillna(0).astype(int)
    host_df['CPU Speed'] = pd.to_numeric(get_column_safe(host_df, 'CPU Speed', 0), errors='coerce').fillna(0)
    host_df['Capacity (MiB)'] = pd.to_numeric(get_column_safe(host_df, 'Capacity (MiB)', 0), errors='coerce').fillna(0)
    host_df['Consumed (MiB)'] = pd.to_numeric(get_column_safe(host_df, 'Consumed (MiB)', 0), errors='coerce').fillna(0)
    host_df['Free Space (MiB)'] = pd.to_numeric(get_column_safe(host_df, 'Free Space (MiB)', 0), errors='coerce').fillna(0)
    host_df['GPU Count'] = pd.to_numeric(get_column_safe(host_df, 'GPU Count', 0), errors='coerce').fillna(0).astype(int)
    host_df['GPU Memory Size (MiB)'] = pd.to_numeric(get_column_safe(host_df, 'GPU Memory Size (MiB)', 0), errors='coerce').fillna(0)
    
    # Ensure optional columns exist for aggregation (fill with defaults if missing)
    for col in ['Model', 'Vendor', 'CPU Model']:
        if col not in host_df.columns:
            host_df[col] = 'Unknown'
    
    # Add cluster mapping to host_df - map cluster name/MOID to actual Cluster_MOID
    # The 'Cluster' column might contain names or MOIDs depending on the collector version
    if 'Cluster' in host_df.columns:
        cluster_sample = host_df['Cluster'].dropna().iloc[0] if not host_df['Cluster'].dropna().empty else ''
        if str(cluster_sample).startswith('domain-') or str(cluster_sample).startswith('ClusterComputeResource'):
            # Cluster column contains MOIDs
            host_df['Cluster_Join_Key'] = host_df['Cluster']
            join_col = 'Cluster_MOID'
        else:
            # Cluster column contains names - need to map to MOID
            host_df['Cluster_Join_Key'] = host_df['Cluster']
            join_col = 'Cluster_Name'
        
        # Create a lookup from clusters DataFrame
        cluster_map = clusters[['Cluster_MOID', 'Cluster_Name']].drop_duplicates()
        host_df = host_df.merge(
            cluster_map,
            left_on='Cluster_Join_Key',
            right_on=join_col,
            how='left',
            suffixes=('', '_mapped')
        )
        # Ensure Cluster_MOID exists for grouping
        if 'Cluster_MOID_mapped' in host_df.columns:
            host_df['Cluster_MOID'] = host_df['Cluster_MOID_mapped']
        elif 'Cluster_MOID' not in host_df.columns:
            host_df['Cluster_MOID'] = host_df['Cluster_Join_Key']
    
    # Aggregate per cluster using Cluster_MOID
    groupby_col = 'Cluster_MOID' if 'Cluster_MOID' in host_df.columns else 'Cluster'
    host_agg = host_df.groupby(groupby_col).agg({
        'MOID': 'nunique',
        'CPUs': 'sum',
        'CPU Cores': 'sum',
        'Cores per CPU': 'first',
        'Memory Size': 'sum',
        'Capacity (MiB)': 'sum',
        'Consumed (MiB)': 'sum',
        'Free Space (MiB)': 'sum',
    }).reset_index()
    host_agg.columns = ['Cluster_MOID', 'No_Hosts', 'Total_CPUs', 'Total_pCores', 'Cores_per_CPU',
                        'Total_Host_RAM_MiB', 'Host_Capacity_MiB', 'Host_Consumed_MiB', 'Host_Free_MiB']
    
    clusters = clusters.merge(host_agg, on='Cluster_MOID', how='left')
    clusters['No_Hosts'] = clusters['No_Hosts'].fillna(0).astype(int)
    clusters['Total_CPUs'] = clusters['Total_CPUs'].fillna(0).astype(int)
    clusters['Total_pCores'] = clusters['Total_pCores'].fillna(0).astype(int)
    clusters['Cores_per_CPU'] = clusters['Cores_per_CPU'].fillna(0).astype(int)
    clusters['Total_Host_RAM_MiB'] = clusters['Total_Host_RAM_MiB'].fillna(0)
    # Convert MiB to GiB for display (host RAM is typically reported in MiB by collectors)
    clusters['Total_Host_RAM_GiB'] = clusters['Total_Host_RAM_MiB'].apply(mib_to_gib)
    clusters['Total_Host_RAM_TiB'] = clusters['Total_Host_RAM_MiB'].apply(mib_to_tib)
    clusters['Host_Capacity_MiB'] = clusters['Host_Capacity_MiB'].fillna(0)
    clusters['Host_Capacity_TiB'] = clusters['Host_Capacity_MiB'].apply(mib_to_tib)
    clusters['Host_Consumed_MiB'] = clusters['Host_Consumed_MiB'].fillna(0)
    clusters['Host_Consumed_TiB'] = clusters['Host_Consumed_MiB'].apply(mib_to_tib)
    clusters['Host_Free_MiB'] = clusters['Host_Free_MiB'].fillna(0)
    clusters['Host_Free_TiB'] = clusters['Host_Free_MiB'].apply(mib_to_tib)
    
    # Get distinct hardware models per cluster (use Cluster_MOID from the mapped host_df)
    models_per_cluster = host_df.groupby(groupby_col)['Model'].apply(
        lambda x: ' | '.join(sorted(x.dropna().unique()))
    ).reset_index()
    models_per_cluster.columns = ['Cluster_MOID', 'Host_Models']
    clusters = clusters.merge(models_per_cluster, on='Cluster_MOID', how='left')
    
    # Get distinct vendors per cluster
    vendors_per_cluster = host_df.groupby(groupby_col)['Vendor'].apply(
        lambda x: ' | '.join(sorted(x.dropna().unique()))
    ).reset_index()
    vendors_per_cluster.columns = ['Cluster_MOID', 'Vendors']
    clusters = clusters.merge(vendors_per_cluster, on='Cluster_MOID', how='left')
    
    # Get distinct CPU models per cluster
    cpu_models_per_cluster = host_df.groupby(groupby_col)['CPU Model'].apply(
        lambda x: ' | '.join(sorted(x.dropna().unique()))
    ).reset_index()
    cpu_models_per_cluster.columns = ['Cluster_MOID', 'CPU_Models']
    clusters = clusters.merge(cpu_models_per_cluster, on='Cluster_MOID', how='left')
    
    # vCPU:pCore ratio
    clusters['vCPU_pCore_Ratio'] = clusters.apply(
        lambda row: round(row['Total_vCPU'] / row['Total_pCores'], 2) if row['Total_pCores'] > 0 else 0,
        axis=1
    )
    
    # Avg vCPU per VM
    clusters['Avg_vCPU_per_VM'] = clusters.apply(
        lambda row: round(row['Total_vCPU'] / row['Total_VMs'], 2) if row['Total_VMs'] > 0 else 0,
        axis=1
    )
    
    # Sort clusters
    clusters = clusters.sort_values(['Datacenter_MOID', 'Cluster_MOID'])
    
    # -------------------------------------------------------------------------
    # 7. DEFINE HOST TABLE COLUMNS
    # -------------------------------------------------------------------------
    
    host_table_columns = [
        ('Host MOID', 'MOID'),
        ('Service Tag', 'Service Tag'),
        ('Vendor', 'Vendor'),
        ('Model', 'Model'),
        ('BIOS', 'BIOS'),
        ('Hypervisor', 'Hypervisor'),
        ('Maint Mode', 'Maintenance Mode'),
        ('CPU Sockets', 'CPUs'),
        ('CPU Model', 'CPU Model'),
        ('Total Cores', 'CPU Cores'),
        ('Cores/CPU', 'Cores per CPU'),
        ('CPU MHz', 'CPU Speed'),
        ('RAM (GB)', 'Memory Size'),
        ('NICs', 'NICs'),
        ('GPUs', 'GPU Count'),
        ('Storage Cap (TiB)', None),
        ('Storage Used (TiB)', None),
        ('Storage Free (TiB)', None),
        ('VMs', 'VMs'),
    ]
    
    # -------------------------------------------------------------------------
    # 8. BUILD EXCEL OUTPUT
    # -------------------------------------------------------------------------
    
    if OPENPYXL_AVAILABLE:
        print("Creating formatted Excel workbook...")
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Cluster Summary"
        
        styles = get_styles()
        current_row = 1
        
        # ===== AGGREGATE SUMMARY TABLE (ALL CLUSTERS) =====
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=10)
        title_cell = ws.cell(row=current_row, column=1, value="TOTAL SUMMARY (ALL CLUSTERS)")
        title_cell.font = Font(bold=True, size=14, color="FFFFFF")
        title_cell.fill = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
        title_cell.alignment = styles['left']
        current_row += 1
        
        # Aggregate totals
        total_hosts = clusters['No_Hosts'].sum()
        total_vms = clusters['Total_VMs'].sum()
        total_vcpu = clusters['Total_vCPU'].sum()
        total_pcores = clusters['Total_pCores'].sum()
        total_host_ram_gib = clusters['Total_Host_RAM_GiB'].sum()
        total_host_ram_tib = clusters['Total_Host_RAM_TiB'].sum()
        total_vm_ram_gib = clusters['Total_RAM_GiB'].sum()
        total_vm_ram_tib = clusters['Total_RAM_TiB'].sum()
        total_provisioned_tib = clusters['Provisioned_Storage_TiB'].sum()
        total_consumed_tib = clusters['Consumed_Storage_TiB'].sum()
        total_capacity_tib = clusters['Capacity_Storage_TiB'].sum()
        total_cluster_capacity_tib = clusters['Host_Capacity_TiB'].sum()
        total_cluster_consumed_tib = clusters['Host_Consumed_TiB'].sum()
        total_cluster_free_tib = clusters['Host_Free_TiB'].sum()
        overall_vcpu_pcore = round(total_vcpu / total_pcores, 2) if total_pcores > 0 else 0
        overall_avg_vcpu_vm = round(total_vcpu / total_vms, 2) if total_vms > 0 else 0
        max_vm_vcpu = clusters['Max_VM_vCPU'].max()
        max_vm_ram_gib = clusters['Max_VM_RAM_GiB'].max()
        max_vm_provisioned_gib = clusters['Max_VM_Provisioned_GiB'].max()
        max_vm_capacity_gib = clusters['Max_VM_Capacity_GiB'].max()
        
        # All unique hardware
        all_vendors = ' | '.join(sorted(set(' | '.join(clusters['Vendors'].dropna()).split(' | '))))
        all_models = ' | '.join(sorted(set(' | '.join(clusters['Host_Models'].dropna()).split(' | '))))
        all_cpus = ' | '.join(sorted(set(' | '.join(clusters['CPU_Models'].dropna()).split(' | '))))
        
        aggregate_data = [
            ('Total Clusters', len(clusters)),
            ('Total Hosts', total_hosts),
            ('Total VMs', total_vms),
            ('Total vCPU', total_vcpu),
            ('Total pCores', total_pcores),
            ('vCPU:pCore Ratio', overall_vcpu_pcore),
            ('Avg vCPU/VM', overall_avg_vcpu_vm),
            ('Total Host RAM (GiB)', safe_round(total_host_ram_gib, 2)),
            ('Total Host RAM (TiB)', safe_round(total_host_ram_tib, 4)),
            ('VM RAM Total (GiB)', safe_round(total_vm_ram_gib, 2)),
            ('VM RAM Total (TiB)', safe_round(total_vm_ram_tib, 4)),
            ('Max VM vCPU', max_vm_vcpu),
            ('Max VM RAM (GiB)', safe_round(max_vm_ram_gib, 2)),
            ('Provisioned Storage (TiB)', safe_round(total_provisioned_tib, 4)),
            ('Consumed Storage (TiB)', safe_round(total_consumed_tib, 4)),
            ('Capacity vDisk (TiB)', safe_round(total_capacity_tib, 4)),
            ('Max VM Provisioned (GiB)', safe_round(max_vm_provisioned_gib, 2)),
            ('Max VM Capacity (GiB)', safe_round(max_vm_capacity_gib, 2)),
            ('Cluster Storage Capacity (TiB)', safe_round(total_cluster_capacity_tib, 4)),
            ('Cluster Storage Consumed (TiB)', safe_round(total_cluster_consumed_tib, 4)),
            ('Cluster Storage Free (TiB)', safe_round(total_cluster_free_tib, 4)),
            ('Hardware Vendors', all_vendors),
            ('Hardware Models', all_models),
            ('CPU Models', all_cpus),
        ]
        
        # Write aggregate summary in 4-column layout
        agg_fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
        for i, (label, value) in enumerate(aggregate_data):
            row_offset = i // 2
            col_offset = (i % 2) * 2
            
            label_cell = ws.cell(row=current_row + row_offset, column=1 + col_offset, value=label)
            label_cell.font = styles['summary_label']
            label_cell.fill = agg_fill
            label_cell.border = styles['border']
            
            value_cell = ws.cell(row=current_row + row_offset, column=2 + col_offset, value=value)
            value_cell.font = styles['summary_value']
            value_cell.border = styles['border']
        
        current_row += (len(aggregate_data) + 1) // 2 + 2  # Extra blank row after summary
        
        # ===== PER-CLUSTER SECTIONS =====
        for idx, cluster in clusters.iterrows():
            cluster_moid = cluster['Cluster_MOID']
            cluster_name = cluster['Cluster_Name']
            
            # ===== CLUSTER TITLE =====
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=10)
            title_cell = ws.cell(row=current_row, column=1, 
                                 value=f"CLUSTER: {cluster_name} ({cluster_moid})")
            title_cell.font = styles['cluster_title']
            title_cell.fill = styles['cluster_title_fill']
            title_cell.alignment = styles['left']
            current_row += 1
            
            # ===== CLUSTER SUMMARY SECTION =====
            summary_data = [
                ('Datacenter MOID', cluster['Datacenter_MOID']),
                ('Cluster MOID', cluster_moid),
                ('Total Hosts', cluster['No_Hosts']),
                ('Total VMs', cluster['Total_VMs']),
                ('Total vCPU', cluster['Total_vCPU']),
                ('Total pCores', cluster['Total_pCores']),
                ('vCPU:pCore Ratio', cluster['vCPU_pCore_Ratio']),
                ('Avg vCPU/VM', cluster['Avg_vCPU_per_VM']),
                ('Total Host RAM (GiB)', safe_round(cluster['Total_Host_RAM_GiB'], 2)),
                ('Total Host RAM (TiB)', safe_round(cluster['Total_Host_RAM_TiB'], 4)),
                ('VM RAM Total (GiB)', safe_round(cluster['Total_RAM_GiB'], 2)),
                ('VM RAM Total (TiB)', safe_round(cluster['Total_RAM_TiB'], 4)),
                ('Max VM vCPU', cluster['Max_VM_vCPU']),
                ('Max VM RAM (GiB)', safe_round(cluster['Max_VM_RAM_GiB'], 2)),
                ('Provisioned Storage (TiB)', safe_round(cluster['Provisioned_Storage_TiB'], 4)),
                ('Consumed Storage (TiB)', safe_round(cluster['Consumed_Storage_TiB'], 4)),
                ('Capacity vDisk (TiB)', safe_round(cluster['Capacity_Storage_TiB'], 4)),
                ('Max VM Provisioned (GiB)', safe_round(cluster['Max_VM_Provisioned_GiB'], 2)),
                ('Max VM Capacity (GiB)', safe_round(cluster['Max_VM_Capacity_GiB'], 2)),
                ('Cluster Storage Capacity (TiB)', safe_round(cluster['Host_Capacity_TiB'], 4)),
                ('Cluster Storage Consumed (TiB)', safe_round(cluster['Host_Consumed_TiB'], 4)),
                ('Cluster Storage Free (TiB)', safe_round(cluster['Host_Free_TiB'], 4)),
                ('Hardware Models', cluster.get('Host_Models', '')),
                ('Vendors', cluster.get('Vendors', '')),
                ('CPU Models', cluster.get('CPU_Models', '')),
            ]
            
            # Write summary in 4-column layout (Label, Value, Label, Value)
            for i, (label, value) in enumerate(summary_data):
                row_offset = i // 2
                col_offset = (i % 2) * 2
                
                label_cell = ws.cell(row=current_row + row_offset, column=1 + col_offset, value=label)
                label_cell.font = styles['summary_label']
                label_cell.fill = styles['summary_fill']
                label_cell.border = styles['border']
                
                value_cell = ws.cell(row=current_row + row_offset, column=2 + col_offset, value=value)
                value_cell.font = styles['summary_value']
                value_cell.border = styles['border']
            
            current_row += (len(summary_data) + 1) // 2 + 1
            
            # ===== HOST DETAILS TABLE =====
            ws.cell(row=current_row, column=1, value="HOST DETAILS").font = Font(bold=True, size=11)
            current_row += 1
            
            # Header row
            for col_idx, (header, _) in enumerate(host_table_columns, 1):
                cell = ws.cell(row=current_row, column=col_idx, value=header)
                cell.font = styles['host_header']
                cell.fill = styles['host_header_fill']
                cell.border = styles['border']
                cell.alignment = styles['center']
            current_row += 1
            
            # Host data rows - filter by Cluster_MOID if available, else by cluster name
            if 'Cluster_MOID' in host_df.columns:
                cluster_hosts = host_df[host_df['Cluster_MOID'] == cluster_moid].copy()
            else:
                # Fall back to matching Cluster column against cluster name
                cluster_hosts = host_df[host_df['Cluster'] == cluster_name].copy()
            cluster_hosts = cluster_hosts.sort_values('MOID') if 'MOID' in cluster_hosts.columns else cluster_hosts
            
            for row_idx, (_, host) in enumerate(cluster_hosts.iterrows()):
                for col_idx, (header, source_col) in enumerate(host_table_columns, 1):
                    if source_col:
                        value = host.get(source_col, '')
                        if pd.isna(value):
                            value = ''
                    else:
                        # Calculated fields
                        if header == 'Storage Cap (TiB)':
                            value = safe_round(mib_to_tib(host['Capacity (MiB)']), 4)
                        elif header == 'Storage Used (TiB)':
                            value = safe_round(mib_to_tib(host['Consumed (MiB)']), 4)
                        elif header == 'Storage Free (TiB)':
                            value = safe_round(mib_to_tib(host['Free Space (MiB)']), 4)
                        else:
                            value = ''
                    
                    cell = ws.cell(row=current_row, column=col_idx, value=value)
                    cell.font = styles['host_data']
                    cell.border = styles['border']
                    
                    # Alternate row coloring
                    if row_idx % 2 == 1:
                        cell.fill = styles['host_alt_fill']
                
                current_row += 1
            
            # Add blank rows between clusters
            current_row += 2
        
        # Auto-size columns
        for col_idx in range(1, len(host_table_columns) + 1):
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = 18
        
        # Make first few columns wider for text content
        ws.column_dimensions['A'].width = 28
        ws.column_dimensions['B'].width = 35
        ws.column_dimensions['C'].width = 12
        ws.column_dimensions['D'].width = 25
        ws.column_dimensions['I'].width = 45  # CPU Model
        
        print(f"Saving Excel file to {output_excel}...")
        
        def save_workbook(path):
            wb.save(path)
        
        try:
            actual_path = save_file_with_retry(save_workbook, output_excel)
            print(f"Output: {actual_path}")
        except PermissionError as e:
            print(f"ERROR: {e}")
            raise
    
    else:
        # Fallback to CSV if openpyxl not available
        print("openpyxl not available - creating simplified CSV output...")
        
        all_columns = ['Row_Type', 'Cluster_Name', 'Cluster_MOID', 'Datacenter_MOID',
                       'Host_MOID', 'Service_Tag', 'Vendor', 'Model', 'CPU_Sockets',
                       'CPU_Model', 'CPU_Cores', 'RAM_GB', 'VMs', 'Total_Hosts',
                       'Total_VMs', 'Total_vCPU', 'Total_pCores', 'vCPU_pCore_Ratio']
        
        output_rows = []
        for idx, cluster in clusters.iterrows():
            cluster_moid = cluster['Cluster_MOID']
            
            # Header row
            output_rows.append({col: col for col in all_columns})
            
            # Cluster summary row
            output_rows.append({
                'Row_Type': 'CLUSTER_SUMMARY',
                'Cluster_Name': cluster['Cluster_Name'],
                'Cluster_MOID': cluster_moid,
                'Datacenter_MOID': cluster['Datacenter_MOID'],
                'Total_Hosts': cluster['No_Hosts'],
                'Total_VMs': cluster['Total_VMs'],
                'Total_vCPU': cluster['Total_vCPU'],
                'Total_pCores': cluster['Total_pCores'],
                'vCPU_pCore_Ratio': cluster['vCPU_pCore_Ratio'],
            })
            
            # Host rows - filter by Cluster_MOID if available, else by cluster name  
            if 'Cluster_MOID' in host_df.columns:
                cluster_hosts = host_df[host_df['Cluster_MOID'] == cluster_moid]
            else:
                cluster_hosts = host_df[host_df['Cluster'] == cluster['Cluster_Name']]
            for _, host in cluster_hosts.iterrows():
                output_rows.append({
                    'Row_Type': 'HOST_DETAIL',
                    'Cluster_Name': cluster['Cluster_Name'],
                    'Cluster_MOID': cluster_moid,
                    'Host_MOID': host['MOID'],
                    'Service_Tag': host.get('Service Tag', ''),
                    'Vendor': host.get('Vendor', ''),
                    'Model': host.get('Model', ''),
                    'CPU_Sockets': host['CPUs'],
                    'CPU_Model': host.get('CPU Model', ''),
                    'CPU_Cores': host['CPU Cores'],
                    'RAM_GB': safe_round(host['Memory Size'], 2),
                    'VMs': host['VMs'],
                })
            
            # Blank separator
            output_rows.append({col: '' for col in all_columns})
        
        def save_csv(path):
            pd.DataFrame(output_rows).to_csv(path, index=False, header=False)
        
        try:
            actual_path = save_file_with_retry(save_csv, output_csv)
            print(f"Output: {actual_path}")
        except PermissionError as e:
            print(f"ERROR: {e}")
            raise
    
    # -------------------------------------------------------------------------
    # 9. VALIDATION REPORT
    # -------------------------------------------------------------------------
    validation_lines.append("=" * 80)
    validation_lines.append("SUMMARY")
    validation_lines.append("=" * 80)
    validation_lines.append(f"Total clusters: {len(clusters)}")
    validation_lines.append(f"Total hosts: {len(vhosts)}")
    validation_lines.append(f"Total VMs: {clusters['Total_VMs'].sum()}")
    validation_lines.append(f"Total vCPUs: {clusters['Total_vCPU'].sum()}")
    validation_lines.append(f"Total Host RAM (GiB): {clusters['Total_Host_RAM_GiB'].sum():.2f}")
    validation_lines.append(f"Total Host RAM (TiB): {clusters['Total_Host_RAM_TiB'].sum():.4f}")
    validation_lines.append("")
    
    validation_lines.append("=" * 80)
    validation_lines.append("HARDWARE BY CLUSTER")
    validation_lines.append("=" * 80)
    for _, row in clusters.iterrows():
        validation_lines.append(f"\n{row['Cluster_Name']} ({row['Cluster_MOID']}):")
        validation_lines.append(f"  Vendors: {row.get('Vendors', 'N/A')}")
        validation_lines.append(f"  Models: {row.get('Host_Models', 'N/A')}")
    
    validation_content = '\n'.join(validation_lines)
    actual_validation_path = save_text_file_safe(validation_report, validation_content)
    if actual_validation_path:
        print(f"Validation: {actual_validation_path}")
    
    print("Done processing this collector!")
    return True


# ============================================================================
# MAIN ENTRY POINT
# ============================================================================

def main():
    """
    Main entry point. Processes:
    1. Excel files from input_files/ folder (extracts sheets to CSVs, then processes)
    2. CSV files in script directory (legacy mode for backward compatibility)
    """
    print("=" * 70)
    print("VMware Cluster Summary Generator")
    print(f"Date: {DATE_STR}")
    print("=" * 70)
    
    processed_count = 0
    
    # Check for Excel collector files in input_files/
    excel_files = list(INPUT_FILES_DIR.glob("*.xlsx")) + list(INPUT_FILES_DIR.glob("*.xls"))
    
    if excel_files:
        print(f"\nFound {len(excel_files)} collector file(s) in input_files/:")
        for f in excel_files:
            print(f"  - {f.name}")
        print()
        
        for excel_file in excel_files:
            print("=" * 70)
            print(f"Processing: {excel_file.name}")
            print("=" * 70)
            
            # Create temp directory for extracted CSVs
            collector_name = excel_file.stem  # filename without extension
            temp_csv_dir = RESULTS_DIR / f"_temp_{collector_name}"
            temp_csv_dir.mkdir(exist_ok=True)
            
            # Extract sheets to CSVs (returns tuple of success, sheet_mapping)
            success, sheet_mapping = extract_excel_sheets(excel_file, temp_csv_dir)
            
            if success:
                # Generate output paths for this collector
                output_excel = RESULTS_DIR / f"cluster_summary_{collector_name}_{DATE_STR}.xlsx"
                output_csv = RESULTS_DIR / f"cluster_summary_{collector_name}_{DATE_STR}.csv"
                validation_report = RESULTS_DIR / f"validation_{collector_name}_{DATE_STR}.txt"
                
                # Process the extracted CSVs
                try:
                    process_collector(
                        input_dir=temp_csv_dir,
                        output_prefix=collector_name,
                        output_excel=output_excel,
                        output_csv=output_csv,
                        validation_report=validation_report
                    )
                    processed_count += 1
                except DataValidationError as e:
                    print(f"VALIDATION ERROR processing {excel_file.name}: {e}")
                    print("  Check the validation report for details.")
                except Exception as e:
                    print(f"ERROR processing {excel_file.name}: {e}")
                    import traceback
                    traceback.print_exc()
            else:
                print(f"  Skipping {excel_file.name} - sheet extraction failed")
                print("  Check that this is a valid collector export with expected sheets")
            
            # Clean up temp CSV directory
            try:
                import shutil
                shutil.rmtree(temp_csv_dir)
                print(f"  Cleaned up temp files")
            except:
                pass
            
            print()
    
    else:
        # Legacy mode: Check for CSVs in script directory
        print("\nNo Excel files in input_files/. Checking for CSVs in script directory...")
        
        if (SCRIPT_DIR / "vCluster.csv").exists():
            print("Found CSV files in script directory (legacy mode)")
            
            output_excel = RESULTS_DIR / f"cluster_summary_{DATE_STR}.xlsx"
            output_csv = RESULTS_DIR / f"cluster_summary_{DATE_STR}.csv"
            validation_report = RESULTS_DIR / f"validation_report_{DATE_STR}.txt"
            
            try:
                process_collector(
                    input_dir=SCRIPT_DIR,
                    output_prefix="local",
                    output_excel=output_excel,
                    output_csv=output_csv,
                    validation_report=validation_report
                )
                processed_count += 1
            except DataValidationError as e:
                print(f"VALIDATION ERROR: {e}")
                print("  Check the validation report for details.")
            except Exception as e:
                print(f"ERROR: {e}")
                import traceback
                traceback.print_exc()
        else:
            print("\nNo input files found!")
            print("Options:")
            print("  1. Place collector Excel files (.xlsx) in: input_files/")
            print("  2. Place CSV files (vCluster.csv, vHosts.csv, etc.) in script directory")
    
    print("\n" + "=" * 70)
    print(f"COMPLETE: Processed {processed_count} collector file(s)")
    print(f"Results saved to: {RESULTS_DIR}")
    print("=" * 70)


if __name__ == "__main__":
    main()
