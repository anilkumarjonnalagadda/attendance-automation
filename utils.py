"""
Utility functions for attendance automation.
Handles CSV loading, email normalization, and duration parsing.
"""

import re
import pandas as pd
from datetime import datetime


# ============================================================
# Field Specification Constants
# ============================================================

REGISTRATION_FIELDS = [
    {'key': 'Email',               'label': 'Email *',              'required': True},
    {'key': 'Training date',       'label': 'Training Date *',      'required': True},
    {'key': 'Attendance - Day 1',  'label': 'Attendance - Day 1',   'required': False},
    {'key': 'Attendance - Day 2',  'label': 'Attendance - Day 2',   'required': False},
    {'key': 'Name',                'label': 'Name',                 'required': False},
    {'key': 'Version',             'label': 'Version / Program',    'required': False},
]

ZOOM_FIELDS = [
    {'key': 'Email',    'label': 'Email *',    'required': True},
    {'key': 'Duration', 'label': 'Duration *', 'required': True},
    {'key': 'Name',     'label': 'Name',       'required': False},
]


# ============================================================
# Core Helpers
# ============================================================

def normalize_email(email):
    """Normalize email: lowercase and strip whitespace."""
    if pd.isna(email) or not isinstance(email, str):
        return ""
    return email.strip().lower()


def parse_duration_to_minutes(duration_value):
    """
    Parse various Zoom duration formats into total minutes.

    Handles:
    - Integer/float (already in minutes): 90, 90.5
    - "HH:MM:SS" format: "1:30:00"
    - "H hours M minutes" / "Hh Mm": "1h 30m", "1 hour 30 minutes"
    - "X mins" / "X minutes": "90 mins", "90 minutes"
    - Raw string number: "90"
    """
    if pd.isna(duration_value):
        return 0

    # If it's already a number
    if isinstance(duration_value, (int, float)):
        return float(duration_value)

    duration_str = str(duration_value).strip()

    # Try pure number
    try:
        return float(duration_str)
    except ValueError:
        pass

    # Try HH:MM:SS or MM:SS format
    time_match = re.match(r'^(\d+):(\d+):(\d+)$', duration_str)
    if time_match:
        hours, minutes, seconds = map(int, time_match.groups())
        return hours * 60 + minutes + seconds / 60

    time_match = re.match(r'^(\d+):(\d+)$', duration_str)
    if time_match:
        minutes, seconds = map(int, time_match.groups())
        return minutes + seconds / 60

    # Try "Xh Ym" / "X hours Y minutes" / "X hr Y min" patterns
    total = 0
    hour_match = re.search(r'(\d+(?:\.\d+)?)\s*(?:h|hr|hour|hours)', duration_str, re.IGNORECASE)
    min_match = re.search(r'(\d+(?:\.\d+)?)\s*(?:m|min|mins|minute|minutes)', duration_str, re.IGNORECASE)

    if hour_match:
        total += float(hour_match.group(1)) * 60
    if min_match:
        total += float(min_match.group(1))

    if hour_match or min_match:
        return total

    # Fallback: try to extract any number
    num_match = re.search(r'(\d+(?:\.\d+)?)', duration_str)
    if num_match:
        return float(num_match.group(1))

    return 0


def parse_date(date_value):
    """
    Parse various date formats into a datetime.date object.

    Handles: dd-mm-yyyy, dd/mm/yyyy, yyyy-mm-dd, dd-mm-yy, etc.
    """
    if pd.isna(date_value):
        return None

    # If already a datetime
    if isinstance(date_value, datetime):
        return date_value.date()
    if hasattr(date_value, 'date'):
        return date_value.date()

    date_str = str(date_value).strip()

    # Common date formats to try
    formats = [
        "%d-%m-%Y",   # dd-mm-yyyy
        "%d/%m/%Y",   # dd/mm/yyyy
        "%Y-%m-%d",   # yyyy-mm-dd
        "%d-%m-%y",   # dd-mm-yy
        "%d/%m/%y",   # dd/mm/yy
        "%m-%d-%Y",   # mm-dd-yyyy
        "%m/%d/%Y",   # mm/dd/yyyy
        "%d %b %Y",   # dd Mon yyyy (e.g., 15 Jun 2025)
        "%d %B %Y",   # dd Month yyyy (e.g., 15 June 2025)
        "%Y/%m/%d",   # yyyy/mm/dd
    ]

    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt).date()
        except ValueError:
            continue

    # Try pandas as fallback
    try:
        return pd.to_datetime(date_str, dayfirst=True).date()
    except Exception:
        return None


# ============================================================
# CSV Reading & Column Detection
# ============================================================

def read_csv_safe(filepath):
    """
    Read a CSV file with encoding fallback (UTF-8 -> Latin-1).

    Returns: (DataFrame, None) on success, or (None, error_message) on failure.
    """
    try:
        df = pd.read_csv(filepath, encoding='utf-8')
    except UnicodeDecodeError:
        try:
            df = pd.read_csv(filepath, encoding='latin-1')
        except Exception as e:
            return None, f"Error reading file: {str(e)}"
    except Exception as e:
        return None, f"Error reading file: {str(e)}"

    # Strip whitespace from column names
    df.columns = df.columns.str.strip()
    return df, None


def detect_registration_columns(df):
    """
    Auto-detect column mapping for registration CSV using keyword matching.

    Returns a dict of {standardized_key: csv_column_name} for detected columns.
    Does NOT raise errors for missing columns.
    """
    col_map = {}
    for col in df.columns:
        col_lower = col.lower().strip()
        if 'email' in col_lower and 'Email' not in col_map:
            col_map['Email'] = col
        elif ('training date' in col_lower or 'training_date' in col_lower) and 'Training date' not in col_map:
            col_map['Training date'] = col
        elif 'attendance' in col_lower and 'day 1' in col_lower and 'Attendance - Day 1' not in col_map:
            col_map['Attendance - Day 1'] = col
        elif 'attendance' in col_lower and 'day 2' in col_lower and 'Attendance - Day 2' not in col_map:
            col_map['Attendance - Day 2'] = col
        elif (col_lower == 'name' or col_lower == 'participant name') and 'Name' not in col_map:
            col_map['Name'] = col
        elif ('version' in col_lower or 'program' in col_lower) and 'Version' not in col_map:
            col_map['Version'] = col
    return col_map


def detect_zoom_columns(df):
    """
    Auto-detect column mapping for Zoom CSV using keyword matching.

    Returns a dict of {standardized_key: csv_column_name} for detected columns.
    Does NOT raise errors for missing columns.
    """
    col_map = {}
    for col in df.columns:
        col_lower = col.lower().strip()
        if 'email' in col_lower and 'Email' not in col_map:
            col_map['Email'] = col
        elif 'duration' in col_lower and 'Duration' not in col_map:
            col_map['Duration'] = col
        elif ('name' in col_lower) and 'Name' not in col_map:
            col_map['Name'] = col
    return col_map


def load_registration_csv(filepath, col_map_override=None):
    """
    Load registration CSV and validate required columns.

    Args:
        filepath: Path to the CSV file.
        col_map_override: Optional dict. If provided, skip auto-detection and use this mapping.

    Returns: (DataFrame, col_map) on success, or (None, error_message) on failure.
    """
    df, err = read_csv_safe(filepath)
    if df is None:
        return None, err

    if col_map_override is not None:
        col_map = col_map_override
    else:
        col_map = detect_registration_columns(df)

    required = ['Email', 'Training date']
    missing = [r for r in required if r not in col_map]
    if missing:
        return None, f"Missing required columns: {', '.join(missing)}"

    return df, col_map


def load_zoom_csv(filepath, col_map_override=None):
    """
    Load Zoom attendance CSV and validate required columns.

    Args:
        filepath: Path to the CSV file.
        col_map_override: Optional dict. If provided, skip auto-detection and use this mapping.

    Returns: (DataFrame, col_map) on success, or (None, error_message) on failure.
    """
    df, err = read_csv_safe(filepath)
    if df is None:
        return None, err

    if col_map_override is not None:
        col_map = col_map_override
    else:
        col_map = detect_zoom_columns(df)

    required = ['Email', 'Duration']
    missing = [r for r in required if r not in col_map]
    if missing:
        return None, f"Missing required columns in Zoom CSV: {', '.join(missing)}"

    return df, col_map


# ============================================================
# Data Extraction Helpers
# ============================================================

def get_unique_programs(df, col_map):
    """Extract unique program/version values from registration data."""
    if 'Version' not in col_map:
        return []

    version_col = col_map['Version']
    programs = df[version_col].dropna().unique().tolist()
    return sorted([str(p) for p in programs])


def get_unique_dates(df, col_map):
    """Extract unique training dates from registration data."""
    date_col = col_map['Training date']
    dates = []
    for val in df[date_col].dropna().unique():
        parsed = parse_date(val)
        if parsed:
            dates.append(parsed)
    return sorted(set(dates))
