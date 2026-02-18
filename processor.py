"""
Core attendance processing logic.
Handles matching registration data with Zoom attendance data.
"""

import os
import pandas as pd
from datetime import date, datetime
from utils import (
    normalize_email,
    parse_duration_to_minutes,
    parse_date,
    load_registration_csv,
    load_zoom_csv,
)


class AttendanceResult:
    """Holds the results of attendance processing."""

    def __init__(self):
        self.total_registrations = 0
        self.total_zoom_participants = 0
        self.matched_present = 0      # Y - met duration threshold
        self.matched_below_threshold = 0  # N - found in Zoom but < 60 min
        self.not_found_in_zoom = 0    # N - not in Zoom at all
        self.unmatched_zoom_emails = []  # Zoom emails not found in registration data
        self.output_filepath = ""
        self.errors = []
        self.warnings = []

    @property
    def success(self):
        return len(self.errors) == 0

    def summary(self):
        lines = []
        lines.append(f"Total registrations for selected date: {self.total_registrations}")
        lines.append(f"Total Zoom participants: {self.total_zoom_participants}")
        lines.append(f"Marked Present (Y): {self.matched_present}")
        lines.append(f"Below duration threshold (N): {self.matched_below_threshold}")
        lines.append(f"Not found in Zoom (N): {self.not_found_in_zoom}")
        lines.append(f"In Zoom but not registered: {len(self.unmatched_zoom_emails)}")
        if self.output_filepath:
            lines.append(f"\nOutput saved to: {self.output_filepath}")
        if self.warnings:
            lines.append("\nWarnings:")
            for w in self.warnings:
                lines.append(f"  - {w}")
        return "\n".join(lines)


def process_attendance(
    reg_filepath,
    zoom_filepath,
    selected_date,
    attendance_day,
    program_filter=None,
    min_duration=60,
    output_dir=None,
    reg_col_map_override=None,
    zoom_col_map_override=None,
):
    """
    Main attendance processing function.

    Args:
        reg_filepath: Path to registration CSV
        zoom_filepath: Path to Zoom attendance CSV
        selected_date: date object - the training date to filter on
        attendance_day: 1 or 2 - which attendance day column to update
        program_filter: Optional program/version to filter (None or "All" = no filter)
        min_duration: Minimum duration in minutes to mark as present (default 60)
        output_dir: Directory to save output file (defaults to same as registration CSV)
        reg_col_map_override: Optional column mapping dict for registration CSV
        zoom_col_map_override: Optional column mapping dict for Zoom CSV

    Returns:
        AttendanceResult object
    """
    result = AttendanceResult()

    # --- Step 1: Load Data ---
    reg_data = load_registration_csv(reg_filepath, col_map_override=reg_col_map_override)
    if reg_data[0] is None:
        result.errors.append(f"Registration CSV: {reg_data[1]}")
        return result
    reg_df, reg_col_map = reg_data

    zoom_data = load_zoom_csv(zoom_filepath, col_map_override=zoom_col_map_override)
    if zoom_data[0] is None:
        result.errors.append(f"Zoom CSV: {zoom_data[1]}")
        return result
    zoom_df, zoom_col_map = zoom_data

    # --- Step 2: Determine attendance column ---
    if attendance_day == 1:
        att_col_key = 'Attendance - Day 1'
    else:
        att_col_key = 'Attendance - Day 2'

    if att_col_key not in reg_col_map:
        # Create the column if it doesn't exist
        att_col_name = f"Attendance - Day {attendance_day}"
        reg_df[att_col_name] = ""
        reg_col_map[att_col_key] = att_col_name

    att_col_name = reg_col_map[att_col_key]

    # Ensure attendance column is string type (empty cols are read as float64)
    reg_df[att_col_name] = reg_df[att_col_name].fillna("").astype(str)

    # --- Step 3: Normalize emails ---
    email_col_reg = reg_col_map['Email']
    email_col_zoom = zoom_col_map['Email']
    duration_col_zoom = zoom_col_map['Duration']

    reg_df['_normalized_email'] = reg_df[email_col_reg].apply(normalize_email)
    zoom_df['_normalized_email'] = zoom_df[email_col_zoom].apply(normalize_email)

    # --- Step 4: Parse training dates and filter ---
    date_col = reg_col_map['Training date']
    reg_df['_parsed_date'] = reg_df[date_col].apply(parse_date)

    # Filter by selected date
    date_mask = reg_df['_parsed_date'] == selected_date

    # Filter by program if specified
    if program_filter and program_filter != "All" and 'Version' in reg_col_map:
        version_col = reg_col_map['Version']
        program_mask = reg_df[version_col].astype(str).str.strip() == program_filter
        date_mask = date_mask & program_mask

    filtered_indices = reg_df[date_mask].index
    result.total_registrations = len(filtered_indices)

    if result.total_registrations == 0:
        result.warnings.append(
            f"No registrations found for date {selected_date.strftime('%d-%m-%Y')}"
            + (f" and program '{program_filter}'" if program_filter and program_filter != "All" else "")
        )

    # --- Step 5: Process Zoom data - handle duplicates by summing durations ---
    zoom_df['_duration_minutes'] = zoom_df[duration_col_zoom].apply(parse_duration_to_minutes)

    zoom_grouped = zoom_df.groupby('_normalized_email')['_duration_minutes'].sum().reset_index()
    zoom_dict = dict(zip(zoom_grouped['_normalized_email'], zoom_grouped['_duration_minutes']))

    # Remove empty email entries
    zoom_dict.pop("", None)

    result.total_zoom_participants = len(zoom_dict)

    # --- Step 6: Match and mark attendance ---
    for idx in filtered_indices:
        email = reg_df.at[idx, '_normalized_email']

        if not email:
            result.not_found_in_zoom += 1
            reg_df.at[idx, att_col_name] = "N"
            continue

        if email in zoom_dict:
            total_duration = zoom_dict[email]
            if total_duration >= min_duration:
                reg_df.at[idx, att_col_name] = "Y"
                result.matched_present += 1
            else:
                reg_df.at[idx, att_col_name] = "N"
                result.matched_below_threshold += 1
        else:
            reg_df.at[idx, att_col_name] = "N"
            result.not_found_in_zoom += 1

    # --- Step 6b: Find Zoom participants not in registration data ---
    reg_emails_for_date = set(
        reg_df.at[idx, '_normalized_email']
        for idx in filtered_indices
        if reg_df.at[idx, '_normalized_email']
    )
    result.unmatched_zoom_emails = sorted(
        email for email in zoom_dict if email not in reg_emails_for_date
    )

    # --- Step 7: Save output ---
    # Remove temporary columns
    output_df = reg_df.drop(columns=['_normalized_email', '_parsed_date'], errors='ignore')

    if output_dir is None:
        output_dir = os.path.dirname(reg_filepath)

    date_str = selected_date.strftime('%Y-%m-%d')
    day_str = f"Day{attendance_day}"
    program_str = f"_{program_filter}" if program_filter and program_filter != "All" else ""
    program_str = program_str.replace(" ", "_").replace("/", "-")

    filename = f"attendance_output_{date_str}_{day_str}{program_str}.csv"
    output_path = os.path.join(output_dir, filename)

    try:
        output_df.to_csv(output_path, index=False, encoding='utf-8-sig')
        result.output_filepath = output_path
    except Exception as e:
        result.errors.append(f"Error saving output: {str(e)}")

    return result
