# Attendance Automation Tool

A desktop GUI application (Python/Tkinter) that automates attendance marking by cross-referencing **registration data** (exported from Google Sheets) with **Zoom attendance reports**. Trainers can quickly mark attendance as Y/N based on participant duration, with support for multiple programs, training days, and flexible CSV formats.

---

## Features

- **Email-based matching** with case-insensitive normalization
- **Duration threshold**: Mark "Y" if Zoom duration >= 60 minutes, "N" otherwise
- **Duplicate Zoom entry handling**: Automatically sums durations for participants who disconnected and rejoined
- **Smart column auto-detection** with manual override via a mapping dialog
- **Multi-format support**: Parses dates (10+ formats) and durations (numeric, HH:MM:SS, text)
- **Program/version filtering**: Process attendance for a specific program or all at once
- **Day 1 / Day 2 toggle**: Mark attendance for either training day
- **Unmatched Zoom detection**: Identifies Zoom participants not found in registration data
- **Modern Slate Teal UI** with step indicators, stat dashboard, tooltips, and hover effects
- **Standalone .exe** via PyInstaller — no Python installation required to run

---

## Screenshots

The application follows a 3-step workflow:

1. **Load Files** — Browse and map CSV columns
2. **Configure** — Set date, day, and program filter
3. **Process** — Run matching and view results in the summary dashboard

---

## Getting Started

### Prerequisites

- Python 3.10+ (tested with 3.13)
- Windows (DPI awareness uses Windows APIs; core logic is cross-platform)

### Installation

```bash
# Clone the repository
git clone https://github.com/anilkumarjonnalagadda/attendance-automation.git
cd attendance-automation

# Install dependencies
pip install -r requirements.txt
```

### Run the Application

```bash
python main.py
```

### Build Standalone Executable

```bash
pyinstaller --onefile --windowed --name AttendanceAutomation --add-data "theme.py;." main.py
```

The `.exe` will be created in the `dist/` folder (~41 MB).

---

## Usage

### Step 1: Prepare Your CSV Files

**Registration CSV** — Export from your Google Sheet. Required columns:

| Column | Required | Description |
|--------|----------|-------------|
| Email | Yes | Participant email address |
| Training date | Yes | Date in any supported format |
| Attendance - Day 1 | No | Populated by the tool (Y/N) |
| Attendance - Day 2 | No | Populated by the tool (Y/N) |
| Name | No | Participant name |
| Version | No | Program/version for filtering |

Any additional columns (Organization, Location, Remarks, etc.) are preserved in the output.

**Zoom Attendance CSV** — Download from Zoom. Required columns:

| Column | Required | Description |
|--------|----------|-------------|
| Email | Yes | Participant email address |
| Duration | Yes | Attendance duration (any format) |
| Name | No | Participant name |

### Step 2: Load Files in the App

1. Click **Browse** for Registration CSV
2. Review the auto-detected column mapping in the dialog, adjust if needed, and click **Confirm**
3. Repeat for the Zoom CSV

### Step 3: Configure Settings

- **Training Date**: Select the date to filter registrations
- **Attendance Day**: Toggle between Day 1 and Day 2
- **Program Filter**: Choose a specific program or "All"

### Step 4: Process

Click **"▶ Mark Attendance"**. The tool will:

1. Filter registrations by the selected date and program
2. Normalize all email addresses (lowercase, trim whitespace)
3. Sum Zoom durations for participants with multiple sessions
4. Mark "Y" if total duration >= 60 minutes, "N" otherwise
5. Save the output CSV with attendance marked

### Step 5: Review Results

The **summary dashboard** shows 5 stat cards:

| Card | Color | Description |
|------|-------|-------------|
| Registrations | Teal | Total registrations for the selected date |
| Present (Y) | Green | Met the 60-minute threshold |
| Below Threshold | Amber | Found in Zoom but < 60 minutes |
| Not in Zoom | Red | Not found in Zoom data at all |
| Not Registered | Amber | In Zoom but not in registration data |

Unmatched Zoom participant emails are listed individually in the status log.

---

## Supported Formats

### Duration Formats

| Format | Example | Parsed As |
|--------|---------|-----------|
| Integer/float | `75` | 75 minutes |
| String number | `"90"` | 90 minutes |
| HH:MM:SS | `"1:30:00"` | 90 minutes |
| MM:SS | `"45:30"` | 45.5 minutes |
| Text (h/m) | `"1h 30m"` | 90 minutes |
| Text (words) | `"2 hours"` | 120 minutes |

### Date Formats

The tool auto-detects dates in these formats:

- `dd-mm-yyyy` (15-01-2025)
- `dd/mm/yyyy` (15/01/2025)
- `yyyy-mm-dd` (2025-01-15)
- `dd-mm-yy` (15-01-25)
- `dd/mm/yy` (15/01/25)
- `dd Mon yyyy` (15 Jan 2025)
- `dd Month yyyy` (15 January 2025)
- And more via pandas fallback

---

## Output

The output CSV is saved in the same directory as the registration CSV with the naming convention:

```
attendance_output_{YYYY-MM-DD}_{DayN}{_Program}.csv
```

**Examples:**
- `attendance_output_2025-01-15_Day1.csv`
- `attendance_output_2025-01-15_Day2_Python_Basics_v1.csv`

The output preserves all original columns and populates the selected Attendance Day column with Y or N.

---

## Column Mapping

If your CSV column headers don't match the expected names, the **Column Mapping Dialog** lets you map them manually:

- Auto-detection pre-fills mappings using keyword matching
- Required fields are highlighted with a teal accent bar
- Per-row status indicators: ✓ (mapped), ○ (unmapped), ⚠ (duplicate)
- Reset button restores auto-detected values

This handles column renames, additions, and structural changes without modifying the tool.

---

## Project Structure

```
attendance_automation/
├── main.py                    # Entry point, DPI config, ttk styling
├── gui.py                     # Main Tkinter GUI (AttendanceApp class)
├── theme.py                   # Shared Slate Teal theme (colors, fonts, helpers)
├── processor.py               # Core matching algorithm (AttendanceResult)
├── utils.py                   # CSV loading, parsing, normalization utilities
├── column_mapping_dialog.py   # Modal column mapping dialog
├── requirements.txt           # Python dependencies
├── test_processor.py          # Test suite
└── test_data/
    ├── sample_registration.csv
    └── sample_zoom.csv
```

---

## Running Tests

```bash
python test_processor.py
```

Tests cover:
- Duration parsing (8 formats)
- Date parsing (4 formats)
- Email normalization (4 edge cases)
- Full processing pipeline (matching, thresholds, unmatched detection)

---

## Dependencies

| Package | Purpose |
|---------|---------|
| pandas | CSV data processing and DataFrame operations |
| openpyxl | Excel format support |
| tkcalendar | Calendar date picker widget |
| pyinstaller | Standalone .exe packaging (dev only) |

Tkinter is included with Python and requires no separate installation.

---

## License

This project is for internal use.
