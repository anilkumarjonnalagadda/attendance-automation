"""
Tkinter GUI for Attendance Automation Tool.
Comprehensive UI with Slate Teal theme, step indicators,
status badges, summary dashboard, and tooltips.
"""

import os
import sys
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import date
from tkcalendar import DateEntry

from theme import (
    BG_COLOR, HEADER_BG, HEADER_FG, ACCENT_COLOR, ACCENT_HOVER,
    ACCENT_LIGHT, SUCCESS_COLOR, ERROR_COLOR, WARNING_COLOR,
    CARD_BG, TEXT_COLOR, TEXT_SECONDARY, MUTED_COLOR,
    BORDER_COLOR, SHADOW_COLOR, STEP_INACTIVE, STEP_ACTIVE,
    BADGE_PENDING, BADGE_LOADED,
    FONT_HEADING, FONT_SUBHEAD, FONT_BODY, FONT_BODY_BOLD,
    FONT_SMALL, FONT_MONO, FONT_BUTTON_LG, FONT_BUTTON,
    ToolTip, make_hover_button, create_shadow_card,
)
from utils import (
    read_csv_safe,
    detect_registration_columns,
    detect_zoom_columns,
    get_unique_programs,
    get_unique_dates,
    REGISTRATION_FIELDS,
    ZOOM_FIELDS,
)
from column_mapping_dialog import ColumnMappingDialog
from processor import process_attendance


class AttendanceApp:
    """Main GUI application for attendance automation."""

    def __init__(self, root):
        self.root = root
        self.root.title("Attendance Automation Tool")
        self.root.geometry("720x820")
        self.root.minsize(700, 780)
        self.root.configure(bg=BG_COLOR)

        # State variables
        self.reg_filepath = tk.StringVar()
        self.zoom_filepath = tk.StringVar()
        self.attendance_day = tk.IntVar(value=1)
        self.program_var = tk.StringVar(value="All")
        self.reg_col_map = None
        self.reg_df = None
        self.zoom_col_map = None
        self.output_dir = None

        # Step states
        self.steps = [
            {"label": "Load Files", "state": "active"},
            {"label": "Configure", "state": "inactive"},
            {"label": "Process", "state": "inactive"},
        ]

        # UI references (set during build)
        self.reg_filename_label = None
        self.reg_badge = None
        self.zoom_filename_label = None
        self.zoom_badge = None
        self.segment_buttons = []
        self.summary_frame = None
        self.stat_cards = {}
        self.log_card_container = None

        self._build_ui()

    # ================================================================
    # UI Building
    # ================================================================

    def _build_ui(self):
        """Build the complete UI from sub-methods."""
        self._build_header()

        # Scrollable main content
        self.main_frame = tk.Frame(self.root, bg=BG_COLOR)
        self.main_frame.pack(fill='both', expand=True, padx=24, pady=(10, 15))

        self._build_step_indicator()
        self._build_file_selection_card()
        self._build_settings_card()
        self._build_process_section()
        self._build_summary_dashboard()
        self._build_log_card()
        self._build_action_bar()

    def _build_header(self):
        """Build the enhanced app header."""
        header_frame = tk.Frame(self.root, bg=HEADER_BG, height=70)
        header_frame.pack(fill='x')
        header_frame.pack_propagate(False)

        header_inner = tk.Frame(header_frame, bg=HEADER_BG)
        header_inner.place(relx=0.5, rely=0.5, anchor='center')

        tk.Label(
            header_inner,
            text="Attendance Automation Tool",
            font=FONT_HEADING,
            bg=HEADER_BG, fg=HEADER_FG,
        ).pack()

        tk.Label(
            header_inner,
            text="Cross-reference registration and Zoom data",
            font=FONT_SMALL,
            bg=HEADER_BG, fg=MUTED_COLOR,
        ).pack()

    def _build_step_indicator(self):
        """Draw a 3-step progress indicator using Canvas."""
        self.step_canvas = tk.Canvas(
            self.main_frame, bg=BG_COLOR, height=60, highlightthickness=0
        )
        self.step_canvas.pack(fill='x', pady=(0, 8))
        self.step_canvas.bind("<Configure>", lambda e: self._draw_steps())
        # Initial draw after a short delay to get correct width
        self.root.after(50, self._draw_steps)

    def _draw_steps(self):
        """Redraw step circles and connecting lines."""
        c = self.step_canvas
        c.delete("all")
        w = c.winfo_width()
        if w < 10:
            w = 670
        y_center = 22
        radius = 14
        step_count = len(self.steps)
        spacing = w / (step_count + 1)

        for i, step in enumerate(self.steps):
            x = spacing * (i + 1)

            # Connecting line to next step
            if i < step_count - 1:
                next_x = spacing * (i + 2)
                line_color = STEP_ACTIVE if step["state"] == "completed" else STEP_INACTIVE
                c.create_line(x + radius + 2, y_center, next_x - radius - 2, y_center,
                              fill=line_color, width=2)

            # Circle
            if step["state"] == "completed":
                c.create_oval(x - radius, y_center - radius, x + radius, y_center + radius,
                              fill=STEP_ACTIVE, outline=STEP_ACTIVE)
                c.create_text(x, y_center, text="\u2713", fill="white",
                              font=("Segoe UI", 11, "bold"))
            elif step["state"] == "active":
                c.create_oval(x - radius, y_center - radius, x + radius, y_center + radius,
                              fill=CARD_BG, outline=STEP_ACTIVE, width=2)
                c.create_text(x, y_center, text=str(i + 1), fill=STEP_ACTIVE,
                              font=("Segoe UI", 10, "bold"))
            else:
                c.create_oval(x - radius, y_center - radius, x + radius, y_center + radius,
                              fill=CARD_BG, outline=STEP_INACTIVE, width=2)
                c.create_text(x, y_center, text=str(i + 1), fill=STEP_INACTIVE,
                              font=("Segoe UI", 10))

            # Label below circle
            label_color = TEXT_COLOR if step["state"] in ("active", "completed") else MUTED_COLOR
            c.create_text(x, y_center + radius + 14, text=step["label"],
                          fill=label_color, font=FONT_SMALL)

    def _update_step(self, step_index, state):
        """Update a step's state and redraw."""
        self.steps[step_index]["state"] = state
        self._draw_steps()

    def _check_files_loaded(self):
        """Check if both files are loaded and advance step indicator."""
        if self.reg_col_map is not None and self.zoom_col_map is not None:
            self._update_step(0, "completed")
            self._update_step(1, "active")

    def _build_file_selection_card(self):
        """Build the file selection card with status badges."""
        _, card = create_shadow_card(self.main_frame, "File Selection")

        # Registration CSV row
        self.reg_filename_label, self.reg_badge = self._build_file_row(
            card, "Registration CSV:",
            self._browse_registration,
            "Select the CSV exported from the Google Sheet registration data"
        )

        # Spacer
        tk.Frame(card, bg=BORDER_COLOR, height=1).pack(fill='x', pady=8)

        # Zoom CSV row
        self.zoom_filename_label, self.zoom_badge = self._build_file_row(
            card, "Zoom Attendance CSV:",
            self._browse_zoom,
            "Select the CSV exported from Zoom attendance report"
        )

    def _build_file_row(self, parent, label_text, browse_cmd, tooltip_text):
        """Build a single file selection row with badge."""
        row = tk.Frame(parent, bg=CARD_BG)
        row.pack(fill='x', pady=2)

        # Label
        tk.Label(
            row, text=label_text,
            font=FONT_BODY_BOLD, bg=CARD_BG, fg=TEXT_COLOR,
            width=18, anchor='w',
        ).pack(side='left')

        # Filename display
        filename_label = tk.Label(
            row, text="No file selected",
            font=FONT_SMALL, bg="#f8fafc", fg=MUTED_COLOR,
            anchor='w', relief='solid', bd=1, padx=8, pady=4,
        )
        filename_label.pack(side='left', fill='x', expand=True, padx=(0, 8))

        # Status badge
        badge = tk.Label(
            row, text=" \u25CB Not loaded ",
            font=FONT_SMALL, bg=BADGE_PENDING, fg=MUTED_COLOR,
            padx=6, pady=2,
        )
        badge.pack(side='left', padx=(0, 8))

        # Browse button with hover
        browse_btn = make_hover_button(
            row, text="Browse",
            bg=ACCENT_COLOR, fg="white", hover_bg=ACCENT_HOVER,
            font=FONT_BODY, command=browse_cmd, padx=14, pady=3,
        )
        browse_btn.pack(side='right')
        ToolTip(browse_btn, tooltip_text)

        return filename_label, badge

    def _build_settings_card(self):
        """Build the settings card with date, day toggle, and program filter."""
        _, card = create_shadow_card(self.main_frame, "Settings")

        # Training Date
        date_row = tk.Frame(card, bg=CARD_BG)
        date_row.pack(fill='x', pady=(0, 10))

        tk.Label(
            date_row, text="Training Date:",
            font=FONT_BODY, bg=CARD_BG, fg=TEXT_COLOR,
            width=18, anchor='w',
        ).pack(side='left')

        self.date_picker = DateEntry(
            date_row, width=18,
            background=ACCENT_COLOR, foreground='white',
            borderwidth=1, date_pattern='dd-mm-yyyy',
            font=FONT_BODY,
        )
        self.date_picker.pack(side='left')
        ToolTip(self.date_picker, "Select the training date to filter registrations")

        # Attendance Day — Segmented Control
        day_row = tk.Frame(card, bg=CARD_BG)
        day_row.pack(fill='x', pady=(0, 10))

        tk.Label(
            day_row, text="Attendance Day:",
            font=FONT_BODY, bg=CARD_BG, fg=TEXT_COLOR,
            width=18, anchor='w',
        ).pack(side='left')

        self._build_segmented_control(day_row)

        # Program Filter
        prog_row = tk.Frame(card, bg=CARD_BG)
        prog_row.pack(fill='x')

        tk.Label(
            prog_row, text="Program Filter:",
            font=FONT_BODY, bg=CARD_BG, fg=TEXT_COLOR,
            width=18, anchor='w',
        ).pack(side='left')

        self.program_dropdown = ttk.Combobox(
            prog_row, textvariable=self.program_var,
            state='readonly', font=FONT_BODY, width=30,
        )
        self.program_dropdown['values'] = ["All"]
        self.program_dropdown.pack(side='left')
        ToolTip(self.program_dropdown, "Filter by program/version, or 'All' for all programs")

    def _build_segmented_control(self, parent):
        """Build a Day 1 / Day 2 segmented toggle."""
        seg_frame = tk.Frame(parent, bg=BORDER_COLOR, bd=1, relief='solid')
        seg_frame.pack(side='left')

        options = [("  Day 1  ", 1), ("  Day 2  ", 2)]
        self.segment_buttons = []

        for text, value in options:
            is_active = self.attendance_day.get() == value
            btn = tk.Label(
                seg_frame, text=text, font=FONT_BODY,
                padx=14, pady=4, cursor='hand2',
                bg=ACCENT_COLOR if is_active else CARD_BG,
                fg="white" if is_active else TEXT_COLOR,
            )
            btn.pack(side='left')
            btn.bind("<Button-1>", lambda e, v=value: self._on_segment_click(v))
            self.segment_buttons.append((btn, value))

        ToolTip(seg_frame, "Choose which attendance day column to update")

    def _on_segment_click(self, value):
        """Handle segment toggle click."""
        self.attendance_day.set(value)
        for btn, v in self.segment_buttons:
            if v == value:
                btn.config(bg=ACCENT_COLOR, fg="white")
            else:
                btn.config(bg=CARD_BG, fg=TEXT_COLOR)

    def _build_process_section(self):
        """Build the process button area with progress bar."""
        proc_frame = tk.Frame(self.main_frame, bg=BG_COLOR)
        proc_frame.pack(fill='x', pady=(5, 0))

        # Separator
        tk.Frame(proc_frame, bg=BORDER_COLOR, height=1).pack(fill='x', pady=(0, 12))

        self.proc_btn_frame = tk.Frame(proc_frame, bg=BG_COLOR)
        self.proc_btn_frame.pack()

        self.process_btn = make_hover_button(
            self.proc_btn_frame,
            text="\u25B6  Mark Attendance",
            bg=ACCENT_COLOR, fg="white", hover_bg=ACCENT_HOVER,
            font=FONT_BUTTON_LG,
            command=self._process_attendance,
            padx=40, pady=10,
        )
        self.process_btn.pack()
        ToolTip(self.process_btn, "Match Zoom attendance with registrations and mark present/absent")

        # Progress bar (hidden initially)
        self.progress_frame = tk.Frame(proc_frame, bg=BG_COLOR)
        self.progress_bar = ttk.Progressbar(
            self.progress_frame, mode='indeterminate', length=350,
            style='Teal.Horizontal.TProgressbar',
        )
        self.progress_bar.pack(pady=5)

        # Processing status text
        self.processing_label = tk.Label(
            self.progress_frame, text="Processing attendance data...",
            font=FONT_SMALL, bg=BG_COLOR, fg=TEXT_SECONDARY,
        )
        self.processing_label.pack()

    def _build_summary_dashboard(self):
        """Create the summary dashboard (hidden until processing completes)."""
        self.summary_frame = tk.Frame(self.main_frame, bg=BG_COLOR)
        # Not packed until results are available

        stats = [
            ("total_reg", "Registrations", ACCENT_COLOR),
            ("present", "Present (Y)", SUCCESS_COLOR),
            ("below", "Below Threshold", WARNING_COLOR),
            ("absent", "Not in Zoom", ERROR_COLOR),
            ("unmatched_zoom", "Not Registered", WARNING_COLOR),
        ]

        self.stat_cards = {}
        for key, label, color in stats:
            # Outer shadow
            shadow = tk.Frame(self.summary_frame, bg=SHADOW_COLOR)
            shadow.pack(side='left', fill='both', expand=True, padx=4, pady=(0, 2))

            card_frame = tk.Frame(shadow, bg=CARD_BG, padx=10, pady=8)
            card_frame.pack(fill='both')

            value_label = tk.Label(
                card_frame, text="--",
                font=("Segoe UI", 22, "bold"),
                bg=CARD_BG, fg=color,
            )
            value_label.pack()

            tk.Label(
                card_frame, text=label,
                font=FONT_SMALL, bg=CARD_BG, fg=TEXT_SECONDARY,
            ).pack()

            self.stat_cards[key] = value_label

    def _show_summary(self, result):
        """Populate and show the summary dashboard."""
        self.stat_cards["total_reg"].config(text=str(result.total_registrations))
        self.stat_cards["present"].config(text=str(result.matched_present))
        self.stat_cards["below"].config(text=str(result.matched_below_threshold))
        self.stat_cards["absent"].config(text=str(result.not_found_in_zoom))
        self.stat_cards["unmatched_zoom"].config(text=str(len(result.unmatched_zoom_emails)))
        self.summary_frame.pack(fill='x', pady=(8, 4), before=self.log_card_container)

    def _hide_summary(self):
        """Hide the summary dashboard."""
        self.summary_frame.pack_forget()

    def _build_log_card(self):
        """Build the status log card."""
        self.log_card_container = tk.Frame(self.main_frame, bg=BG_COLOR)
        self.log_card_container.pack(fill='both', expand=True, pady=(0, 5))

        # Title row with Clear button
        title_row = tk.Frame(self.log_card_container, bg=BG_COLOR)
        title_row.pack(fill='x', pady=(0, 6))

        tk.Label(
            title_row, text="Status Log",
            font=FONT_SUBHEAD, bg=BG_COLOR, fg=TEXT_COLOR,
        ).pack(side='left')

        clear_btn = make_hover_button(
            title_row, text="Clear",
            bg=BG_COLOR, fg=MUTED_COLOR, hover_bg=BORDER_COLOR,
            font=FONT_SMALL, command=self._clear_log, padx=8, pady=2,
        )
        clear_btn.pack(side='right')

        # Shadow card for log
        shadow = tk.Frame(self.log_card_container, bg=SHADOW_COLOR)
        shadow.pack(fill='both', expand=True, padx=(0, 2), pady=(0, 2))

        log_inner = tk.Frame(shadow, bg=CARD_BG, padx=2, pady=2)
        log_inner.pack(fill='both', expand=True)

        self.log_text = tk.Text(
            log_inner, height=8, font=FONT_MONO,
            bg="#fafbfc", fg=TEXT_COLOR,
            relief='flat', bd=0, wrap='word',
            state='disabled', padx=10, pady=8,
        )
        self.log_text.pack(fill='both', expand=True)

        # Colored log tags
        self.log_text.tag_config('info', foreground=TEXT_COLOR)
        self.log_text.tag_config('success', foreground=SUCCESS_COLOR)
        self.log_text.tag_config('error', foreground=ERROR_COLOR)
        self.log_text.tag_config('warning', foreground=WARNING_COLOR)

    def _build_action_bar(self):
        """Build the bottom action bar with Reset and Open Folder buttons."""
        action_bar = tk.Frame(self.main_frame, bg=BG_COLOR)
        action_bar.pack(fill='x', pady=(3, 0))

        # Reset / New Session
        self.reset_btn = make_hover_button(
            action_bar, text="\u21BB  New Session",
            bg=BORDER_COLOR, fg=TEXT_COLOR, hover_bg=SHADOW_COLOR,
            font=FONT_BUTTON, command=self._reset_session, padx=15, pady=6,
        )
        self.reset_btn.pack(side='left')
        ToolTip(self.reset_btn, "Clear all loaded files and settings to start over")

        # Open Output Folder
        self.open_folder_btn = make_hover_button(
            action_bar, text="Open Output Folder",
            bg=BORDER_COLOR, fg=TEXT_COLOR, hover_bg=SHADOW_COLOR,
            font=FONT_BUTTON, command=self._open_output_folder, padx=15, pady=6,
        )
        self.open_folder_btn.pack(side='right')
        self.open_folder_btn.config(state='disabled')
        ToolTip(self.open_folder_btn, "Open the folder containing the generated output CSV")

    # ================================================================
    # Logging
    # ================================================================

    def _log(self, message, tag='info'):
        """Append a message to the log."""
        self.log_text.config(state='normal')
        self.log_text.insert('end', message + "\n", tag)
        self.log_text.see('end')
        self.log_text.config(state='disabled')
        self.root.update_idletasks()

    def _clear_log(self):
        """Clear the log."""
        self.log_text.config(state='normal')
        self.log_text.delete('1.0', 'end')
        self.log_text.config(state='disabled')

    # ================================================================
    # File Browsing with Column Mapping
    # ================================================================

    def _browse_registration(self):
        """Open file dialog for registration CSV, then show column mapping dialog."""
        filepath = filedialog.askopenfilename(
            title="Select Registration CSV",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if not filepath:
            return

        self._clear_log()
        self._log("Loading registration data...", 'info')

        df, err = read_csv_safe(filepath)
        if df is None:
            self._log(f"Error: {err}", 'error')
            return

        auto_map = detect_registration_columns(df)

        dialog = ColumnMappingDialog(
            parent=self.root,
            title="Column Mapping - Registration CSV",
            filename=os.path.basename(filepath),
            csv_columns=list(df.columns),
            field_specs=REGISTRATION_FIELDS,
            auto_detected=auto_map,
        )

        if dialog.result is None:
            self._log("Column mapping cancelled.", 'warning')
            return

        # Store confirmed data
        self.reg_filepath.set(filepath)
        self.reg_df = df
        self.reg_col_map = dialog.result

        # Update UI
        self.reg_filename_label.config(text=os.path.basename(filepath), fg=TEXT_COLOR)
        ToolTip(self.reg_filename_label, filepath)
        self.reg_badge.config(
            text=f" \u2713 Loaded ({len(df)} rows) ",
            bg=BADGE_LOADED, fg=SUCCESS_COLOR,
        )

        row_count = len(self.reg_df)
        self._log(f"Registration CSV loaded: {row_count} rows", 'success')

        mapped_fields = [f"{k} \u2192 {v}" for k, v in self.reg_col_map.items()]
        self._log(f"Mapping: {', '.join(mapped_fields)}", 'info')

        # Populate program dropdown
        programs = get_unique_programs(self.reg_df, self.reg_col_map)
        self.program_dropdown['values'] = ["All"] + programs
        self.program_var.set("All")

        if programs:
            self._log(f"Programs: {', '.join(programs)}", 'info')

        dates = get_unique_dates(self.reg_df, self.reg_col_map)
        if dates:
            date_strs = [d.strftime('%d-%m-%Y') for d in dates[:10]]
            self._log(f"Dates: {', '.join(date_strs)}" +
                      (" ..." if len(dates) > 10 else ""), 'info')

        self._check_files_loaded()

    def _browse_zoom(self):
        """Open file dialog for Zoom attendance CSV, then show column mapping dialog."""
        filepath = filedialog.askopenfilename(
            title="Select Zoom Attendance CSV",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if not filepath:
            return

        df, err = read_csv_safe(filepath)
        if df is None:
            self._log(f"Error: {err}", 'error')
            return

        auto_map = detect_zoom_columns(df)

        dialog = ColumnMappingDialog(
            parent=self.root,
            title="Column Mapping - Zoom CSV",
            filename=os.path.basename(filepath),
            csv_columns=list(df.columns),
            field_specs=ZOOM_FIELDS,
            auto_detected=auto_map,
        )

        if dialog.result is None:
            self._log("Column mapping cancelled.", 'warning')
            return

        self.zoom_filepath.set(filepath)
        self.zoom_col_map = dialog.result

        # Update UI
        self.zoom_filename_label.config(text=os.path.basename(filepath), fg=TEXT_COLOR)
        ToolTip(self.zoom_filename_label, filepath)
        self.zoom_badge.config(
            text=f" \u2713 Loaded ({len(df)} rows) ",
            bg=BADGE_LOADED, fg=SUCCESS_COLOR,
        )

        self._log(f"Zoom CSV loaded: {os.path.basename(filepath)} ({len(df)} rows)", 'success')
        mapped_fields = [f"{k} \u2192 {v}" for k, v in self.zoom_col_map.items()]
        self._log(f"Mapping: {', '.join(mapped_fields)}", 'info')

        self._check_files_loaded()

    # ================================================================
    # Processing
    # ================================================================

    def _process_attendance(self):
        """Run the attendance processing in a background thread."""
        # Validate inputs
        if not self.reg_filepath.get():
            messagebox.showwarning("Input Required", "Please select the Registration CSV file.")
            return
        if not self.zoom_filepath.get():
            messagebox.showwarning("Input Required", "Please select the Zoom Attendance CSV file.")
            return
        if self.reg_col_map is None:
            messagebox.showwarning("Mapping Required", "Please load and map the Registration CSV columns first.")
            return
        if self.zoom_col_map is None:
            messagebox.showwarning("Mapping Required", "Please load and map the Zoom CSV columns first.")
            return

        # Update steps
        self._update_step(1, "completed")
        self._update_step(2, "active")

        # Show processing state
        self._clear_log()
        self._log("Processing attendance...", 'info')
        self._log(f"Date: {self.date_picker.get_date().strftime('%d-%m-%Y')}", 'info')
        self._log(f"Day: {self.attendance_day.get()}", 'info')
        self._log(f"Program: {self.program_var.get()}", 'info')
        self._log("-" * 40, 'info')

        # Swap button for progress bar
        self.process_btn.pack_forget()
        self.progress_frame.pack()
        self.progress_bar.start(15)
        self.root.update_idletasks()

        # Hide previous summary
        self._hide_summary()

        # Capture params for thread
        params = {
            'reg_filepath': self.reg_filepath.get(),
            'zoom_filepath': self.zoom_filepath.get(),
            'selected_date': self.date_picker.get_date(),
            'attendance_day': self.attendance_day.get(),
            'program_filter': self.program_var.get() if self.program_var.get() != "All" else None,
            'min_duration': 60,
            'reg_col_map_override': self.reg_col_map,
            'zoom_col_map_override': self.zoom_col_map,
        }

        # Run in background thread
        thread = threading.Thread(target=self._run_processing, args=(params,), daemon=True)
        thread.start()

    def _run_processing(self, params):
        """Execute processing in background thread."""
        try:
            result = process_attendance(**params)
            self.root.after(0, lambda: self._on_processing_complete(result))
        except Exception as e:
            self.root.after(0, lambda: self._on_processing_error(str(e)))

    def _on_processing_complete(self, result):
        """Handle processing completion on the main thread."""
        # Restore button
        self.progress_bar.stop()
        self.progress_frame.pack_forget()
        self.process_btn.pack()

        if result.errors:
            for err in result.errors:
                self._log(f"ERROR: {err}", 'error')
            self._update_step(2, "active")
        else:
            self._log(result.summary(), 'success')

            if result.unmatched_zoom_emails:
                self._log(
                    f"\nZoom participants not in registration ({len(result.unmatched_zoom_emails)}):",
                    'warning'
                )
                for email in result.unmatched_zoom_emails:
                    self._log(f"  • {email}", 'warning')

            self._show_summary(result)
            self._update_step(2, "completed")

            if result.output_filepath:
                self.output_dir = os.path.dirname(result.output_filepath)
                self.open_folder_btn.config(state='normal')

            if result.total_registrations == 0:
                self._log(
                    "\nTip: Check that the selected date matches the format in your CSV.",
                    'warning'
                )

    def _on_processing_error(self, error_msg):
        """Handle processing error on the main thread."""
        self.progress_bar.stop()
        self.progress_frame.pack_forget()
        self.process_btn.pack()
        self._log(f"Unexpected error: {error_msg}", 'error')
        self._update_step(2, "active")

    # ================================================================
    # Reset & Actions
    # ================================================================

    def _reset_session(self):
        """Reset all state to start a new session."""
        self.reg_filepath.set("")
        self.zoom_filepath.set("")
        self.reg_col_map = None
        self.reg_df = None
        self.zoom_col_map = None
        self.output_dir = None

        # Reset file displays
        self.reg_filename_label.config(text="No file selected", fg=MUTED_COLOR)
        self.reg_badge.config(text=" \u25CB Not loaded ", bg=BADGE_PENDING, fg=MUTED_COLOR)
        self.zoom_filename_label.config(text="No file selected", fg=MUTED_COLOR)
        self.zoom_badge.config(text=" \u25CB Not loaded ", bg=BADGE_PENDING, fg=MUTED_COLOR)

        # Reset settings
        self.attendance_day.set(1)
        self._on_segment_click(1)
        self.program_var.set("All")
        self.program_dropdown['values'] = ["All"]
        self.date_picker.set_date(date.today())

        # Reset steps
        self._update_step(0, "active")
        self._update_step(1, "inactive")
        self._update_step(2, "inactive")

        # Hide summary
        self._hide_summary()

        # Clear log
        self._clear_log()
        self._log("Session reset. Load your files to begin.", 'info')

        # Disable output folder
        self.open_folder_btn.config(state='disabled')

    def _open_output_folder(self):
        """Open the output folder in file explorer."""
        if self.output_dir and os.path.isdir(self.output_dir):
            if sys.platform == 'win32':
                os.startfile(self.output_dir)
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', self.output_dir])
            else:
                subprocess.Popen(['xdg-open', self.output_dir])
