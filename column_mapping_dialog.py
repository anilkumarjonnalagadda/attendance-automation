"""
Column Mapping Dialog for Attendance Automation Tool.

A modal Tkinter dialog that lets trainers map CSV columns to
standardized field names. Auto-detection pre-fills the mapping,
but trainers can override any field.

Styled with the shared Slate Teal theme.
"""

import os
import tkinter as tk
from tkinter import ttk

from theme import (
    BG_COLOR, HEADER_BG, HEADER_FG, ACCENT_COLOR, ACCENT_HOVER,
    SUCCESS_COLOR, ERROR_COLOR, WARNING_COLOR,
    CARD_BG, TEXT_COLOR, TEXT_SECONDARY, MUTED_COLOR,
    BORDER_COLOR, SHADOW_COLOR,
    FONT_HEADING, FONT_SUBHEAD, FONT_BODY, FONT_BODY_BOLD,
    FONT_SMALL, FONT_BUTTON,
    ToolTip, make_hover_button,
)


# Sentinel value for unmapped fields
NOT_MAPPED = "-- Not Mapped --"


class ColumnMappingDialog:
    """
    Modal dialog for mapping CSV columns to standardized field names.

    Usage:
        dialog = ColumnMappingDialog(
            parent=root,
            title="Column Mapping - Registration CSV",
            filename="registration_data.csv",
            csv_columns=list(df.columns),
            field_specs=REGISTRATION_FIELDS,
            auto_detected=detect_registration_columns(df),
        )
        # dialog.result is None if cancelled, or a col_map dict if confirmed
        confirmed_col_map = dialog.result
    """

    def __init__(self, parent, title, filename, csv_columns, field_specs, auto_detected):
        """
        Args:
            parent: Parent Tk widget (for modality).
            title: Window title string.
            filename: Name of the CSV file (for display).
            csv_columns: List of all column header strings from the CSV.
            field_specs: List of field spec dicts, each with 'key', 'label', 'required'.
            auto_detected: Dict from auto-detection {standardized_key: csv_column_name}.
        """
        self.parent = parent
        self.csv_columns = csv_columns
        self.field_specs = field_specs
        self.auto_detected = auto_detected
        self.result = None  # None = cancelled, dict = confirmed mapping

        # Store combo references and indicator labels keyed by field key
        self.combos = {}
        self.indicators = {}

        self._build_dialog(title, filename)

    def _build_dialog(self, title, filename):
        """Build and display the modal dialog."""
        self.top = tk.Toplevel(self.parent)
        self.top.title(title)
        self.top.configure(bg=BG_COLOR)
        self.top.resizable(False, False)

        # Make modal
        self.top.transient(self.parent)
        self.top.grab_set()

        # --- Header (two-line: title + filename) ---
        header = tk.Frame(self.top, bg=HEADER_BG, height=64)
        header.pack(fill='x')
        header.pack_propagate(False)

        header_inner = tk.Frame(header, bg=HEADER_BG)
        header_inner.place(relx=0.5, rely=0.5, anchor='center')

        tk.Label(
            header_inner, text=title,
            font=("Segoe UI", 14, "bold"),
            bg=HEADER_BG, fg=HEADER_FG,
        ).pack()

        tk.Label(
            header_inner, text=filename,
            font=FONT_SMALL,
            bg=HEADER_BG, fg=MUTED_COLOR,
        ).pack()

        # --- Content area ---
        content = tk.Frame(self.top, bg=BG_COLOR, padx=22, pady=16)
        content.pack(fill='both', expand=True)

        # Info text
        tk.Label(
            content,
            text=f"Detected {len(self.csv_columns)} columns  •  Map them to the required fields below",
            font=FONT_BODY,
            bg=BG_COLOR, fg=TEXT_SECONDARY,
            anchor='w',
        ).pack(fill='x', pady=(0, 4))

        tk.Label(
            content,
            text="Fields marked with * are required.",
            font=FONT_SMALL,
            bg=BG_COLOR, fg=MUTED_COLOR,
            anchor='w',
        ).pack(fill='x', pady=(0, 12))

        # --- Mapping Card (shadow wrapper) ---
        shadow_frame = tk.Frame(content, bg=SHADOW_COLOR)
        shadow_frame.pack(fill='x', padx=(0, 2), pady=(0, 2))

        card = tk.Frame(shadow_frame, bg=CARD_BG, relief='flat', bd=0, padx=16, pady=12)
        card.pack(fill='both')

        # Dropdown options: "-- Not Mapped --" + all CSV columns
        dropdown_values = [NOT_MAPPED] + list(self.csv_columns)

        for i, spec in enumerate(self.field_specs):
            row = tk.Frame(card, bg=CARD_BG)
            row.pack(fill='x', pady=3)

            # Teal left-border bar for required fields
            if spec['required']:
                accent_bar = tk.Frame(row, bg=ACCENT_COLOR, width=3)
                accent_bar.pack(side='left', fill='y', padx=(0, 8))
            else:
                spacer = tk.Frame(row, bg=CARD_BG, width=11)
                spacer.pack(side='left')

            # Label
            label_text = spec['label']
            tk.Label(
                row, text=label_text,
                font=FONT_BODY_BOLD if spec['required'] else FONT_BODY,
                bg=CARD_BG, fg=TEXT_COLOR,
                width=22, anchor='w',
            ).pack(side='left')

            # Per-row status indicator
            indicator = tk.Label(
                row, text="○",
                font=("Segoe UI", 11),
                bg=CARD_BG, fg=MUTED_COLOR,
                width=2,
            )
            indicator.pack(side='left', padx=(0, 6))
            self.indicators[spec['key']] = indicator

            # Combobox
            combo = ttk.Combobox(
                row,
                values=dropdown_values,
                state='readonly',
                font=FONT_BODY,
                width=32,
            )
            combo.pack(side='left', padx=(0, 5))

            # Pre-fill from auto-detection
            auto_value = self.auto_detected.get(spec['key'])
            if auto_value and auto_value in self.csv_columns:
                combo.set(auto_value)
            else:
                combo.set(NOT_MAPPED)

            # Bind change event
            combo.bind("<<ComboboxSelected>>", self._on_combo_change)

            self.combos[spec['key']] = combo

            # Thin separator between rows (not after last)
            if i < len(self.field_specs) - 1:
                tk.Frame(card, bg=BORDER_COLOR, height=1).pack(fill='x', pady=2)

        # --- Status Label ---
        self.status_label = tk.Label(
            content, text="",
            font=FONT_BODY,
            bg=BG_COLOR, anchor='w',
        )
        self.status_label.pack(fill='x', pady=(12, 6))

        # --- Buttons ---
        btn_frame = tk.Frame(content, bg=BG_COLOR)
        btn_frame.pack(fill='x', pady=(4, 0))

        # Reset button (hover)
        reset_btn = make_hover_button(
            btn_frame, text="↻ Reset to Auto-Detect",
            bg=BORDER_COLOR, fg=TEXT_COLOR, hover_bg=SHADOW_COLOR,
            font=FONT_BUTTON,
            command=self._reset_auto_detect,
            padx=14, pady=6,
        )
        reset_btn.pack(side='left')
        ToolTip(reset_btn, "Reset all mappings to auto-detected values")

        # Confirm button (hover, starts disabled if required fields missing)
        self.confirm_btn = make_hover_button(
            btn_frame, text="✓ Confirm Mapping",
            bg=ACCENT_COLOR, fg="white", hover_bg=ACCENT_HOVER,
            font=("Segoe UI", 11, "bold"),
            command=self._on_confirm,
            padx=22, pady=6,
        )
        self.confirm_btn.pack(side='right')
        ToolTip(self.confirm_btn, "Confirm the column mapping and continue")

        # Cancel button (hover)
        cancel_btn = make_hover_button(
            btn_frame, text="Cancel",
            bg=BORDER_COLOR, fg=TEXT_COLOR, hover_bg=SHADOW_COLOR,
            font=FONT_BUTTON,
            command=self._on_cancel,
            padx=14, pady=6,
        )
        cancel_btn.pack(side='right', padx=(0, 8))
        ToolTip(cancel_btn, "Cancel and return without saving mappings")

        # Handle window close
        self.top.protocol("WM_DELETE_WINDOW", self._on_cancel)

        # Initial status update
        self._update_status()

        # Center dialog on parent
        self.top.update_idletasks()
        w = self.top.winfo_width()
        h = self.top.winfo_height()
        px = self.parent.winfo_rootx() + (self.parent.winfo_width() // 2) - (w // 2)
        py = self.parent.winfo_rooty() + (self.parent.winfo_height() // 2) - (h // 2)
        self.top.geometry(f"+{px}+{py}")

        # Block until dialog closes
        self.parent.wait_window(self.top)

    def _on_combo_change(self, event=None):
        """Called when any combobox selection changes."""
        self._update_status()

    def _update_status(self):
        """Update per-row indicators, status label, and confirm button state."""
        # --- Gather state ---
        missing_required = []
        mapped_cols = {}
        duplicates = []
        duplicate_keys = set()

        for spec in self.field_specs:
            combo = self.combos[spec['key']]
            val = combo.get()

            if val != NOT_MAPPED:
                if val in mapped_cols:
                    duplicates.append(
                        f"'{val}' mapped to both {mapped_cols[val]} and {spec['label']}"
                    )
                    duplicate_keys.add(spec['key'])
                    # Also mark the first field that used this column
                    for s in self.field_specs:
                        if self.combos[s['key']].get() == val and s['key'] != spec['key']:
                            duplicate_keys.add(s['key'])
                else:
                    mapped_cols[val] = spec['label']

            if spec['required'] and val == NOT_MAPPED:
                missing_required.append(spec['label'].replace(' *', ''))

        # --- Update per-row indicators ---
        for spec in self.field_specs:
            key = spec['key']
            indicator = self.indicators[key]
            val = self.combos[key].get()

            if key in duplicate_keys:
                # Duplicate — amber warning
                indicator.config(text="⚠", fg=WARNING_COLOR)
            elif val != NOT_MAPPED:
                # Mapped — green check
                indicator.config(text="✓", fg=SUCCESS_COLOR)
            else:
                # Not mapped — gray circle
                indicator.config(text="○", fg=MUTED_COLOR)

        # --- Update status text and confirm button ---
        if missing_required:
            self.status_label.config(
                text=f"⚠ Missing required: {', '.join(missing_required)}",
                fg=ERROR_COLOR,
            )
            self.confirm_btn.config(state='disabled', bg=MUTED_COLOR)
        elif duplicates:
            self.status_label.config(
                text=f"⚠ {duplicates[0]}",
                fg=WARNING_COLOR,
            )
            # Allow confirm even with duplicates (edge cases may need it)
            self.confirm_btn.config(state='normal', bg=ACCENT_COLOR)
        else:
            mapped_count = sum(
                1 for spec in self.field_specs
                if self.combos[spec['key']].get() != NOT_MAPPED
            )
            total = len(self.field_specs)
            self.status_label.config(
                text=f"✓ All required fields mapped ({mapped_count}/{total} configured)",
                fg=SUCCESS_COLOR,
            )
            self.confirm_btn.config(state='normal', bg=ACCENT_COLOR)

    def _reset_auto_detect(self):
        """Reset all combos to auto-detected values."""
        for spec in self.field_specs:
            combo = self.combos[spec['key']]
            auto_value = self.auto_detected.get(spec['key'])
            if auto_value and auto_value in self.csv_columns:
                combo.set(auto_value)
            else:
                combo.set(NOT_MAPPED)
        self._update_status()

    def _on_confirm(self):
        """Build the col_map from current selections and close."""
        col_map = {}
        for spec in self.field_specs:
            combo = self.combos[spec['key']]
            val = combo.get()
            if val != NOT_MAPPED:
                col_map[spec['key']] = val

        self.result = col_map
        self.top.destroy()

    def _on_cancel(self):
        """Cancel and close without setting a result."""
        self.result = None
        self.top.destroy()
