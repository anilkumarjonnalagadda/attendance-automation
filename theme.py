"""
Shared theme module for Attendance Automation Tool.
Contains all color constants, font definitions, and reusable UI helpers.
"""

import tkinter as tk


# ============================================================
# Slate Teal Color Palette
# ============================================================

BG_COLOR       = "#f1f5f9"   # Slate-50: soft background
HEADER_BG      = "#0f172a"   # Slate-900: near-black header
HEADER_FG      = "#f8fafc"   # Slate-50: crisp white
ACCENT_COLOR   = "#0d9488"   # Teal-600: primary actions
ACCENT_HOVER   = "#0f766e"   # Teal-700: hover state
ACCENT_LIGHT   = "#ccfbf1"   # Teal-100: highlights/badges
ACCENT_MUTED   = "#5eead4"   # Teal-300: progress bars
SUCCESS_COLOR  = "#16a34a"   # Green-600
ERROR_COLOR    = "#dc2626"   # Red-600
WARNING_COLOR  = "#d97706"   # Amber-600
CARD_BG        = "#ffffff"   # Pure white
TEXT_COLOR      = "#0f172a"  # Slate-900: primary text
TEXT_SECONDARY = "#475569"   # Slate-600: secondary text
MUTED_COLOR    = "#94a3b8"   # Slate-400: disabled/placeholder
BORDER_COLOR   = "#e2e8f0"   # Slate-200: subtle borders
SHADOW_COLOR   = "#cbd5e1"   # Slate-300: card shadows
STEP_INACTIVE  = "#cbd5e1"   # Slate-300: inactive steps
STEP_ACTIVE    = "#0d9488"   # Teal-600: active/completed steps
BADGE_PENDING  = "#f1f5f9"   # Slate-50: pending badge bg
BADGE_LOADED   = "#dcfce7"   # Green-100: loaded badge bg


# ============================================================
# Font Constants
# ============================================================

FONT_HEADING   = ("Segoe UI", 18, "bold")
FONT_SUBHEAD   = ("Segoe UI", 11, "bold")
FONT_BODY      = ("Segoe UI", 10)
FONT_BODY_BOLD = ("Segoe UI", 10, "bold")
FONT_SMALL     = ("Segoe UI", 9)
FONT_MONO      = ("Consolas", 10)
FONT_BUTTON_LG = ("Segoe UI", 12, "bold")
FONT_BUTTON    = ("Segoe UI", 10)


# ============================================================
# ToolTip Class
# ============================================================

class ToolTip:
    """Lightweight hover tooltip for any Tkinter widget."""

    def __init__(self, widget, text, delay=500):
        self.widget = widget
        self.text = text
        self.delay = delay
        self.tip_window = None
        self._after_id = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")

    def _schedule(self, event):
        self._hide()
        self._after_id = self.widget.after(self.delay, self._show)

    def _show(self):
        if self.tip_window:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            tw, text=self.text, font=FONT_SMALL,
            bg=HEADER_BG, fg=HEADER_FG,
            padx=8, pady=4, relief='flat',
        )
        label.pack()

    def _hide(self, event=None):
        if self._after_id:
            self.widget.after_cancel(self._after_id)
            self._after_id = None
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None

    def update_text(self, new_text):
        """Update the tooltip text."""
        self.text = new_text


# ============================================================
# UI Helper Functions
# ============================================================

def make_hover_button(parent, text, bg, fg, hover_bg, hover_fg=None,
                      font=None, command=None, **kwargs):
    """
    Create a tk.Button with hover color transitions.

    Args:
        parent: Parent widget.
        text: Button label.
        bg: Normal background color.
        fg: Normal foreground color.
        hover_bg: Background color on hover.
        hover_fg: Foreground color on hover (defaults to fg).
        font: Font tuple (defaults to FONT_BUTTON).
        command: Click callback.
        **kwargs: Additional Button kwargs (padx, pady, state, etc.)
    """
    hover_fg = hover_fg or fg
    btn = tk.Button(
        parent, text=text, font=font or FONT_BUTTON,
        bg=bg, fg=fg,
        activebackground=hover_bg, activeforeground=hover_fg,
        relief='flat', cursor='hand2',
        command=command, **kwargs,
    )

    def on_enter(e):
        if btn.cget('state') != 'disabled':
            btn.config(bg=hover_bg, fg=hover_fg)

    def on_leave(e):
        if btn.cget('state') != 'disabled':
            btn.config(bg=bg, fg=fg)

    btn.bind("<Enter>", on_enter)
    btn.bind("<Leave>", on_leave)
    return btn


def create_shadow_card(parent, title=None):
    """
    Create a card frame with a simulated drop shadow and optional section title.

    Returns the inner card frame to pack widgets into.
    """
    container = tk.Frame(parent, bg=BG_COLOR)
    container.pack(fill='x', pady=(0, 12))

    # Section title above card
    if title:
        tk.Label(
            container, text=title,
            font=FONT_SUBHEAD,
            bg=BG_COLOR, fg=TEXT_COLOR,
            anchor='w',
        ).pack(fill='x', pady=(0, 6))

    # Shadow layer
    shadow_frame = tk.Frame(container, bg=SHADOW_COLOR)
    shadow_frame.pack(fill='x', padx=(0, 2), pady=(0, 2))

    # Main card frame
    card = tk.Frame(
        shadow_frame, bg=CARD_BG,
        relief='flat', bd=0,
        padx=18, pady=14,
    )
    card.pack(fill='both')

    return container, card
