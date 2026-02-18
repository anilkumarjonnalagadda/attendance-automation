"""
Attendance Automation Tool - Entry Point

A desktop application that automates the process of marking
trainer attendance by cross-referencing registration data
with Zoom attendance records.

Usage:
    python main.py
"""

import sys
import os
import ctypes
import tkinter as tk
from tkinter import ttk

# DPI awareness for crisp rendering on high-DPI Windows displays
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

# Ensure the script's directory is in the path for imports
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

from gui import AttendanceApp


def configure_ttk_style():
    """Configure ttk theme and widget styles for the Slate Teal theme."""
    style = ttk.Style()
    style.theme_use('clam')

    # Combobox styling
    style.configure('TCombobox',
        fieldbackground='white',
        background='#0d9488',
        foreground='#0f172a',
        arrowcolor='#0d9488',
        borderwidth=1,
        relief='solid',
    )
    style.map('TCombobox',
        fieldbackground=[('readonly', 'white')],
        selectbackground=[('readonly', 'white')],
        selectforeground=[('readonly', '#0f172a')],
    )

    # Progressbar styling
    style.configure('Teal.Horizontal.TProgressbar',
        troughcolor='#e2e8f0',
        background='#0d9488',
        thickness=8,
    )

    # Button styling
    style.configure('Teal.TButton',
        background='#0d9488',
        foreground='white',
        font=('Segoe UI', 10),
        borderwidth=0,
        relief='flat',
    )
    style.map('Teal.TButton',
        background=[('active', '#0f766e')],
    )


def main():
    root = tk.Tk()

    # Configure ttk styles
    configure_ttk_style()

    # Set application icon (if available)
    try:
        root.iconbitmap(os.path.join(script_dir, 'icon.ico'))
    except Exception:
        pass

    # Center the window on screen
    root.update_idletasks()
    width = 720
    height = 820
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')

    app = AttendanceApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
