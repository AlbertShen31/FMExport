#!/usr/bin/env python3
"""
Screen Data Scanner - Extract tabular data from screen and export to CSV
Works on both Mac and Windows
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import mss
import numpy as np
from PIL import Image, ImageTk, ImageEnhance, ImageFilter
import pytesseract
import pandas as pd
import cv2
import re
import os
import sys
from datetime import datetime
import threading
import subprocess
import json

# Configure Tesseract path for Windows
if sys.platform == 'win32':
    # Common Tesseract installation paths on Windows
    tesseract_paths = [
        r'C:\Program Files\Tesseract-OCR\tesseract.exe',
        r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
        r'C:\Users\{}\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'.format(os.getenv('USERNAME', '')),
    ]
    
    for path in tesseract_paths:
        if os.path.exists(path):
            pytesseract.pytesseract.tesseract_cmd = path
            break


class ScreenScannerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Screen Data Scanner")
        self.root.geometry("800x600")
        
        # Variables
        self.captured_image = None
        self.processed_data = None
        self.screenshot_path = None
        self.selected_window = None  # Store selected window info
        
        # Setup GUI
        self.setup_ui()
        
        # Check for Tesseract
        self.check_tesseract()
    
    def setup_ui(self):
        """Create the user interface"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = ttk.Label(main_frame, text="Screen Data Scanner", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Instructions
        instructions = ttk.Label(main_frame, 
                               text="1. Click 'Select Window' to choose a window, or 'Select Area' for manual selection\n"
                                    "2. Preview and adjust if needed\n"
                                    "3. Click 'Extract Data' to process\n"
                                    "4. Click 'Export to CSV' to save",
                               justify=tk.LEFT)
        instructions.grid(row=1, column=0, columnspan=2, pady=(0, 20))
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)
        
        # Select Window button
        self.select_window_btn = ttk.Button(button_frame, text="Select Window", 
                                           command=self.select_window)
        self.select_window_btn.pack(side=tk.LEFT, padx=5)
        
        # Select Area button
        self.select_btn = ttk.Button(button_frame, text="Select Area to Scan", 
                                     command=self.select_area)
        self.select_btn.pack(side=tk.LEFT, padx=5)
        
        # Extract Data button
        self.extract_btn = ttk.Button(button_frame, text="Extract Data", 
                                      command=self.extract_data,
                                      state=tk.DISABLED)
        self.extract_btn.pack(side=tk.LEFT, padx=5)
        
        # Export CSV button
        self.export_btn = ttk.Button(button_frame, text="Export to CSV", 
                                     command=self.export_csv,
                                     state=tk.DISABLED)
        self.export_btn.pack(side=tk.LEFT, padx=5)
        
        # Wage budget estimator frame
        budget_frame = ttk.LabelFrame(main_frame, text="Wage Budget Estimator", padding="10")
        budget_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 10))
        budget_frame.columnconfigure(1, weight=1)
        budget_frame.columnconfigure(3, weight=1)

        ttk.Label(budget_frame, text="Current balance:").grid(row=0, column=0, sticky=tk.W)
        ttk.Label(budget_frame, text="Balance a year ago:").grid(row=1, column=0, sticky=tk.W)
        ttk.Label(budget_frame, text="Prize money (actual):").grid(row=2, column=0, sticky=tk.W)
        ttk.Label(budget_frame, text="Prize money (expected):").grid(row=3, column=0, sticky=tk.W)
        ttk.Label(budget_frame, text="Net transfer spend:").grid(row=0, column=2, sticky=tk.W)
        ttk.Label(budget_frame, text="Current wages:").grid(row=1, column=2, sticky=tk.W)
        ttk.Label(budget_frame, text="Squad size:").grid(row=2, column=2, sticky=tk.W)

        self.current_balance_var = tk.StringVar()
        self.prior_balance_var = tk.StringVar()
        self.prize_actual_var = tk.StringVar()
        self.prize_expected_var = tk.StringVar()
        self.net_transfer_var = tk.StringVar()
        self.current_wages_var = tk.StringVar()
        self.squad_size_var = tk.StringVar(value="25")
        self.wage_period_var = tk.StringVar(value="annual")

        self.current_balance_entry = ttk.Entry(budget_frame, textvariable=self.current_balance_var, width=18)
        self.current_balance_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(5, 15))
        self.prior_balance_entry = ttk.Entry(budget_frame, textvariable=self.prior_balance_var, width=18)
        self.prior_balance_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(5, 15))
        self.prize_actual_entry = ttk.Entry(budget_frame, textvariable=self.prize_actual_var, width=18)
        self.prize_actual_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(5, 15))
        self.prize_expected_entry = ttk.Entry(budget_frame, textvariable=self.prize_expected_var, width=18)
        self.prize_expected_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), padx=(5, 15))
        self.net_transfer_entry = ttk.Entry(budget_frame, textvariable=self.net_transfer_var, width=18)
        self.net_transfer_entry.grid(row=0, column=3, sticky=(tk.W, tk.E), padx=(5, 0))
        self.current_wages_entry = ttk.Entry(budget_frame, textvariable=self.current_wages_var, width=18)
        self.current_wages_entry.grid(row=1, column=3, sticky=(tk.W, tk.E), padx=(5, 0))
        self.squad_size_entry = ttk.Entry(budget_frame, textvariable=self.squad_size_var, width=18)
        self.squad_size_entry.grid(row=2, column=3, sticky=(tk.W, tk.E), padx=(5, 0))

        wage_period_label = ttk.Label(budget_frame, text="Wage period:")
        wage_period_label.grid(row=3, column=2, sticky=tk.W, pady=(4, 0))
        wage_period_frame = ttk.Frame(budget_frame)
        wage_period_frame.grid(row=3, column=3, sticky=tk.W, pady=(4, 0))
        ttk.Radiobutton(
            wage_period_frame,
            text="Weekly",
            variable=self.wage_period_var,
            value="weekly",
            command=self.calculate_wage_budget,
        ).pack(side=tk.LEFT)
        ttk.Radiobutton(
            wage_period_frame,
            text="Annual",
            variable=self.wage_period_var,
            value="annual",
            command=self.calculate_wage_budget,
        ).pack(side=tk.LEFT, padx=(8, 0))

        self._money_trace_lock = False
        for entry, var in (
            (self.current_balance_entry, self.current_balance_var),
            (self.prior_balance_entry, self.prior_balance_var),
            (self.prize_actual_entry, self.prize_actual_var),
            (self.prize_expected_entry, self.prize_expected_var),
            (self.net_transfer_entry, self.net_transfer_var),
            (self.current_wages_entry, self.current_wages_var),
        ):
            var.trace_add("write", lambda *args, v=var, e=entry: self._format_money_live(v, e))
        self.squad_size_var.trace_add(
            "write",
            lambda *args: self._format_int_live(self.squad_size_var, self.squad_size_entry),
        )

        calc_btn = ttk.Button(budget_frame, text="Calculate", command=self.calculate_wage_budget)
        calc_btn.grid(row=4, column=0, columnspan=4, pady=(8, 6))

        self.current_balance_display_var = tk.StringVar(value="Current balance: -")
        self.expected_balance_var = tk.StringVar(value="Expected balance (1y): -")
        self.profit_var = tk.StringVar(value="Profit: -")
        self.adjusted_profit_var = tk.StringVar(value="Adjusted profit: -")
        self.total_budget_var = tk.StringVar(value="Available wages: -")
        self.current_wages_display_var = tk.StringVar(value="Current wages: -")
        self.expected_wages_var = tk.StringVar(value="Expected wages budget: -")
        self.avg_spend_var = tk.StringVar(value="Average per player: -")
        self.max_spend_var = tk.StringVar(value="Max per player (2x avg): -")

        ttk.Label(budget_frame, textvariable=self.current_balance_display_var).grid(row=5, column=0, columnspan=2, sticky=tk.W)
        ttk.Label(budget_frame, textvariable=self.expected_balance_var).grid(row=5, column=2, columnspan=2, sticky=tk.W)
        ttk.Label(budget_frame, textvariable=self.profit_var).grid(row=6, column=0, columnspan=2, sticky=tk.W)
        ttk.Label(budget_frame, textvariable=self.adjusted_profit_var).grid(row=6, column=2, columnspan=2, sticky=tk.W)
        ttk.Label(budget_frame, textvariable=self.total_budget_var).grid(row=7, column=0, columnspan=2, sticky=tk.W)
        ttk.Label(budget_frame, textvariable=self.current_wages_display_var).grid(row=7, column=2, columnspan=2, sticky=tk.W)
        ttk.Label(budget_frame, textvariable=self.expected_wages_var).grid(row=8, column=0, columnspan=2, sticky=tk.W)
        ttk.Label(budget_frame, textvariable=self.avg_spend_var).grid(row=8, column=2, columnspan=2, sticky=tk.W)
        ttk.Label(budget_frame, textvariable=self.max_spend_var).grid(row=9, column=0, columnspan=2, sticky=tk.W)

        # Preview frame
        preview_label = ttk.Label(main_frame, text="Preview:", font=("Arial", 10, "bold"))
        preview_label.grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=(10, 5))
        
        # Canvas for image preview
        self.canvas = tk.Canvas(main_frame, width=750, height=400, bg="white", 
                               relief=tk.SUNKEN, borderwidth=2)
        self.canvas.grid(row=5, column=0, columnspan=2, pady=10)
        
        # Scrollbar for canvas
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        scrollbar.grid(row=5, column=2, sticky=(tk.N, tk.S))
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready - Click 'Select Area' to begin")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                              relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=7, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=1)

    def _parse_money_input(self, value):
        """Parse money input allowing commas and currency symbols."""
        if value is None:
            return 0.0
        cleaned = re.sub(r'[^0-9.\-]', '', str(value))
        if cleaned in ('', '-', '.', '-.'):
            return 0.0
        try:
            return float(cleaned)
        except ValueError:
            return 0.0

    def _format_money(self, value):
        """Format money value with commas and sensible decimals."""
        if float(value).is_integer():
            return f"{value:,.0f}"
        return f"{value:,.2f}"

    def _format_money_var(self, var):
        """Format a money StringVar in-place with commas."""
        raw = (var.get() or "").strip()
        if raw == "":
            return
        value = self._parse_money_input(raw)
        var.set(self._format_money(value))

    def _normalize_money_input(self, raw):
        """Keep only digits, one decimal point, and leading minus."""
        cleaned = re.sub(r'[^0-9.\-]', '', raw)
        if not cleaned:
            return ""
        negative = cleaned.startswith('-')
        cleaned = cleaned.replace('-', '')
        if cleaned.count('.') > 1:
            parts = cleaned.split('.')
            cleaned = parts[0] + '.' + ''.join(parts[1:])
        cleaned = ('-' if negative else '') + cleaned
        return cleaned

    def _format_money_live(self, var, entry):
        """Live format input with commas while typing."""
        if self._money_trace_lock:
            return
        raw = var.get()
        if raw is None:
            return
        normalized = self._normalize_money_input(raw)
        if normalized in ("", "-", "."):
            return
        value = self._parse_money_input(normalized)
        formatted = self._format_money(value)
        if formatted == raw:
            return
        self._money_trace_lock = True
        try:
            var.set(formatted)
            if entry is not None:
                entry.icursor(tk.END)
        finally:
            self._money_trace_lock = False

    def _format_int_live(self, var, entry):
        """Live format integer input while typing."""
        if self._money_trace_lock:
            return
        raw = var.get()
        if raw is None:
            return
        cleaned = re.sub(r'[^0-9]', '', str(raw))
        if cleaned == "":
            return
        if cleaned == raw:
            return
        self._money_trace_lock = True
        try:
            var.set(cleaned)
            if entry is not None:
                entry.icursor(tk.END)
        finally:
            self._money_trace_lock = False

    def calculate_wage_budget(self):
        """Calculate wage budget estimates for a 25-player squad."""
        current_balance = self._parse_money_input(self.current_balance_var.get())
        prior_raw = (self.prior_balance_var.get() or "").strip()
        if prior_raw == "":
            prior_balance = current_balance
        else:
            prior_balance = self._parse_money_input(prior_raw)
        prize_actual = self._parse_money_input(self.prize_actual_var.get())
        prize_expected = self._parse_money_input(self.prize_expected_var.get())
        net_transfer_spend = self._parse_money_input(self.net_transfer_var.get())
        current_wages_input = self._parse_money_input(self.current_wages_var.get())
        squad_size_raw = (self.squad_size_var.get() or "").strip()
        if squad_size_raw == "":
            squad_size = 25
        else:
            try:
                squad_size = max(1, int(squad_size_raw))
            except ValueError:
                squad_size = 25
        wage_period = self.wage_period_var.get()
        if wage_period == "weekly":
            current_wages_annual = current_wages_input * 52
        else:
            current_wages_annual = current_wages_input
        current_wages_weekly = current_wages_annual / 52 if current_wages_annual else 0.0

        year_profit = current_balance - prior_balance
        adjusted_profit = (
            year_profit - prize_actual + prize_expected - net_transfer_spend
        )
        base_budget = adjusted_profit + current_wages_annual
        minimum_budget = current_wages_annual * 0.75
        if base_budget < minimum_budget:
            base_budget = minimum_budget

        total_available = base_budget

        avg_per_player = total_available / squad_size if total_available else 0.0
        max_per_player = avg_per_player * 2

        expected_balance = (
            current_balance
            + adjusted_profit
            - (total_available - current_wages_annual)
        )
        expected_wages_weekly = total_available / 52 if total_available else 0.0

        if wage_period == "weekly":
            display_factor = 1 / 52
            period_label = "weekly"
        else:
            display_factor = 1
            period_label = "annual"

        self.current_balance_display_var.set(
            f"Current balance: {self._format_money(current_balance)}"
        )
        self.expected_balance_var.set(
            f"Expected balance (1y): {self._format_money(expected_balance)}"
        )
        self.profit_var.set(f"Profit: {self._format_money(year_profit)}")
        self.adjusted_profit_var.set(
            f"Adjusted profit: {self._format_money(adjusted_profit)}"
        )
        self.total_budget_var.set(
            f"Available wages ({period_label}): "
            f"{self._format_money(total_available * display_factor)}"
        )
        self.current_wages_display_var.set(
            f"Current wages ({period_label}): "
            f"{self._format_money((current_wages_weekly if wage_period == 'weekly' else current_wages_annual))}"
        )
        self.expected_wages_var.set(
            f"Expected wages budget ({period_label}): "
            f"{self._format_money((expected_wages_weekly if wage_period == 'weekly' else total_available))}"
        )
        self.avg_spend_var.set(
            f"Average per player ({period_label}, {squad_size}): {self._format_money(avg_per_player * display_factor)}"
        )
        self.max_spend_var.set(
            f"Max per player (2x avg, {period_label}): {self._format_money(max_per_player * display_factor)}"
        )
    
    def check_tesseract(self):
        """Check if Tesseract OCR is installed"""
        try:
            pytesseract.get_tesseract_version()
            self.status_var.set("Ready - Tesseract OCR detected")
        except Exception as e:
            self.status_var.set("Warning: Tesseract OCR not found. Please install Tesseract.")
            messagebox.showwarning(
                "Tesseract Not Found",
                "Tesseract OCR is required for this application.\n\n"
                "Installation:\n"
                "Mac: brew install tesseract\n"
                "Windows: Download from https://github.com/UB-Mannheim/tesseract/wiki"
            )
    
    def get_windows_list(self):
        """Get list of open windows (platform-specific)"""
        windows = []
        
        if sys.platform == 'darwin':  # macOS
            try:
                # Use AppleScript to get window list
                script = '''
                tell application "System Events"
                    set windowList to {}
                    repeat with proc in processes
                        try
                            set procName to name of proc
                            repeat with win in windows of proc
                                try
                                    set winTitle to title of win
                                    set winPos to position of win
                                    set winSize to size of win
                                    set end of windowList to {procName, winTitle, winPos, winSize}
                                end try
                            end repeat
                        end try
                    end repeat
                    return windowList
                end tell
                '''
                result = subprocess.run(
                    ['osascript', '-e', script],
                    capture_output=True,
                    text=True,
                    timeout=10
                )
                
                if result.returncode == 0:
                    # Parse AppleScript output (it's a bit complex, so we'll use a simpler approach)
                    # Let's use a different AppleScript that returns JSON-like data
                    script2 = '''
                    tell application "System Events"
                        set windowInfo to {}
                        repeat with proc in processes
                            try
                                set procName to name of proc
                                if procName is not "Screen Data Scanner" and procName is not "SystemUIServer" then
                                    repeat with win in windows of proc
                                        try
                                            set winTitle to title of win
                                            if winTitle is not "" then
                                                set end of windowInfo to procName & " - " & winTitle
                                            end if
                                        end try
                                    end repeat
                                end if
                            end try
                        end repeat
                        return windowInfo
                    end tell
                    '''
                    result = subprocess.run(
                        ['osascript', '-e', script2],
                        capture_output=True,
                        text=True,
                        timeout=10
                    )
                    
                    if result.returncode == 0:
                        lines = result.stdout.strip().split(', ')
                        for line in lines:
                            if line and line.strip():
                                windows.append({
                                    'title': line.strip(),
                                    'app': line.split(' - ')[0] if ' - ' in line else line
                                })
            except Exception as e:
                print(f"Error getting windows on macOS: {e}")
        
        elif sys.platform == 'win32':  # Windows
            try:
                # Try using pywin32 if available
                try:
                    import win32gui
                    import win32process
                    import win32api
                    
                    def enum_windows_callback(hwnd, windows):
                        if win32gui.IsWindowVisible(hwnd):
                            title = win32gui.GetWindowText(hwnd)
                            if title:
                                try:
                                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                                    handle = win32api.OpenProcess(0x0410, False, pid)
                                    exe_path = win32process.GetModuleFileNameEx(handle, 0)
                                    app_name = os.path.basename(exe_path)
                                    windows.append({
                                        'title': title,
                                        'app': app_name,
                                        'hwnd': hwnd
                                    })
                                except:
                                    windows.append({
                                        'title': title,
                                        'app': 'Unknown',
                                        'hwnd': hwnd
                                    })
                        return True
                    
                    win32gui.EnumWindows(enum_windows_callback, windows)
                except ImportError:
                    # Fallback: use a simple method
                    messagebox.showinfo(
                        "Window Selection",
                        "For better window selection on Windows, install pywin32:\n"
                        "uv pip install pywin32\n\n"
                        "Falling back to area selection."
                    )
                    return []
            except Exception as e:
                print(f"Error getting windows on Windows: {e}")
        
        return windows
    
    def select_window(self):
        """Show window selection dialog"""
        self.status_var.set("Getting list of windows...")
        windows = self.get_windows_list()
        
        if not windows:
            messagebox.showwarning(
                "No Windows Found",
                "Could not retrieve window list. Please use 'Select Area' instead."
            )
            self.status_var.set("Ready - Use 'Select Area' to manually select region")
            return
        
        # Create window selection dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Select Window to Scan")
        dialog.geometry("600x500")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # Label
        label = ttk.Label(dialog, text="Select a window to scan:", font=("Arial", 10, "bold"))
        label.pack(pady=10)
        
        # Listbox with scrollbar
        frame = ttk.Frame(dialog)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        scrollbar = ttk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        listbox = tk.Listbox(frame, yscrollcommand=scrollbar.set, font=("Arial", 10))
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        
        # Populate listbox
        for window in windows:
            display_text = f"{window['app']} - {window['title']}"
            listbox.insert(tk.END, display_text)
        
        if not listbox.size():
            listbox.insert(tk.END, "No windows found")
            listbox.config(state=tk.DISABLED)
        
        # Buttons
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        
        def on_select():
            selection = listbox.curselection()
            if selection:
                selected_index = selection[0]
                self.selected_window = windows[selected_index]
                dialog.destroy()
                self.capture_selected_window()
            else:
                messagebox.showinfo("Selection", "Please select a window")
        
        def on_cancel():
            dialog.destroy()
        
        select_btn = ttk.Button(button_frame, text="Select", command=on_select)
        select_btn.pack(side=tk.LEFT, padx=5)
        
        cancel_btn = ttk.Button(button_frame, text="Cancel", command=on_cancel)
        cancel_btn.pack(side=tk.LEFT, padx=5)
        
        # Double-click to select
        def on_double_click(event):
            on_select()
        
        listbox.bind("<Double-Button-1>", on_double_click)
        
        # Focus on listbox
        listbox.focus_set()
        if listbox.size() > 0:
            listbox.selection_set(0)
    
    def capture_selected_window(self):
        """Capture the selected window"""
        if not self.selected_window:
            return
        
        self.status_var.set("Capturing selected window...")
        
        try:
            if sys.platform == 'darwin':  # macOS
                try:
                    from Quartz import CGWindowListCopyWindowInfo, kCGWindowListOptionOnScreenOnly, kCGNullWindowID, kCGWindowListExcludeDesktopElements
                    import Quartz
                    
                    # Get window bounds using AppleScript
                    window_title = self.selected_window['title']
                    app_name = self.selected_window['app']
                    
                    script = f'''
                    tell application "System Events"
                        tell process "{app_name}"
                            repeat with win in windows
                                try
                                    if title of win is "{window_title}" then
                                        set winPos to position of win
                                        set winSize to size of win
                                        return (item 1 of winPos) & "," & (item 2 of winPos) & "," & (item 1 of winSize) & "," & (item 2 of winSize)
                                    end if
                                end try
                            end repeat
                        end tell
                    end tell
                    '''
                    
                    result = subprocess.run(
                        ['osascript', '-e', script],
                        capture_output=True,
                        text=True,
                        timeout=5
                    )
                    
                    if result.returncode == 0 and result.stdout.strip():
                        coords = result.stdout.strip().split(',')
                        if len(coords) == 4:
                            left = int(coords[0].strip())
                            top = int(coords[1].strip())
                            width = int(coords[2].strip())
                            height = int(coords[3].strip())
                            
                            # Capture using mss
                            with mss.mss() as sct:
                                region = {
                                    "top": top,
                                    "left": left,
                                    "width": width,
                                    "height": height
                                }
                                screenshot = sct.grab(region)
                                img = Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")
                                
                                self.captured_image = img
                                self.display_preview(img)
                                self.extract_btn.config(state=tk.NORMAL)
                                self.status_var.set(f"Window captured: {img.width}x{img.height} pixels")
                                self.selected_window = None
                                return
                    
                    # Fallback: show message and use area selection
                    messagebox.showinfo(
                        "Window Selected",
                        f"Selected: {self.selected_window['title']}\n\n"
                        "Please use 'Select Area' to capture the window content.\n"
                        "The window selection helps identify which window to scan."
                    )
                    self.selected_window = None
                    self.select_area()
                    
                except ImportError:
                    messagebox.showinfo(
                        "Window Selected",
                        f"Selected: {self.selected_window['title']}\n\n"
                        "Please use 'Select Area' to capture the window content.\n"
                        "Install pyobjc-framework-Quartz for automatic window capture:\n"
                        "uv pip install pyobjc-framework-Quartz"
                    )
                    self.selected_window = None
                    self.select_area()
                except Exception as e:
                    print(f"Error capturing window on macOS: {e}")
                    messagebox.showinfo(
                        "Window Selected",
                        f"Selected: {self.selected_window['title']}\n\n"
                        "Please use 'Select Area' to capture the window content."
                    )
                    self.selected_window = None
                    self.select_area()
                
            elif sys.platform == 'win32':  # Windows
                try:
                    import win32gui
                    hwnd = self.selected_window.get('hwnd')
                    if hwnd:
                        # Get window rectangle
                        rect = win32gui.GetWindowRect(hwnd)
                        left, top, right, bottom = rect
                        
                        # Capture the window
                        with mss.mss() as sct:
                            region = {
                                "top": top,
                                "left": left,
                                "width": right - left,
                                "height": bottom - top
                            }
                            screenshot = sct.grab(region)
                            img = Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")
                            
                            self.captured_image = img
                            self.display_preview(img)
                            self.extract_btn.config(state=tk.NORMAL)
                            self.status_var.set(f"Window captured: {img.width}x{img.height} pixels")
                            self.selected_window = None
                    else:
                        raise Exception("Window handle not available")
                except ImportError:
                    messagebox.showwarning(
                        "PyWin32 Required",
                        "Please install pywin32 for window capture:\n"
                        "uv pip install pywin32"
                    )
                    self.selected_window = None
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to capture window: {str(e)}")
                    self.status_var.set("Error capturing window")
                    self.selected_window = None
            else:
                # Fallback to area selection
                messagebox.showinfo(
                    "Window Selected",
                    f"Selected: {self.selected_window['title']}\n\n"
                    "Please use 'Select Area' to capture the window content."
                )
                self.selected_window = None
                self.select_area()
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to capture window: {str(e)}")
            self.status_var.set("Error capturing window")
            self.selected_window = None
    
    def select_area(self):
        """Open area selection window"""
        self.status_var.set("Selecting area... Click and drag to select region")
        self.root.withdraw()  # Hide main window
        
        # Wait a moment for window to hide
        self.root.update()
        import time
        time.sleep(0.1)  # Small delay to ensure window is hidden
        
        # Get all monitor information
        with mss.mss() as sct:
            monitors = sct.monitors
            # Use monitor 0 (all monitors combined) or monitor 1 (primary)
            all_monitors = monitors[0]
        
        # Create selection window covering all screens
        selection_window = tk.Toplevel()
        selection_window.overrideredirect(True)  # Remove window decorations
        selection_window.attributes('-alpha', 0.2)
        selection_window.configure(bg='black')
        selection_window.attributes('-topmost', True)
        selection_window.attributes('-fullscreen', True)
        
        # Variables for selection
        start_x_root = start_y_root = 0
        start_x_canvas = start_y_canvas = 0
        rect_id = None
        is_selecting = False
        
        canvas = tk.Canvas(selection_window, highlightthickness=0, cursor="cross", bg='black')
        canvas.pack(fill=tk.BOTH, expand=True)
        
        def on_button_press(event):
            nonlocal start_x_root, start_y_root, start_x_canvas, start_y_canvas, rect_id, is_selecting
            # Store absolute screen coordinates for capture
            start_x_root = event.x_root
            start_y_root = event.y_root
            # Store canvas coordinates for drawing
            start_x_canvas = event.x
            start_y_canvas = event.y
            is_selecting = True
            if rect_id:
                canvas.delete(rect_id)
        
        def on_move_press(event):
            nonlocal rect_id, is_selecting
            if not is_selecting:
                return
            if rect_id:
                canvas.delete(rect_id)
            # Draw rectangle using canvas coordinates
            rect_id = canvas.create_rectangle(
                start_x_canvas, start_y_canvas, event.x, event.y,
                outline='red', width=3, fill='', stipple='gray25'
            )
        
        def on_button_release(event):
            nonlocal start_x_root, start_y_root, is_selecting
            if not is_selecting:
                return
            is_selecting = False
            
            # Use absolute screen coordinates for capture
            end_x_root = event.x_root
            end_y_root = event.y_root
            
            # Calculate region bounds
            x1 = min(start_x_root, end_x_root)
            y1 = min(start_y_root, end_y_root)
            x2 = max(start_x_root, end_x_root)
            y2 = max(start_y_root, end_y_root)
            
            if abs(x2 - x1) > 10 and abs(y2 - y1) > 10:
                # Close selection window
                selection_window.destroy()
                self.root.deiconify()  # Show main window
                # Capture selected region using absolute coordinates
                self.capture_region(x1, y1, x2, y2)
            else:
                messagebox.showinfo("Selection", "Please select a larger area")
        
        def cancel_selection(event):
            if event.keysym == 'Escape':
                selection_window.destroy()
                self.root.deiconify()
        
        canvas.bind("<ButtonPress-1>", on_button_press)
        canvas.bind("<B1-Motion>", on_move_press)
        canvas.bind("<ButtonRelease-1>", on_button_release)
        selection_window.bind("<KeyPress>", cancel_selection)
        selection_window.focus_set()  # Get focus for keyboard events
        
        # Instructions
        instruction_label = tk.Label(
            selection_window,
            text="Click and drag to select area, or press ESC to cancel",
            bg='black', fg='white', font=("Arial", 14)
        )
        instruction_label.place(relx=0.5, rely=0.05, anchor=tk.CENTER)
    
    def capture_region(self, x1, y1, x2, y2):
        """Capture the selected screen region"""
        try:
            with mss.mss() as sct:
                # Get monitor dimensions
                monitor = sct.monitors[0]
                
                # Calculate region
                region = {
                    "top": y1,
                    "left": x1,
                    "width": abs(x2 - x1),
                    "height": abs(y2 - y1)
                }
                
                # Capture screenshot
                screenshot = sct.grab(region)
                img = Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")
                
                self.captured_image = img
                self.display_preview(img)
                self.extract_btn.config(state=tk.NORMAL)
                self.status_var.set(f"Area captured: {img.width}x{img.height} pixels")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to capture screen: {str(e)}")
            self.status_var.set("Error capturing screen")
    
    def display_preview(self, image):
        """Display image preview in canvas"""
        # Resize if too large
        max_width = 750
        max_height = 400
        
        if image.width > max_width or image.height > max_height:
            ratio = min(max_width / image.width, max_height / image.height)
            new_width = int(image.width * ratio)
            new_height = int(image.height * ratio)
            image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
        
        # Convert to PhotoImage
        photo = ImageTk.PhotoImage(image)
        
        # Update canvas
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor=tk.NW, image=photo)
        self.canvas.image = photo  # Keep a reference
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def extract_data(self):
        """Extract tabular data from captured image"""
        if self.captured_image is None:
            messagebox.showwarning("No Image", "Please capture an area first")
            return
        
        self.progress.start()
        self.status_var.set("Extracting data... This may take a moment")
        self.extract_btn.config(state=tk.DISABLED)
        
        # Run extraction in separate thread to keep UI responsive
        thread = threading.Thread(target=self._extract_data_thread)
        thread.daemon = True
        thread.start()
    
    def _extract_data_thread(self):
        """Extract data in background thread"""
        try:
            # Preprocess image for better OCR
            img = self.preprocess_image(self.captured_image)
            
            # Use OCR to extract text
            ocr_config = r'--oem 3 --psm 6'  # Assume uniform block of text
            text = pytesseract.image_to_string(img, config=ocr_config)
            
            # Parse text into rows
            rows = self.parse_text_to_rows(text)
            
            # Convert to DataFrame
            if rows:
                self.processed_data = pd.DataFrame(rows)
                self.root.after(0, self._extraction_complete, True)
            else:
                self.root.after(0, self._extraction_complete, False)
        
        except Exception as e:
            self.root.after(0, self._extraction_complete_error, str(e))
    
    def preprocess_image(self, image):
        """Preprocess image to improve OCR accuracy"""
        # Convert to numpy array
        img_array = np.array(image)
        
        # Convert to grayscale
        if len(img_array.shape) == 3:
            gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
        else:
            gray = img_array
        
        # Apply thresholding
        _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        
        # Denoise
        denoised = cv2.fastNlMeansDenoising(thresh, None, 10, 7, 21)
        
        # Convert back to PIL Image
        processed = Image.fromarray(denoised)
        
        # Enhance contrast
        enhancer = ImageEnhance.Contrast(processed)
        processed = enhancer.enhance(2.0)
        
        return processed
    
    def parse_text_to_rows(self, text):
        """Parse OCR text into rows of data"""
        lines = text.strip().split('\n')
        rows = []
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # Try to split by common delimiters
            # First, try tab
            if '\t' in line:
                cells = [cell.strip() for cell in line.split('\t')]
            # Then try multiple spaces (at least 2)
            elif '  ' in line:
                cells = re.split(r'\s{2,}', line)
            # Then try pipe
            elif '|' in line:
                cells = [cell.strip() for cell in line.split('|')]
            # Then try comma
            elif ',' in line:
                cells = [cell.strip() for cell in line.split(',')]
            else:
                # Single column
                cells = [line]
            
            # Filter out empty cells
            cells = [cell for cell in cells if cell]
            
            if cells:
                rows.append(cells)
        
        # Normalize row lengths (pad shorter rows)
        if rows:
            max_cols = max(len(row) for row in rows)
            normalized_rows = []
            for row in rows:
                while len(row) < max_cols:
                    row.append('')
                normalized_rows.append(row)
            return normalized_rows
        
        return []
    
    def _extraction_complete(self, success):
        """Called when extraction completes"""
        self.progress.stop()
        self.extract_btn.config(state=tk.NORMAL)
        
        if success and self.processed_data is not None:
            self.export_btn.config(state=tk.NORMAL)
            row_count = len(self.processed_data)
            col_count = len(self.processed_data.columns)
            self.status_var.set(
                f"Extraction complete: {row_count} rows, {col_count} columns. Ready to export."
            )
            messagebox.showinfo("Success", f"Extracted {row_count} rows with {col_count} columns")
        else:
            self.status_var.set("Extraction failed: No data found. Try adjusting the selection.")
            messagebox.showwarning("No Data", "Could not extract data. Try selecting a different area.")
    
    def _extraction_complete_error(self, error_msg):
        """Called when extraction fails with error"""
        self.progress.stop()
        self.extract_btn.config(state=tk.NORMAL)
        self.status_var.set(f"Error: {error_msg}")
        messagebox.showerror("Extraction Error", f"Failed to extract data:\n{error_msg}")
    
    def export_csv(self):
        """Export extracted data to CSV file"""
        if self.processed_data is None:
            messagebox.showwarning("No Data", "Please extract data first")
            return
        
        # Get save location
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"extracted_data_{timestamp}.csv"
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile=default_filename
        )
        
        if filename:
            try:
                self.processed_data.to_csv(filename, index=False)
                self.status_var.set(f"Data exported to: {os.path.basename(filename)}")
                messagebox.showinfo("Success", f"Data exported successfully to:\n{filename}")
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export CSV:\n{str(e)}")
                self.status_var.set("Export failed")


def main():
    """Main entry point"""
    root = tk.Tk()
    app = ScreenScannerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

