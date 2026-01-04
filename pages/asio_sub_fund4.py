import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry
from datetime import datetime
import os
import json
import pandas as pd
from collections import defaultdict
import threading
import csv
import io
import zipfile
import re
import tempfile
import shutil

from my_app.pages.loading import LoadingSpinner
from my_app.pages.helper import output_save_in_template, multiple_excels_to_zip, is_missing

from my_app.CONSTANTS import (
    RECORDTYPE, TradeQuantity, sub_fund_4_headers, keep_blank_for_headers,
    RECORDACTION, USERTRANID1, PORTFOLIO, STRATEGY, PRICEDENOMINATION,
    COUNTERINVESTMENT, NETINVESTMENTAMOUNT, NETCOUNTERAMOUNT, FUNDSTRUCTURE,
    PRICEDIRECTLY, COUNTERFXDENOMINATION, QUANTITY, PRICE, EVENTDATE, SETTLEDATE, ACTUALSETTLEDATE
)
from CONSTANTS import BROKER, KEYVALUE, KEYVALUE_KEYNAME, LOCATIONACCOUNT, BuySellIndicator, TradingCode
from CONSTANTS import INVESTMENT

class ASIOSubFund4Page(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg="#ffffff")

        # Title (fixed at top)
        title_frame = tk.Frame(self, bg="#ffffff")
        title_frame.pack(fill="x", pady=(15, 0))
        
        title = tk.Label(
            title_frame, 
            text="📑 ASIO Sub Fund 4", 
            font=("Arial", 22, "bold"), 
            bg="#ffffff", 
            fg="#2c3e50"
        )
        title.pack(pady=(0, 20))

        # Create canvas for scrolling with visible scrollbar
        canvas_frame = tk.Frame(self, bg="#ffffff")
        canvas_frame.pack(fill="both", expand=True, padx=0, pady=0)
        
        # Canvas for scrolling
        self.canvas = tk.Canvas(
            canvas_frame,
            bg="#ffffff",
            highlightthickness=0,
            borderwidth=0
        )
        
        # Scrollbar for canvas
        self.scrollbar = tk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        
        # Configure canvas scrolling
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Pack canvas first, then scrollbar
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # Scrollable frame inside canvas
        self.scrollable_frame = tk.Frame(self.canvas, bg="#ffffff")
        
        # Configure canvas scrolling
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        # Create window in canvas for scrollable frame
        self.canvas_window = self.canvas.create_window(
            (0, 0),
            window=self.scrollable_frame,
            anchor="nw"
        )
        
        # Bind canvas resize to update scrollable frame width
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        
        # Bind mouse wheel for scrolling
        self._bind_mousewheel()
        
        # Bind keyboard events for scrolling with arrow keys (same as dashboard)
        self._bind_keyboard_scrolling()
        
        # Main container inside scrollable frame
        main_container = tk.Frame(self.scrollable_frame, bg="#ffffff")
        main_container.pack(fill="x", padx=40, pady=20)

        # ========== File Upload ==========
        # Section title with subtle styling
        file_title = tk.Label(
            main_container,
            text="📁 File Upload",
            font=("Arial", 14, "bold"),
            bg="#ffffff",
            fg="#2c3e50",
            anchor="w"
        )
        file_title.pack(fill="x", pady=(0, 15))

        # ListBox container
        listbox_container = tk.Frame(main_container, bg="#ffffff")
        listbox_container.pack(fill="x", pady=(0, 12))

        listbox_label = tk.Label(
            listbox_container,
            text="Selected Files:",
            font=("Arial", 11, "bold"),
            bg="#ffffff",
            fg="#2c3e50",
            anchor="w"
        )
        listbox_label.pack(fill="x", pady=(0, 8))

        # ListBox with scrollbar
        listbox_frame = tk.Frame(listbox_container, bg="#ecf0f1", relief="flat", bd=1)
        listbox_frame.pack(fill="x")

        scrollbar = tk.Scrollbar(listbox_frame, orient="vertical")
        scrollbar.pack(side="right", fill="y")

        self.file_listbox = tk.Listbox(
            listbox_frame,
            yscrollcommand=scrollbar.set,
            font=("Arial", 10),
            bg="white",
            fg="#2c3e50",
            selectmode=tk.SINGLE,
            height=5,
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground="#bdc3c7"
        )
        self.file_listbox.pack(side="left", fill="both", expand=True, padx=1, pady=1)
        scrollbar.config(command=self.file_listbox.yview)

        # Button group for file actions (Browse and Remove together)
        file_buttons_frame = tk.Frame(main_container, bg="#ffffff")
        file_buttons_frame.pack(fill="x", pady=(12, 0))

        browse_btn = tk.Button(
            file_buttons_frame,
            text="📁 Browse File",
            command=self._browse_files,
            bg="#3498db",
            fg="white",
            activebackground="#2980b9",
            activeforeground="white",
            relief="flat",
            padx=20,
            pady=8,
            font=("Arial", 11, "bold"),
            cursor="hand2"
        )
        browse_btn.pack(side="left", padx=(0, 10))

        remove_btn = tk.Button(
            file_buttons_frame,
            text="🗑️ Remove Selected",
            command=self._remove_selected_file,
            bg="#e74c3c",
            fg="white",
            activebackground="#c0392b",
            activeforeground="white",
            relief="flat",
            padx=20,
            pady=8,
            font=("Arial", 11, "bold"),
            cursor="hand2"
        )
        remove_btn.pack(side="left", padx=(0, 10))

        remove_all_btn = tk.Button(
            file_buttons_frame,
            text="🗑️ Remove All",
            command=self._remove_all_files,
            bg="#c0392b",
            fg="white",
            activebackground="#a93226",
            activeforeground="white",
            relief="flat",
            padx=20,
            pady=8,
            font=("Arial", 11, "bold"),
            cursor="hand2"
        )
        remove_all_btn.pack(side="left")

        # Store selected files
        self.selected_files = []
        
        # Store loader data for export
        self.loader_data = defaultdict(list)

        # Divider line
        divider1 = tk.Frame(main_container, bg="#e0e0e0", height=1)
        divider1.pack(fill="x", pady=(25, 25))

        # ========== Bulk Processing Checkbox (above Date Selection) ==========
        bulk_frame = tk.Frame(main_container, bg="#ffffff")
        bulk_frame.pack(fill="x", pady=(0, 20))
        
        self.bulk_var = tk.BooleanVar(value=False)
        bulk_checkbox = tk.Checkbutton(
            bulk_frame,
            text="Bulk Processing (ZIP files with automatic date extraction)",
            variable=self.bulk_var,
            font=("Arial", 11),
            bg="#ffffff",
            fg="#34495e",
            activebackground="#ffffff",
            activeforeground="#34495e",
            selectcolor="#ffffff",
            cursor="hand2"
        )
        bulk_checkbox.pack(side="left")

        # ========== Container for Date Selection and File Reading (side by side) ==========
        date_reading_container = tk.Frame(main_container, bg="#ffffff")
        date_reading_container.pack(fill="x", pady=(0, 25))

        # ========== Date Input (Left Side) ==========
        date_left_container = tk.Frame(date_reading_container, bg="#ffffff")
        date_left_container.pack(side="left", fill="both", expand=True, padx=(0, 30))
        
        date_title = tk.Label(
            date_left_container,
            text="📅 Date Selection",
            font=("Arial", 14, "bold"),
            bg="#ffffff",
            fg="#2c3e50",
            anchor="w"
        )
        date_title.pack(fill="x", pady=(0, 15))

        # Date fields row
        date_row = tk.Frame(date_left_container, bg="#ffffff")
        date_row.pack(fill="x", pady=(0, 0))

        # Event Date - create DateEntry widgets asynchronously for instant frame opening
        event_date_frame = tk.Frame(date_row, bg="#ffffff")
        event_date_frame.pack(side="left", padx=(0, 30))
        tk.Label(
            event_date_frame,
            text="Event Date",
            font=("Arial", 11, "bold"),
            bg="#ffffff",
            fg="#34495e"
        ).pack(side="top", pady=(0, 8))
        # Store frame reference for async widget creation
        self._event_date_frame = event_date_frame
        self.event_date_entry = None
        
        # Settlement Date
        settlement_date_frame = tk.Frame(date_row, bg="#ffffff")
        settlement_date_frame.pack(side="left", padx=(0, 30))
        tk.Label(
            settlement_date_frame,
            text="Settlement Date",
            font=("Arial", 11, "bold"),
            bg="#ffffff",
            fg="#34495e"
        ).pack(side="top", pady=(0, 8))
        self._settlement_date_frame = settlement_date_frame
        self.settlement_date_entry = None

        # Actual Date
        actual_date_frame = tk.Frame(date_row, bg="#ffffff")
        actual_date_frame.pack(side="left")
        tk.Label(
            actual_date_frame,
            text="Actual Date",
            font=("Arial", 11, "bold"),
            bg="#ffffff",
            fg="#34495e"
        ).pack(side="top", pady=(0, 8))
        self._actual_date_frame = actual_date_frame
        self.actual_date_entry = None
        
        # Create DateEntry widgets asynchronously (DateEntry creation is slow - this avoids blocking)
        self.after_idle(self._create_date_entries_async)

        # ========== File Reading (Right Side) ==========
        reading_right_container = tk.Frame(date_reading_container, bg="#ffffff")
        reading_right_container.pack(side="right", fill="both", expand=True)
        
        reading_title = tk.Label(
            reading_right_container,
            text="⚙️ File Reading Configuration",
            font=("Arial", 14, "bold"),
            bg="#ffffff",
            fg="#2c3e50",
            anchor="w"
        )
        reading_title.pack(fill="x", pady=(0, 15))

        # Number inputs row
        reading_row = tk.Frame(reading_right_container, bg="#ffffff")
        reading_row.pack(fill="x", pady=(0, 0))

        # Read From Row
        read_row_frame = tk.Frame(reading_row, bg="#ffffff")
        read_row_frame.pack(side="left", padx=(0, 40))
        tk.Label(
            read_row_frame,
            text="Read From Row",
            font=("Arial", 11, "bold"),
            bg="#ffffff",
            fg="#34495e"
        ).pack(side="top", pady=(0, 8))
        # Load saved read configuration (lazy load - use defaults initially)
        # Load config in background to avoid blocking UI
        self._read_config_cache = None
        default_read_row = "1"
        default_read_col = "A"
        
        self.read_row_var = tk.StringVar(value=default_read_row)
        read_row_entry = tk.Entry(
            read_row_frame,
            textvariable=self.read_row_var,
            width=15,
            font=("Arial", 11),
            justify="center",
            relief="solid",
            bd=1,
            highlightthickness=1,
            highlightbackground="#bdc3c7",
            highlightcolor="#3498db"
        )
        read_row_entry.pack(side="top")

        # Read From Column (A, B, C format)
        read_col_frame = tk.Frame(reading_row, bg="#ffffff")
        read_col_frame.pack(side="left")
        tk.Label(
            read_col_frame,
            text="Read From Column",
            font=("Arial", 11, "bold"),
            bg="#ffffff",
            fg="#34495e"
        ).pack(side="top", pady=(0, 8))
        self.read_col_var = tk.StringVar(value=default_read_col)
        read_col_entry = tk.Entry(
            read_col_frame,
            textvariable=self.read_col_var,
            width=15,
            font=("Arial", 11, "bold"),
            justify="center",
            relief="solid",
            bd=1,
            highlightthickness=1,
            highlightbackground="#bdc3c7",
            highlightcolor="#3498db"
        )
        read_col_entry.pack(side="top")
        
        # Bind to uppercase and validate column letters
        def on_col_change(*args):
            current = self.read_col_var.get()
            # Convert to uppercase
            upper = current.upper()
            if upper != current:
                self.read_col_var.set(upper)
        
        self.read_col_var.trace_add("write", on_col_change)

        # Validation for inputs
        read_row_entry.config(validate="key", validatecommand=(self.register(self._validate_number), "%P"))
        read_col_entry.config(validate="key", validatecommand=(self.register(self._validate_column_letter), "%P"))

        # ========== Container for Output Format and File Reading Fallback (side by side) ==========
        output_fallback_container = tk.Frame(main_container, bg="#ffffff")
        output_fallback_container.pack(fill="x", pady=(5, 0))

        # Output Format checkboxes (Left Side)
        output_format_frame = tk.Frame(output_fallback_container, bg="#ffffff")
        output_format_frame.pack(side="left", fill="x", expand=True)
        
        tk.Label(
            output_format_frame,
            text="Output Format:",
            font=("Arial", 11, "bold"),
            bg="#ffffff",
            fg="#34495e"
        ).pack(side="left", padx=(0, 15))
        
        self.export_excel_var = tk.BooleanVar(value=False)
        excel_checkbox = tk.Checkbutton(
            output_format_frame,
            text="Excel (.xlsx)",
            variable=self.export_excel_var,
            font=("Arial", 11),
            bg="#ffffff",
            fg="#34495e",
            activebackground="#ffffff",
            activeforeground="#34495e",
            selectcolor="#ffffff",
            cursor="hand2"
        )
        excel_checkbox.pack(side="left", padx=(0, 15))
        
        self.export_csv_var = tk.BooleanVar(value=True)
        csv_checkbox = tk.Checkbutton(
            output_format_frame,
            text="CSV (.csv)",
            variable=self.export_csv_var,
            font=("Arial", 11),
            bg="#ffffff",
            fg="#34495e",
            activebackground="#ffffff",
            activeforeground="#34495e",
            selectcolor="#ffffff",
            cursor="hand2"
        )
        csv_checkbox.pack(side="left")

        # File Reading Fallback checkbox (Right Side, aligned with Read From Row)
        # Align with Read From Row by matching the right container structure
        # Since reading_right_container is packed side="right" and read_row_frame starts at its left edge,
        # we need to match that alignment. Use a spacer frame that matches reading_right_container width behavior
        fallback_spacer = tk.Frame(output_fallback_container, bg="#ffffff")
        fallback_spacer.pack(side="right", fill="both", expand=True)
        
        fallback_frame = tk.Frame(fallback_spacer, bg="#ffffff")
        fallback_frame.pack(side="left", padx=(170, 0))
        
        self.fallback_var = tk.BooleanVar(value=False)
        fallback_checkbox = tk.Checkbutton(
            fallback_frame,
            text="File Reading Fallback (Row 10, Column B)",
            variable=self.fallback_var,
            font=("Arial", 11),
            bg="#ffffff",
            fg="#34495e",
            activebackground="#ffffff",
            activeforeground="#34495e",
            selectcolor="#ffffff",
            cursor="hand2"
        )
        fallback_checkbox.pack(side="left")

        # Divider line before submit
        divider3 = tk.Frame(main_container, bg="#e0e0e0", height=1)
        divider3.pack(fill="x", pady=(0, 25))

        # ========== Submit Button Section ==========
        submit_section = tk.Frame(main_container, bg="#ffffff")
        submit_section.pack(fill="x", pady=(0, 15))

        # Submit button with hover effect - centered and prominent
        submit_btn = tk.Button(
            submit_section,
            text="Submit",
            command=self._submit,
            bg="#27ae60",
            fg="white",
            activebackground="#229954",
            activeforeground="white",
            relief="flat",
            padx=50,
            pady=14,
            font=("Arial", 14, "bold"),
            cursor="hand2"
        )
        submit_btn.pack()
        
        # Add hover effect
        def on_enter(e):
            submit_btn.config(bg="#229954")
        
        def on_leave(e):
            submit_btn.config(bg="#27ae60")
        
        submit_btn.bind("<Enter>", on_enter)
        submit_btn.bind("<Leave>", on_leave)

        # Status bar
        self.status_var = tk.StringVar(value="Ready - Please select files and configure settings")
        self.status_label = tk.Label(
            main_container,
            textvariable=self.status_var,
            font=("Arial", 10),
            bg="#f8f9fa",
            fg="#6c757d",
            anchor="w",
            relief="flat",
            bd=0,
            padx=15,
            pady=10
        )
        self.status_label.pack(fill="x", pady=(0, 20))
    
    def _on_canvas_configure(self, event):
        """Update scrollable frame width when canvas is resized."""
        canvas_width = event.width
        self.canvas.itemconfig(self.canvas_window, width=canvas_width)
    
    def _bind_mousewheel(self):
        """Bind mouse wheel events for scrolling."""
        def _on_mousewheel(event):
            try:
                # Check if canvas still exists before scrolling
                if hasattr(self, 'canvas') and self.canvas.winfo_exists():
                    self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            except Exception:
                pass  # Ignore errors if canvas is destroyed
        
        def _bind_to_mousewheel(event):
            try:
                if hasattr(self, 'canvas') and self.canvas.winfo_exists():
                    self.canvas.bind_all("<MouseWheel>", _on_mousewheel)
            except Exception:
                pass
        
        def _unbind_from_mousewheel(event):
            try:
                if hasattr(self, 'canvas') and self.canvas.winfo_exists():
                    self.canvas.unbind_all("<MouseWheel>")
            except Exception:
                pass
        
        # Bind mousewheel when mouse enters canvas (with safety check)
        try:
            if hasattr(self, 'canvas') and self.canvas.winfo_exists():
                self.canvas.bind("<Enter>", _bind_to_mousewheel)
                self.canvas.bind("<Leave>", _unbind_from_mousewheel)
        except Exception:
            pass
    
    def _bind_keyboard_scrolling(self):
        """Bind keyboard events for scrolling with arrow keys (same as dashboard)"""
        # Bind arrow keys to canvas and scrollable frame
        self.canvas.bind("<Up>", lambda e: self._scroll_up())
        self.canvas.bind("<Down>", lambda e: self._scroll_down())
        self.canvas.bind("<Page_Up>", lambda e: self._scroll_page_up())
        self.canvas.bind("<Page_Down>", lambda e: self._scroll_page_down())
        self.canvas.bind("<Home>", lambda e: self._scroll_home())
        self.canvas.bind("<End>", lambda e: self._scroll_end())
        
        # Also bind to scrollable frame
        self.scrollable_frame.bind("<Up>", lambda e: self._scroll_up())
        self.scrollable_frame.bind("<Down>", lambda e: self._scroll_down())
        self.scrollable_frame.bind("<Page_Up>", lambda e: self._scroll_page_up())
        self.scrollable_frame.bind("<Page_Down>", lambda e: self._scroll_page_down())
        self.scrollable_frame.bind("<Home>", lambda e: self._scroll_home())
        self.scrollable_frame.bind("<End>", lambda e: self._scroll_end())
        
        # Set focus to canvas so keyboard events work
        self.canvas.focus_set()
        
        # Also bind when clicking on canvas or scrollable frame
        self.canvas.bind("<Button-1>", lambda e: self.canvas.focus_set())
        self.scrollable_frame.bind("<Button-1>", lambda e: self.canvas.focus_set())
    
    def _scroll_up(self):
        """Scroll up one unit"""
        self.canvas.yview_scroll(-1, "units")
        return "break"
    
    def _scroll_down(self):
        """Scroll down one unit"""
        self.canvas.yview_scroll(1, "units")
        return "break"
    
    def _scroll_page_up(self):
        """Scroll up one page"""
        self.canvas.yview_scroll(-1, "pages")
        return "break"
    
    def _scroll_page_down(self):
        """Scroll down one page"""
        self.canvas.yview_scroll(1, "pages")
        return "break"
    
    def _scroll_home(self):
        """Scroll to top"""
        self.canvas.yview_moveto(0)
        return "break"
    
    def _scroll_end(self):
        """Scroll to bottom"""
        self.canvas.yview_moveto(1)
        return "break"

    def _browse_files(self):
        """Browse and select xls, xlsx, csv files, or zip files if bulk mode is enabled."""
        if self.bulk_var.get():
            # Bulk mode: allow all files (user can select zip files)
            files = filedialog.askopenfilenames(
                title="Select ZIP Files",
                filetypes=[
                    ("All Files", "*.*"),
                    ("ZIP Files", "*.zip")
                ]
            )
        else:
            # Normal mode: allow Excel and CSV files
            files = filedialog.askopenfilenames(
                title="Select Files",
                filetypes=[
                    # ("All Supported Files", "*.xls *.xlsx *.csv"),
                    # ("Excel Files", "*.xls *.xlsx"),
                    # ("CSV Files", "*.csv"),
                    ("All Files", "*.*")
                ]
            )

        if files:
            for file_path in files:
                if file_path not in self.selected_files:
                    self.selected_files.append(file_path)
                    # Display only filename in ListBox
                    filename = os.path.basename(file_path)
                    self.file_listbox.insert(tk.END, filename)
            
            self.status_var.set(f"{len(self.selected_files)} file(s) selected")

    def _remove_selected_file(self):
        """Remove selected file from list."""
        selection = self.file_listbox.curselection()
        if selection:
            index = selection[0]
            self.file_listbox.delete(index)
            self.selected_files.pop(index)
            self.status_var.set(f"{len(self.selected_files)} file(s) remaining")
        else:
            messagebox.showinfo("Info", "Please select a file to remove")

    def _remove_all_files(self):
        """Remove all files from list."""
        if not self.selected_files:
            messagebox.showinfo("Info", "No files to remove")
            return
        
        # Clear the listbox
        self.file_listbox.delete(0, tk.END)
        
        # Clear the selected files list
        self.selected_files.clear()
        
        self.status_var.set("All files removed")

    def _validate_number(self, value):
        """Validate that input is a positive integer."""
        if value == "":
            return True
        try:
            int(value)
            return True
        except ValueError:
            return False
    
    def _validate_column_letter(self, value):
        """Validate that input is a valid column letter (A-Z, AA-ZZ, etc.)."""
        if value == "":
            return True
        # Convert to uppercase for validation
        upper_value = value.upper()
        # Allow only letters A-Z (single or double letter columns like AA, AB, etc.)
        if len(upper_value) <= 2 and upper_value.isalpha():
            return True
        return False
    
    def _on_event_date_change(self, event=None):
        """Update Settlement Date and Actual Date to match Event Date when Event Date changes.
        This only works one-way: Event Date -> Settlement Date and Actual Date.
        Changes to Settlement Date or Actual Date do not affect Event Date.
        """
        try:
            # Check if widgets are created
            if not (self.event_date_entry and self.settlement_date_entry and self.actual_date_entry):
                return
            # Get the selected date from Event Date
            event_date = self.event_date_entry.get_date()
            
            # Only update if we have a valid date
            if event_date:
                # Update Settlement Date and Actual Date to match Event Date
                self.settlement_date_entry.set_date(event_date)
                self.actual_date_entry.set_date(event_date)
        except Exception:
            # Silently ignore errors (e.g., if dates are not yet initialized)
            pass
    
    def _create_date_entries_async(self):
        """Create DateEntry widgets asynchronously after frame is shown - this avoids blocking initialization."""
        try:
            # Event Date
            self.event_date_entry = DateEntry(
                self._event_date_frame,
                width=18,
                background="#3498db",
                foreground="white",
                borderwidth=2,
                date_pattern='dd/MM/yyyy',
                font=("Arial", 10)
            )
            self.event_date_entry.pack(side="top")
            
            # Settlement Date
            self.settlement_date_entry = DateEntry(
                self._settlement_date_frame,
                width=18,
                background="#3498db",
                foreground="white",
                borderwidth=2,
                date_pattern='dd/MM/yyyy',
                font=("Arial", 10)
            )
            self.settlement_date_entry.pack(side="top")
            
            # Actual Date
            self.actual_date_entry = DateEntry(
                self._actual_date_frame,
                width=18,
                background="#3498db",
                foreground="white",
                borderwidth=2,
                date_pattern='dd/MM/yyyy',
                font=("Arial", 10)
            )
            self.actual_date_entry.pack(side="top")
            
            # Now set up bindings
            self._setup_date_entry_bindings()
            self.after(50, self._setup_calendar_binding)
            self.after(100, self._lazy_load_read_config)
        except Exception:
            pass
    
    def _setup_date_entry_bindings(self):
        """Set up date entry bindings (non-blocking)."""
        try:
            if not self.event_date_entry:
                return
            # Hook into the calendar selection by overriding the _select_calendar method
            # Store original method if it exists
            original_select = getattr(self.event_date_entry, '_select_calendar', None)
            
            def wrapped_select(sel_date):
                # Call original method if it exists
                if original_select:
                    original_select(sel_date)
                # Update other dates
                self._on_event_date_change()
            
            # Replace the method
            self.event_date_entry._select_calendar = wrapped_select
            
            # Also bind to focus out event as backup (when user types date manually)
            self.event_date_entry.bind("<FocusOut>", lambda e: self._on_event_date_change())
        except Exception:
            pass
    
    def _setup_calendar_binding(self):
        """Set up binding to the calendar widget after it's created."""
        try:
            # Access the calendar widget if it exists
            if hasattr(self.event_date_entry, '_top_cal') and self.event_date_entry._top_cal:
                # Bind to calendar date selection
                self.event_date_entry._top_cal.bind("<<CalendarSelected>>", 
                                                    lambda e: self._on_event_date_change())
        except Exception:
            pass
    
    def _load_read_config(self):
        """Load read configuration from consolidated_data.json (cached)."""
        # Return cached config if available
        if self._read_config_cache is not None:
            return self._read_config_cache
        
        try:
            from my_app.file_utils import get_app_directory
            app_dir = get_app_directory()
            consolidated_path = os.path.join(app_dir, "consolidated_data.json")
            
            if os.path.exists(consolidated_path):
                with open(consolidated_path, "r", encoding="utf-8") as f:
                    consolidated_data = json.load(f)
                    read_config = consolidated_data.get("asio_sub_fund4_read_config", {})
                    # Cache the config
                    self._read_config_cache = read_config
                    return read_config
        except Exception:
            pass
        
        # Return default if loading fails
        default_config = {"read_from_row": 1, "read_from_column": "A"}
        self._read_config_cache = default_config
        return default_config
    
    def _lazy_load_read_config(self):
        """Lazy load read config and update UI fields if different from defaults."""
        try:
            read_config = self._load_read_config()
            saved_row = str(read_config.get("read_from_row", 1))
            saved_col = read_config.get("read_from_column", "A")
            
            # Only update if current values are defaults and saved values differ
            if self.read_row_var.get() == "1" and saved_row != "1":
                self.read_row_var.set(saved_row)
            if self.read_col_var.get() == "A" and saved_col != "A":
                self.read_col_var.set(saved_col)
        except Exception:
            pass  # Silently fail - defaults are already set
    
    def _save_read_config(self, read_row, read_col_letter):
        """Save read configuration to consolidated_data.json."""
        try:
            from my_app.file_utils import get_app_directory
            app_dir = get_app_directory()
            consolidated_path = os.path.join(app_dir, "consolidated_data.json")
            
            # Load existing data
            consolidated_data = {}
            if os.path.exists(consolidated_path):
                with open(consolidated_path, "r", encoding="utf-8") as f:
                    consolidated_data = json.load(f)
            
            # Update read config
            consolidated_data["asio_sub_fund4_read_config"] = {
                "read_from_row": read_row,
                "read_from_column": read_col_letter
            }
            
            # Save back to file
            with open(consolidated_path, "w", encoding="utf-8") as f:
                json.dump(consolidated_data, f, indent=4)
        except Exception as e:
            # Silently fail - don't interrupt user workflow
            pass
    
    def _column_letter_to_index(self, column_letter):
        """Convert Excel column letter (A, B, C, ..., Z, AA, AB, etc.) to 0-based index.
        
        Args:
            column_letter: Column letter string (e.g., "A", "B", "AA")
        
        Returns:
            int: 0-based column index (A=0, B=1, Z=25, AA=26, etc.)
        """
        if not column_letter:
            return 0
        
        column_letter = column_letter.upper().strip()
        result = 0
        for char in column_letter:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result - 1  # Convert to 0-based index
    
    def _extract_date_from_zip_filename(self, zip_filename):
        """Extract date from zip filename in DDMMYYYY format.
        
        Args:
            zip_filename: Zip filename (e.g., "FT TRADES  _02122025.zip" or "Fw_ FT TRADES  _01122025.zip")
        
        Returns:
            datetime: Parsed date object, or None if extraction fails
        """
        try:
            # Look for 8-digit date pattern (DDMMYYYY) in filename
            # Pattern: 8 consecutive digits
            match = re.search(r'(\d{8})', zip_filename)
            if match:
                date_str = match.group(1)
                # Parse as DDMMYYYY
                day = int(date_str[0:2])
                month = int(date_str[2:4])
                year = int(date_str[4:8])
                return datetime(year, month, day)
        except Exception:
            pass
        return None
    
    def _extract_files_from_zip(self, zip_path, temp_dir):
        """Extract Excel and CSV files from zip archive.
        
        Args:
            zip_path: Path to zip file
            temp_dir: Temporary directory to extract files to
        
        Returns:
            list: List of extracted file paths [(file_path, is_excel), ...]
        """
        extracted_files = []
        try:
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                # Extract all files to temp directory
                zip_ref.extractall(temp_dir)
                
                # Get list of extracted files
                for member in zip_ref.namelist():
                    # Skip directories
                    if member.endswith('/'):
                        continue
                    
                    extracted_path = os.path.join(temp_dir, member)
                    if os.path.exists(extracted_path):
                        # Check file extension
                        ext = os.path.splitext(extracted_path)[1].lower()
                        if ext in ['.xls', '.xlsx']:
                            extracted_files.append((extracted_path, True))  # is_excel = True
                        elif ext == '.csv':
                            extracted_files.append((extracted_path, False))  # is_excel = False
        except Exception as e:
            raise Exception(f"Failed to extract zip file {os.path.basename(zip_path)}: {e}")
        
        return extracted_files
    
    def read_dynamic_file(
            self,
            file_path: str,
            header_row: int = 1,
            header_start_col: int = 0,
            sheet_name: int | str | None = 0,
            **kwargs
        ):
        """
        Dynamic function to read CSV, XLS, and XLSX files with configurable header row and header start column.
        
        Args:
            file_path (str): Path to the file (CSV, XLS, or XLSX)
            header_row (int): Row number where header is located (1-based, Excel row number).
                            For example, header_row=10 means header is at row 10 in Excel.
            header_start_col (int): Column index where headers start (0-based, 0=A, 1=B, 2=C, etc.).
                                  For example, header_start_col=1 means headers start from column B.
            sheet_name (int | str | None): Sheet name or index for Excel files (default: 0 for first sheet)
            **kwargs: Additional pandas read_* options (usecols, dtype, parse_dates, etc.)
        
        Returns:
            pd.DataFrame: DataFrame with headers from specified row and column
        """
        from pathlib import Path
        import pandas as pd
        
        file_path = Path(file_path)
        
        if not file_path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")
        
        # Get file extension
        ext = file_path.suffix.lower()
        
        # Convert header_row (1-based) to 0-based index for pandas
        # If header_row=10 (1-based), then header_row_index=9 (0-based)
        header_row_index = header_row - 1
        
        # Skip rows before the header row
        skiprows = list(range(header_row_index)) if header_row_index > 0 else None
        
        # For CSV, XLS, XLSX - handle column selection
        if ext == ".csv":
            # Read CSV with header at specified row and columns from header_start_col onwards
            df = pd.read_csv(
                file_path,
                header=0,  # After skipping, header is at row 0
                skiprows=skiprows,
                usecols=range(header_start_col, 10_000) if header_start_col > 0 else None,
                **kwargs
            )
            
        elif ext in [".xls", ".xlsx"]:
            # Determine engine for Excel files
            engine = "openpyxl" if ext == ".xlsx" else "xlrd"
            
            # Read Excel file with skiprows and header
            # After skipping rows, the header row becomes row 0
            df = pd.read_excel(
                file_path,
                sheet_name=sheet_name,
                header=0,  # After skipping, header is at row 0
                skiprows=skiprows,
                engine=engine,
                **kwargs
            )
            
            # If header_start_col > 0, select columns starting from that index
            if header_start_col > 0:
                # Select columns from header_start_col onwards (drop columns before it)
                df = df.iloc[:, header_start_col:]
            
        else:
            raise ValueError(f"Unsupported file format: {ext}. Supported formats: CSV, XLS, XLSX")
        
        # Find the first completely blank row and stop reading there
        # This handles cases where there are blank rows followed by "Total" rows
        first_blank_row_idx = None
        for idx, row in df.iterrows():
            # Check if all values in the row are blank/NaN
            is_blank = True
            for val in row:
                if pd.notna(val) and str(val).strip() != "":
                    is_blank = False
                    break
            
            # When we encounter a blank row, stop reading (don't include this row or any after it)
            if is_blank:
                first_blank_row_idx = idx
                break
        
        # If we found a blank row, keep only rows BEFORE it (stop reading at that point)
        if first_blank_row_idx is not None:
            df = df.iloc[:first_blank_row_idx]
        
        # Clean up: remove any completely empty rows that might have been missed
        df = df.dropna(how='all')
        
        # Reset index after all operations
        df = df.reset_index(drop=True)
        
        return df

    def _get_location_account_from_trading_code(self, trading_code_key):
        """
        Generate location account value based on trading code key.
        
        Args:
            trading_code_key (str): Trading code key (e.g., "FT", "FT1", "FT2", "FT3", etc.)
        
        Returns:
            str: Location account value
        """
        base_location = "Asio_Sub Fund_4_OHM_FO_DBSBK0000289"
        
        # For all keys including FT, append the key as suffix
        return f"{base_location}_{trading_code_key}"
    
    def concatenate_security_name(self, row):
        try:
            symbol = row["Symbol"]
        except Exception as e:
            raise Exception(f"Error accessing Symbol: {e}")
        
        # Handle NaT (Not a Time) values in ExpiryDate
        try:
            value = row.get("ExpiryDate")

            if pd.isna(row["ExpiryDate"]):
                expiry = ""
            else:
                # If already a string → normalize/return safely
                if isinstance(value, str):
                    # try to parse into datetime first
                    try:
                        # value = pd.to_datetime(value)
                        value = pd.to_datetime(value, dayfirst=True, errors="coerce")
                        expiry = value.strftime("%Y%m%d") if value is not pd.NaT else ""
                    except:
                        # If string cannot be parsed, return original string (or "")
                        expiry = value.replace("-", "").replace("/", "")
                else:
                    # It is a datetime → directly format
                    expiry = value.strftime("%Y%m%d")
                    expiry = row["ExpiryDate"].strftime("%Y%m%d")  # format date
        except Exception as e:
            print(e)
            raise Exception(f"Error processing ExpiryDate: {e}, type: {type(row.get('ExpiryDate', 'N/A'))}, value: {row.get('ExpiryDate', 'N/A')}")
        
        # Handle NaT/NaN values in OptionType
        try:
            option_type = ""
            if pd.notna(row["OptionType"]):
                option_type = str(row["OptionType"])[0] if str(row["OptionType"]) else ""  # first character
        except Exception as e:
            raise Exception(f"Error processing OptionType: {e}")
        
        # Handle NaT/NaN values in StrikePrice
        try:
            strike = ""
            if pd.notna(row["StrikePrice"]):
                strike = str(int(row["StrikePrice"]))
        except Exception as e:
            raise Exception(f"Error processing StrikePrice: {e}")

        try:
            instrument = str(row.get("Instrument")).strip()
            if strike == "0" and option_type.strip() == "":
                option_type = instrument[0] # F
                # this when instrument is balnk or none or any other value which is not incorrect
                # value_is_missing = is_missing(instrument)
                # if value_is_missing:
                #     option_type = "F"
        except Exception as e:
            raise Exception(f"Error processing Instrument: {e}")

        return f"NSE{symbol}{expiry}{option_type}{strike}"

    def _prepare_data_row(self, row, config, trading_code, trading_code_mapping=None, columns=None, event_date=None, settlement_date=None, actual_date=None):
        """
        Prepare a single data row based on sub_fund_4_headers list order.
        
        Args:
            row: DataFrame row
            config: Configuration dictionary (e.g., asio_sf4_ft_config)
            trading_code: Trading code string (e.g., "FT", "FT1", "FT2")
            trading_code_mapping: Dictionary mapping trading codes to location accounts (optional)
            columns: DataFrame columns list to check for column existence (optional)
            event_date: Event date (datetime object, optional)
            settlement_date: Settlement date (datetime object, optional)
            actual_date: Actual date (datetime object, optional)
        
        Returns:
            list: Prepared row data in sub_fund_4_headers order
        """
        prepared_row = []
        try:
            security_name = self.concatenate_security_name(row)
        except Exception as e:
            raise Exception(f"Error in concatenate_security_name: {e}")
        
        # Get location account based on trading code
        location_account = None
        if trading_code_mapping and trading_code in trading_code_mapping:
            # Use value from consolidated_data.json mapping
            location_account = trading_code_mapping[trading_code]
        else:
            # Generate using function if not in mapping
            location_account = self._get_location_account_from_trading_code(trading_code)
        
        # Format dates if provided (format: YYYY-MM-DD)
        event_date_str = ""
        settlement_date_str = ""
        actual_date_str = ""
        
        if event_date:
            event_date_str = event_date.strftime("%d-%m-%Y")
        if settlement_date:
            settlement_date_str = settlement_date.strftime("%d-%m-%Y")
        if actual_date:
            actual_date_str = actual_date.strftime("%d-%m-%Y")
        # Process each header in sub_fund_4_headers order
        for header in sub_fund_4_headers:
            # Check if this header should be kept blank
            if header in keep_blank_for_headers:
                prepared_row.append("")
                continue
            
            # Special handling for RecordType
            if header == RECORDTYPE:
                # if security_name == "NSENIFTY20251230n0":
                #     breakpoint()
                if columns and BuySellIndicator in columns:
                    buy_sell_value = row.get(BuySellIndicator, "")
                    if buy_sell_value is not None:
                        prepared_row.append(buy_sell_value)
                    else:
                        prepared_row.append(config.get("RecordType", ""))
                else:
                    prepared_row.append(config.get("RecordType", ""))
                continue
            
            # Special handling for RecordAction
            if header == RECORDACTION:
                prepared_row.append(config.get(RECORDACTION, ""))
                continue
            
            # Special handling for KeyValue
            if header == KEYVALUE:
                prepared_row.append(security_name)
                continue
            
            # Special handling for KeyValue.KeyName
            if header == KEYVALUE_KEYNAME:
                prepared_row.append(config.get(KEYVALUE_KEYNAME, ""))
                continue
            
            # Special handling for UserTranId1
            if header == USERTRANID1:
                prepared_row.append(config.get(USERTRANID1, ""))
                continue
            
            # Special handling for Portfolio
            if header == PORTFOLIO:
                prepared_row.append(config.get(PORTFOLIO, ""))
                continue
            
            # Special handling for LocationAccount
            if header == LOCATIONACCOUNT:
                prepared_row.append(location_account)
                continue
            
            # Special handling for Strategy
            if header == STRATEGY:
                prepared_row.append(config.get(STRATEGY, ""))
                continue
            
            # Special handling for Investment
            if header == INVESTMENT:
                prepared_row.append(security_name)
                continue
            
            if header == PRICEDENOMINATION:
                prepared_row.append(config.get(PRICEDENOMINATION, ""))
                continue

            # Special handling for EventDate
            if header == EVENTDATE:
                prepared_row.append(event_date_str)
                continue
            
            # Special handling for SettleDate
            if header == SETTLEDATE:
                prepared_row.append(settlement_date_str)
                continue
            
            # Special handling for ActualSettleDate
            if header == ACTUALSETTLEDATE:
                prepared_row.append(actual_date_str)
                continue
            
            # Special handling for Quantity
            if header == QUANTITY:
                if columns and TradeQuantity in columns:
                    prepared_row.append(row.get(TradeQuantity, ""))
                continue
            
            # Special handling for Price
            if header == PRICE:
                if columns and PRICE in columns:
                    prepared_row.append(row.get(PRICE, ""))
                continue
            
            # Special handling for PriceDenomination
            if header == PRICEDENOMINATION:
                prepared_row.append(config.get(PRICEDENOMINATION, ""))
                continue
            
            # Special handling for CounterInvestment
            if header == COUNTERINVESTMENT:
                prepared_row.append(config.get(COUNTERINVESTMENT, ""))
                continue
            
            # Special handling for NetInvestmentAmount
            if header == NETINVESTMENTAMOUNT:
                prepared_row.append(config.get(NETINVESTMENTAMOUNT, ""))
                continue
            
            # Special handling for NetCounterAmount
            if header == NETCOUNTERAMOUNT:
                prepared_row.append(config.get(NETCOUNTERAMOUNT, ""))
                continue
            
            # Special handling for FundStructure
            if header == FUNDSTRUCTURE:
                prepared_row.append(config.get(FUNDSTRUCTURE, ""))
                continue
            
            # Special handling for CounterFXDenomination
            if header == COUNTERFXDENOMINATION:
                prepared_row.append(config.get(COUNTERFXDENOMINATION, ""))
                continue
        
        return prepared_row
    
    def _submit(self):
        """Handle form submission."""
        # Validate files
        if not self.selected_files:
            messagebox.showwarning("Warning", "Please select at least one file to process.")
            self.status_var.set("Error: No files selected")
            self.status_label.config(fg="#dc3545")  # Red color for errors
            return

        # Process the submission
        self.status_var.set("Processing... Please wait")
        self.status_label.config(fg="#6c757d")  # Reset to default gray color
        
        # Load configurations from consolidated_data.json (once for all files)
        from my_app.file_utils import get_app_directory
        app_dir = get_app_directory()
        consolidated_path = os.path.join(app_dir, "consolidated_data.json")
        
        try:
            if os.path.exists(consolidated_path):
                with open(consolidated_path, "r", encoding="utf-8") as f:
                    consolidated_data = json.load(f)
            else:
                raise FileNotFoundError("consolidated_data.json not found")
            
            # Extract ASIO Sub Fund 4 configuration (single config for all trading codes)
            asio_sf4_ft_config = consolidated_data.get("asio_sf4_ft", {})
            
            # Extract trading code mapping
            trading_code_mapping = consolidated_data.get("asio_sf4_trading_code_mapping", {})
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load consolidated_data.json: {e}")
            self.status_var.set(f"Error: Failed to load configurations")
            self.status_label.config(fg="#dc3545")  # Red color for errors
            return
        
        # Process all files in the list
        total_rows_processed = 0
        total_files_processed = 0
        temp_dirs = []  # Track temp directories for cleanup
        bulk_generated_zips = []  # Track all generated ZIP files for bulk mode
        
        try:
            # Check if bulk processing mode is enabled
            if self.bulk_var.get():
                # Ask for output directory once at the start
                self.bulk_export_dir = filedialog.askdirectory(
                    title="Select Directory to Save All Output ZIP Files"
                )
                if not self.bulk_export_dir:
                    self.status_var.set("Bulk processing cancelled by user")
                    return
                
                # Bulk processing mode: process zip files separately
                for zip_index, zip_path in enumerate(self.selected_files):
                    zip_filename = os.path.basename(zip_path)
                    self.status_var.set(f"Processing ZIP {zip_index + 1}/{len(self.selected_files)}: {zip_filename}")
                    
                    # Extract date from zip filename
                    extracted_date = self._extract_date_from_zip_filename(zip_filename)
                    if not extracted_date:
                        raise Exception(f"Could not extract date from zip filename: {zip_filename}")
                    
                    # Use extracted date for all three date fields
                    event_date = extracted_date
                    settlement_date = extracted_date
                    actual_date = extracted_date
                    
                    # Format date string for output filename
                    event_date_str = extracted_date.strftime("%d%m%Y")
                    
                    # Create temporary directory for this zip
                    temp_dir = tempfile.mkdtemp()
                    temp_dirs.append(temp_dir)
                    
                    # Extract files from zip
                    extracted_files = self._extract_files_from_zip(zip_path, temp_dir)
                    
                    if not extracted_files:
                        raise Exception(f"No Excel or CSV files found in zip: {zip_filename}")
                    
                    # Process data for this zip only (separate loader_data per zip)
                    zip_loader_data = defaultdict(list)
                    
                    # Process each extracted file
                    for file_path, is_excel in extracted_files:
                        self.status_var.set(f"Processing {os.path.basename(file_path)} from {zip_filename}")
                        
                        # Set read parameters based on file type
                        if is_excel:
                            # Excel: Row 10, Column B
                            read_row = 10
                            read_col = self._column_letter_to_index("B")
                        else:
                            # CSV: Row 1, Column A
                            read_row = 1
                            read_col = self._column_letter_to_index("A")
                        
                        # Read file
                        df = self.read_dynamic_file(
                            file_path=file_path,
                            header_row=read_row,
                            header_start_col=read_col,
                            sheet_name=0
                        )
                        
                        # Process each row
                        for index, row in df.iterrows():
                            try:
                                trading_code = str(row.get(TradingCode, "")).strip()
                            except Exception as e:
                                raise Exception(f"Error getting TradingCode at row {index} in file {os.path.basename(file_path)}: {e}")
                            
                            try:
                                data_row = self._prepare_data_row(
                                    row, 
                                    asio_sf4_ft_config, 
                                    trading_code, 
                                    trading_code_mapping, 
                                    columns=df.columns.tolist(),
                                    event_date=event_date,
                                    settlement_date=settlement_date,
                                    actual_date=actual_date
                                )
                            except Exception as e:
                                raise Exception(f"Error in _prepare_data_row at row {index} in file {os.path.basename(file_path)}: {e}")

                            zip_loader_data[trading_code].append(data_row)
                            total_rows_processed += 1
                        
                        total_files_processed += 1
                    
                    # Create output files for this zip and save as separate ZIP file
                    generated_zip_path = self._export_zip_to_separate_zip(
                        zip_loader_data, 
                        event_date_str,
                        zip_filename,
                        zip_index + 1,
                        len(self.selected_files)
                    )
                    
                    # Track generated ZIP file
                    if generated_zip_path:
                        bulk_generated_zips.append(generated_zip_path)
                
                # Cleanup temporary directories
                for temp_dir in temp_dirs:
                    try:
                        shutil.rmtree(temp_dir)
                    except Exception:
                        pass  # Ignore cleanup errors
                
                # Create master ZIP containing all generated ZIPs and show email dialog
                if bulk_generated_zips:
                    self._create_master_zip_and_email(bulk_generated_zips)
                
                self.status_var.set(f"Processed {len(self.selected_files)} ZIP file(s) - {total_rows_processed} total rows processed")
                return
            else:
                # Normal processing mode: process individual files
                # Validate dates (ensure widgets are created)
                if not (self.event_date_entry and self.settlement_date_entry and self.actual_date_entry):
                    # Wait a moment for async creation
                    self.update_idletasks()
                    if not (self.event_date_entry and self.settlement_date_entry and self.actual_date_entry):
                        messagebox.showerror("Error", "Date widgets are still initializing. Please wait a moment.")
                        self.status_var.set("Error: Date widgets not ready")
                        self.status_label.config(fg="#dc3545")  # Red color for errors
                        return
                try:
                    event_date = self.event_date_entry.get_date()
                    settlement_date = self.settlement_date_entry.get_date()
                    actual_date = self.actual_date_entry.get_date()
                except Exception as e:
                    messagebox.showerror("Error", f"Invalid date selection: {str(e)}")
                    self.status_var.set("Error: Invalid date")
                    self.status_label.config(fg="#dc3545")  # Red color for errors
                    return

                # Validate row and column inputs
                # Check if fallback checkbox is checked
                if self.fallback_var.get():
                    # Use fallback values: Row 10, Column B
                    read_row = 10
                    column_letter = "B"
                    read_col = self._column_letter_to_index(column_letter)
                else:
                    # Use values from input fields
                    try:
                        read_row = int(self.read_row_var.get())
                        
                        if read_row < 1:
                            messagebox.showwarning("Warning", "Read From Row must be at least 1")
                            return
                        
                        # Convert column letter to index
                        column_letter = self.read_col_var.get().strip().upper()
                        if not column_letter:
                            messagebox.showwarning("Warning", "Please enter a valid column letter (A, B, C, etc.)")
                            return
                        
                        read_col = self._column_letter_to_index(column_letter)
                        
                        if read_col < 0:
                            messagebox.showwarning("Warning", "Invalid column letter. Please use A-Z or AA-ZZ format.")
                            return
                        
                        # Save read configuration for next time (only if not using fallback)
                        self._save_read_config(read_row, column_letter)
                    except ValueError:
                        messagebox.showerror("Error", "Read From Row must be a valid number")
                        self.status_var.set("Error: Invalid row value")
                        self.status_label.config(fg="#dc3545")  # Red color for errors
                        return
                    except Exception as e:
                        messagebox.showerror("Error", f"Invalid column input: {str(e)}")
                        self.status_var.set("Error: Invalid column value")
                        self.status_label.config(fg="#dc3545")  # Red color for errors
                        return
                
                # Process individual files
                for file_index, file_path in enumerate(self.selected_files):
                    self.status_var.set(f"Processing file {file_index + 1}/{len(self.selected_files)}: {os.path.basename(file_path)}")
                    
                    # Read file with correct parameters:
                    # header_row: 1-based row number where headers are (use read_row)
                    # header_start_col: 0-based column index where headers start (use read_col)
                    df = self.read_dynamic_file(
                        file_path=file_path,
                        header_row=read_row,  # Row number (1-based) where header is located
                        header_start_col=read_col,  # Column index (0-based) where headers start
                        sheet_name=0
                    )
                    
                    # Process each row in the current file
                    for index, row in df.iterrows():
                        try:
                            trading_code = str(row.get(TradingCode, "")).strip()
                        except Exception as e:
                            raise Exception(f"Error getting TradingCode at row {index} in file {os.path.basename(file_path)}: {e}")
                        
                        try:
                            # Use single asio_sf4_ft config for all trading codes
                            # LocationAccount will be set dynamically based on trading code
                            # Pass dates and DataFrame columns
                            data_row = self._prepare_data_row(
                                row, 
                                asio_sf4_ft_config, 
                                trading_code, 
                                trading_code_mapping, 
                                columns=df.columns.tolist(),
                                event_date=event_date,
                                settlement_date=settlement_date,
                                actual_date=actual_date
                            )
                        except Exception as e:
                            raise Exception(f"Error in _prepare_data_row at row {index} in file {os.path.basename(file_path)}: {e}")

                        self.loader_data[trading_code].append(data_row)
                        total_rows_processed += 1
                    
                    total_files_processed += 1
            
                # Normal mode: combine all data first, then export
                # Dynamic dictionary to store data by trading code (shared across all files)
                self.loader_data = defaultdict(list)
                
                for file_index, file_path in enumerate(self.selected_files):
                    self.status_var.set(f"Processing file {file_index + 1}/{len(self.selected_files)}: {os.path.basename(file_path)}")
                    
                    # Read file with correct parameters:
                    # header_row: 1-based row number where headers are (use read_row)
                    # header_start_col: 0-based column index where headers start (use read_col)
                    df = self.read_dynamic_file(
                        file_path=file_path,
                        header_row=read_row,  # Row number (1-based) where header is located
                        header_start_col=read_col,  # Column index (0-based) where headers start
                        sheet_name=0
                    )
                    
                    # Process each row in the current file
                    for index, row in df.iterrows():
                        try:
                            trading_code = str(row.get(TradingCode, "")).strip()
                        except Exception as e:
                            raise Exception(f"Error getting TradingCode at row {index} in file {os.path.basename(file_path)}: {e}")
                        
                        try:
                            # Use single asio_sf4_ft config for all trading codes
                            # LocationAccount will be set dynamically based on trading code
                            # Pass dates and DataFrame columns
                            data_row = self._prepare_data_row(
                                row, 
                                asio_sf4_ft_config, 
                                trading_code, 
                                trading_code_mapping, 
                                columns=df.columns.tolist(),
                                event_date=event_date,
                                settlement_date=settlement_date,
                                actual_date=actual_date
                            )
                        except Exception as e:
                            raise Exception(f"Error in _prepare_data_row at row {index} in file {os.path.basename(file_path)}: {e}")

                        self.loader_data[trading_code].append(data_row)
                        total_rows_processed += 1
                    
                    total_files_processed += 1
            
            self.status_var.set(f"Processed {total_files_processed} file(s) successfully - {total_rows_processed} total rows processed")
            
            # Automatically export to template after processing all files
            self._export_to_template()
        except Exception as e:
            # Cleanup temporary directories on error
            for temp_dir in temp_dirs:
                try:
                    shutil.rmtree(temp_dir)
                except Exception:
                    pass  # Ignore cleanup errors
            
            messagebox.showerror("Error", f"Failed to process files: {str(e)}")
            self.status_var.set(f"Error: {str(e)}")
            self.status_label.config(fg="#dc3545")  # Red color for errors
            return
    
    def _create_output_files_for_zip(self, zip_loader_data, event_date_str, zip_filename):
        """Create output files (Excel/CSV) for a single zip file's data.
        
        Args:
            zip_loader_data: Dictionary of trading_code -> list of data rows
            event_date_str: Date string for filename (DDMMYYYY format)
            zip_filename: Original zip filename (for naming output files)
        
        Returns:
            list: List of tuples (file_io, filename) for all output files created
        """
        output_files = []
        
        # Get zip base name without extension for naming
        zip_basename = os.path.splitext(zip_filename)[0]
        
        # Create files for each trading code
        for trading_code, data_list in sorted(zip_loader_data.items()):
            if not data_list:
                continue
            
            # Convert list of lists to list of dicts
            template_data_dicts = []
            for row in data_list:
                row_dict = {}
                for idx, header in enumerate(sub_fund_4_headers):
                    if idx < len(row):
                        row_dict[header] = row[idx]
                    else:
                        row_dict[header] = ''
                template_data_dicts.append(row_dict)
            
            if not template_data_dicts:
                continue
            
            # Create Excel file if Excel format is selected
            if self.export_excel_var.get():
                excel_filename = f"ASIO_SF4_{trading_code}_{event_date_str}.xlsx"
                excel_file, excel_name = output_save_in_template(
                    template_data_dicts,
                    sub_fund_4_headers,
                    excel_filename
                )
                output_files.append((excel_file, excel_name))
            
            # Create CSV file if CSV format is selected
            if self.export_csv_var.get():
                csv_filename = f"ASIO_SF4_{trading_code}_{event_date_str}.csv"
                # Create CSV in memory
                csv_buffer = io.StringIO()
                csv_writer = csv.DictWriter(csv_buffer, fieldnames=sub_fund_4_headers)
                csv_writer.writeheader()
                csv_writer.writerows(template_data_dicts)
                csv_content = csv_buffer.getvalue()
                csv_buffer.close()
                
                # Convert to bytes
                csv_bytes = io.BytesIO(csv_content.encode('utf-8'))
                output_files.append((csv_bytes, csv_filename))
        
        return output_files
    
    def _export_zip_to_separate_zip(self, zip_loader_data, event_date_str, zip_filename, zip_index, total_zips):
        """Export output files from a single zip into its own separate ZIP file.
        
        Args:
            zip_loader_data: Dictionary of trading_code -> list of data rows
            event_date_str: Date string for filename (DDMMYYYY format)
            zip_filename: Original zip filename
            zip_index: Current zip index (1-based)
            total_zips: Total number of zips being processed
        
        Returns:
            str: Path to generated ZIP file, or None if cancelled/failed
        """
        if not zip_loader_data:
            return None
        
        # Create output files for this zip
        output_files = self._create_output_files_for_zip(
            zip_loader_data, 
            event_date_str,
            zip_filename
        )
        
        if not output_files:
            return
        
        # Get output directory from parent (set during bulk processing)
        bulk_export_dir = getattr(self, 'bulk_export_dir', None)
        if not bulk_export_dir:
            return None
        
        # Create filename for this zip (based on date from zip filename)
        zip_output_filename = f"ASIO_Sub_Fund_4_FT_Trades_{event_date_str}.zip"
        out_path = os.path.join(bulk_export_dir, zip_output_filename)
        
        # Update status
        self.status_var.set(f"Creating output ZIP {zip_index}/{total_zips}: {zip_output_filename}")
        
        try:
            # Create zip file with output files from this zip
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                # Add all output files
                for file_io, filename in output_files:
                    file_io.seek(0)
                    zip_file.writestr(filename, file_io.read())
            
            # Save zip file
            zip_buffer.seek(0)
            with open(out_path, 'wb') as f:
                f.write(zip_buffer.read())
            
            # Close all file buffers
            for file_io, _ in output_files:
                file_io.close()
            zip_buffer.close()
            
            # Return the generated ZIP path
            return out_path
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export ZIP {zip_index}: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def _export_to_template(self):
        """Export data to template format (ZIP with Excel files separated by trading code)."""
        if not hasattr(self, 'loader_data') or not self.loader_data:
            messagebox.showinfo("Nothing to export", "Process files first to generate template data.")
            return
        
        # Get event date for filename
        try:
            if self.event_date_entry:
                event_date = self.event_date_entry.get_date()
                # Format date as DDMMYYYY for filename
                event_date_str = event_date.strftime("%d%m%Y")
            else:
                from datetime import datetime
                event_date_str = datetime.now().strftime("%d%m%Y")
        except Exception:
            # If event date is not available, use current date or default
            from datetime import datetime
            event_date_str = datetime.now().strftime("%d%m%Y")
        
        # Ask output path for zip file
        zip_filename = f"ASIO_Sub_Fund_4_FT_Trades_{event_date_str}.zip"
        out_path = filedialog.asksaveasfilename(
            title="Save Template Data",
            defaultextension=".zip",
            filetypes=[["ZIP Files", "*.zip"]],
            initialfile=zip_filename
        )
        if not out_path:
            self.status_var.set("Export cancelled by user")
            return
        
        # Validate output format selection
        if not self.export_excel_var.get() and not self.export_csv_var.get():
            messagebox.showwarning("Warning", "Please select at least one output format (Excel or CSV).")
            return
        
        # Show spinner (non-blocking)
        loader = LoadingSpinner(self, text="Exporting templates...")

        def task():
            try:
                excel_files = []
                csv_files = []
                
                # Create files for each trading code
                for trading_code, data_list in sorted(self.loader_data.items()):
                    if not data_list:
                        continue
                    
                    # Convert list of lists to list of dicts
                    # Each row in data_list is a list matching sub_fund_4_headers order
                    template_data_dicts = []
                    for row in data_list:
                        row_dict = {}
                        for idx, header in enumerate(sub_fund_4_headers):
                            if idx < len(row):
                                row_dict[header] = row[idx]
                            else:
                                row_dict[header] = ''
                        template_data_dicts.append(row_dict)
                    
                    if not template_data_dicts:
                        continue
                    
                    # Create Excel file if Excel format is selected
                    if self.export_excel_var.get():
                        excel_filename = f"ASIO_SF4_{trading_code}_{event_date_str}.xlsx"
                        excel_file, excel_name = output_save_in_template(
                            template_data_dicts,
                            sub_fund_4_headers,
                            excel_filename
                        )
                        excel_files.append((excel_file, excel_name))
                    
                    # Create CSV file if CSV format is selected
                    if self.export_csv_var.get():
                        csv_filename = f"ASIO_SF4_{trading_code}_{event_date_str}.csv"
                        # Create CSV in memory
                        csv_buffer = io.StringIO()
                        csv_writer = csv.DictWriter(csv_buffer, fieldnames=sub_fund_4_headers)
                        csv_writer.writeheader()
                        csv_writer.writerows(template_data_dicts)
                        csv_content = csv_buffer.getvalue()
                        csv_buffer.close()
                        
                        # Convert to bytes
                        csv_bytes = io.BytesIO(csv_content.encode('utf-8'))
                        csv_files.append((csv_bytes, csv_filename))
                
                if not excel_files and not csv_files:
                    loader.close()
                    messagebox.showwarning("Warning", "No template data to export.")
                    return
                
                # Create zip file with all files
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    # Add Excel files
                    for excel_file, excel_name in excel_files:
                        excel_file.seek(0)
                        zip_file.writestr(excel_name, excel_file.read())
                    
                    # Add CSV files
                    for csv_file, csv_name in csv_files:
                        csv_file.seek(0)
                        zip_file.writestr(csv_name, csv_file.read())
                
                # Save zip file
                zip_buffer.seek(0)
                with open(out_path, 'wb') as f:
                    f.write(zip_buffer.read())
                
                # Close all file buffers
                for excel_file, _ in excel_files:
                    excel_file.close()
                for csv_file, _ in csv_files:
                    csv_file.close()
                zip_buffer.close()
                
                # Create file list for message
                file_list = []
                if excel_files:
                    file_list.extend([name for _, name in excel_files])
                if csv_files:
                    file_list.extend([name for _, name in csv_files])
                
                loader.close()
                
                email_zip_path = out_path
                email_file_list = file_list.copy()
                
                def show_success_and_dialog():
                    zip_filename = os.path.basename(email_zip_path)
                    success_msg = f"✓ Success! Template exported: {zip_filename} ({len(email_file_list)} files)"
                    self.status_var.set(success_msg)
                    self.status_label.config(fg="#28a745")  # Green color
                    self.after(100, lambda: self._show_email_dialog(email_zip_path, email_file_list))
                
                self.after(0, show_success_and_dialog)
            except Exception as e:
                loader.close()
                messagebox.showerror("Error", f"Failed to export template data: {e}")
                import traceback
                traceback.print_exc()
        
        # Run heavy work in thread
        threading.Thread(target=task, daemon=True).start()
    
    def _create_master_zip_and_email(self, zip_paths):
        """Extract all files from all generated ZIPs and create a combined ZIP with all individual files.
        
        Args:
            zip_paths: List of paths to all generated ZIP files
        """
        if not zip_paths:
            return
        
        try:
            # Use the same directory that was selected at the start (no dialog)
            bulk_export_dir = getattr(self, 'bulk_export_dir', None)
            if not bulk_export_dir:
                # Fallback: use directory of first ZIP
                bulk_export_dir = os.path.dirname(zip_paths[0]) if zip_paths else os.getcwd()
            
            # Create combined ZIP in the same directory
            combined_zip_filename = f"ASIO_Sub_Fund_4_FT_Trades_All_{datetime.now().strftime('%d%m%Y')}.zip"
            combined_zip_path = os.path.join(bulk_export_dir, combined_zip_filename)
            
            # Extract all files from all ZIPs and create combined ZIP
            all_files_list = []
            with zipfile.ZipFile(combined_zip_path, 'w', zipfile.ZIP_DEFLATED) as combined_zip:
                for zip_path in zip_paths:
                    if os.path.exists(zip_path):
                        # Extract all files from this ZIP
                        with zipfile.ZipFile(zip_path, 'r') as source_zip:
                            for file_info in source_zip.namelist():
                                # Skip directories
                                if file_info.endswith('/'):
                                    continue
                                
                                # Read file content from source ZIP
                                file_content = source_zip.read(file_info)
                                
                                # Add to combined ZIP (use just filename to avoid conflicts)
                                # If duplicate filenames exist, keep all with unique names
                                filename = os.path.basename(file_info)
                                if filename in all_files_list:
                                    # Add index to make unique
                                    base, ext = os.path.splitext(filename)
                                    counter = 1
                                    while filename in all_files_list:
                                        filename = f"{base}_{counter}{ext}"
                                        counter += 1
                                
                                all_files_list.append(filename)
                                combined_zip.writestr(filename, file_content)
            
            # Show success message in green on status bar and open email dialog
            def show_success_and_dialog():
                success_msg = f"✓ Success! All {len(zip_paths)} ZIP file(s) processed. Combined ZIP: {combined_zip_filename} ({len(all_files_list)} files)"
                self.status_var.set(success_msg)
                self.status_label.config(fg="#28a745")  # Green color
                self.after(100, lambda: self._show_email_dialog(combined_zip_path, all_files_list))
            
            self.after(0, show_success_and_dialog)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create combined ZIP: {e}")
            import traceback
            traceback.print_exc()
    
    def _show_email_dialog(self, zip_path, file_list):
        """Show email dialog for sending files via Outlook."""
        try:
            from .email_dialog import EmailDialog
            EmailDialog(self, zip_path, file_list)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open email dialog: {e}")
            import traceback
            traceback.print_exc()
