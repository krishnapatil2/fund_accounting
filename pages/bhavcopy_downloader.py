import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import threading
from datetime import datetime, date, timedelta
from tkcalendar import DateEntry

from my_app.pages.loading import LoadingSpinner

# LAZY IMPORTS - Heavy libraries imported only when needed
# pandas, requests, zipfile will be imported in download methods


class BhavcopyDownloaderPage(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg="#ecf0f1")

        # Title
        title = tk.Label(
            self, 
            text="📥 Download NSE Bhavcopy", 
            font=("Arial", 20, "bold"), 
            bg="#ecf0f1", 
            fg="#2c3e50"
        )
        title.pack(pady=30)

        # Main container with padding
        main_container = tk.Frame(self, bg="#ecf0f1")
        main_container.pack(fill="both", expand=True, padx=30, pady=20)

        # Controls Frame - Simple card design
        controls_card = tk.Frame(main_container, bg="white", relief="flat", bd=1, highlightbackground="#bdc3c7", highlightthickness=1)
        controls_card.pack(fill="both", expand=True, pady=10)
        
        controls = tk.Frame(controls_card, bg="white")
        controls.pack(fill="both", expand=True, padx=25, pady=35)

        # Checkbox for date range
        checkbox_row = tk.Frame(controls, bg="white")
        checkbox_row.pack(fill="x", pady=(0, 10))
        
        self.date_range_var = tk.BooleanVar(value=False)
        date_range_checkbox = tk.Checkbutton(
            checkbox_row,
            text="Download date range (From Date to To Date)",
            variable=self.date_range_var,
            font=("Arial", 11, "bold"),
            bg="white",
            fg="#2c3e50",
            selectcolor="white",
            activebackground="white",
            activeforeground="#2c3e50",
            command=self._toggle_date_range
        )
        date_range_checkbox.pack(side="left", padx=(0, 10))

        # Container frame for date selection (stays in same position)
        self.date_container = tk.Frame(controls, bg="white")
        self.date_container.pack(fill="x", pady=10)
        
        # Set default to yesterday (last trading day)
        today = datetime.now().date()
        default_date = today - timedelta(days=1)

        # Single Date Selection Row
        self.single_date_row = tk.Frame(self.date_container, bg="white")
        self.single_date_row.pack(fill="x")
        
        date_label = tk.Label(
            self.single_date_row, 
            text="📅 Select Date:", 
            font=("Arial", 12, "bold"), 
            bg="white", 
            fg="#2c3e50",
            width=15,
            anchor="w"
        )
        date_label.pack(side="left", padx=(0, 10))
        
        self.date_entry = DateEntry(
            self.single_date_row,
            date_pattern='dd/MM/yyyy',
            font=("Arial", 11),
            width=15,
            background="#3498db",
            foreground="white",
            borderwidth=2,
            relief="flat"
        )
        self.date_entry.set_date(default_date)
        self.date_entry.pack(side="left", padx=5)

        # Date Range Selection (initially hidden)
        self.date_range_frame = tk.Frame(self.date_container, bg="white")
        
        # From Date Row
        from_date_row = tk.Frame(self.date_range_frame, bg="white")
        from_date_row.pack(fill="x", pady=(0, 5))
        
        from_date_label = tk.Label(
            from_date_row,
            text="📅 From Date:",
            font=("Arial", 12, "bold"),
            bg="white",
            fg="#2c3e50",
            width=15,
            anchor="w"
        )
        from_date_label.pack(side="left", padx=(0, 10))
        
        self.from_date_entry = DateEntry(
            from_date_row,
            date_pattern='dd/MM/yyyy',
            font=("Arial", 11),
            width=15,
            background="#3498db",
            foreground="white",
            borderwidth=2,
            relief="flat"
        )
        self.from_date_entry.set_date(default_date)
        self.from_date_entry.pack(side="left", padx=5)

        # To Date Row
        to_date_row = tk.Frame(self.date_range_frame, bg="white")
        to_date_row.pack(fill="x")
        
        to_date_label = tk.Label(
            to_date_row,
            text="📅 To Date:",
            font=("Arial", 12, "bold"),
            bg="white",
            fg="#2c3e50",
            width=15,
            anchor="w"
        )
        to_date_label.pack(side="left", padx=(0, 10))
        
        self.to_date_entry = DateEntry(
            to_date_row,
            date_pattern='dd/MM/yyyy',
            font=("Arial", 11),
            width=15,
            background="#3498db",
            foreground="white",
            borderwidth=2,
            relief="flat"
        )
        self.to_date_entry.set_date(default_date)
        self.to_date_entry.pack(side="left", padx=5)

        # Save Path Selection Row
        path_row = tk.Frame(controls, bg="white")
        path_row.pack(fill="x", pady=10)
        
        path_label = tk.Label(
            path_row, 
            text="📁 Save Location:", 
            font=("Arial", 12, "bold"), 
            bg="white", 
            fg="#2c3e50",
            width=15,
            anchor="w"
        )
        path_label.pack(side="left", padx=(0, 10))

        self.save_path_var = tk.StringVar()
        
        path_entry = tk.Entry(
            path_row, 
            textvariable=self.save_path_var, 
            width=50,
            font=("Arial", 10),
            relief="solid",
            bd=1,
            bg="#f8f9fa"
        )
        path_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        browse_btn = tk.Button(
            path_row, 
            text="Browse", 
            command=self._browse_folder,
            bg="#3498db", 
            fg="white", 
            relief="flat", 
            padx=15, 
            pady=5,
            font=("Arial", 10, "bold"),
            cursor="hand2"
        )
        browse_btn.pack(side="left", padx=5)

        # Download Button Row
        button_row = tk.Frame(controls, bg="white")
        button_row.pack(fill="x", pady=(20, 10))
        
        self.download_btn = tk.Button(
            button_row,
            text="⬇ Download Bhavcopy",
            command=self._download_bhavcopy,
            bg="#27ae60",
            fg="white",
            relief="flat",
            padx=25,
            pady=12,
            font=("Arial", 12, "bold"),
            cursor="hand2"
        )
        self.download_btn.pack(side="left", padx=5)

        # Scrollable Status Text Widget
        status_frame = tk.Frame(controls, bg="white")
        status_frame.pack(fill="both", expand=True, pady=(15, 10))
        
        status_label_title = tk.Label(
            status_frame,
            text="Status:",
            font=("Arial", 11, "bold"),
            bg="white",
            fg="#2c3e50",
            anchor="w"
        )
        status_label_title.pack(fill="x", pady=(0, 5))
        
        self.status_text = scrolledtext.ScrolledText(
            status_frame,
            font=("Arial", 11),
            bg="#f8f9fa",
            fg="#34495e",
            wrap=tk.WORD,
            relief="solid",
            bd=1,
            height=8,
            state="disabled",
            padx=10,
            pady=10
        )
        self.status_text.pack(fill="both", expand=True)
        
        # Configure text tags for formatting
        self.status_text.tag_config("success", font=("Arial", 11, "bold"), foreground="#27ae60")

        # Data holders
        self.loading_spinner = None
        self.download_thread = None

    def _toggle_date_range(self):
        """Show/hide date range fields based on checkbox state"""
        if self.date_range_var.get():
            # Show date range fields, hide single date
            self.single_date_row.pack_forget()
            self.date_range_frame.pack(fill="x")
        else:
            # Show single date, hide date range fields
            self.date_range_frame.pack_forget()
            self.single_date_row.pack(fill="x")

    def _browse_folder(self):
        """Browse for folder to save bhavcopy"""
        folder = filedialog.askdirectory(
            title="Select Folder to Save Bhavcopy",
            initialdir=self.save_path_var.get() if self.save_path_var.get() else os.getcwd()
        )
        if folder:
            self.save_path_var.set(folder)

    def _update_status(self, message, is_success=False):
        """Update status message in scrollable text widget"""
        self.status_text.config(state="normal")
        self.status_text.delete(1.0, tk.END)
        self.status_text.insert(1.0, message)
        
        # Apply formatting based on success status
        if is_success:
            # Apply success tag (bold and green)
            self.status_text.tag_add("success", 1.0, tk.END)
        else:
            # Remove success tag for regular messages (default color)
            self.status_text.tag_remove("success", 1.0, tk.END)
        
        self.status_text.config(state="disabled")
        # Auto-scroll to bottom to show latest message
        self.status_text.see(tk.END)

    def _download_bhavcopy(self):
        """Start download in a separate thread"""
        save_path = self.save_path_var.get().strip()
        
        if not save_path:
            messagebox.showerror("Error", "Please select a folder to save the bhavcopy.")
            return

        today = date.today()
        
        # Check if date range is selected
        if self.date_range_var.get():
            from_date = self.from_date_entry.get_date()
            to_date = self.to_date_entry.get_date()
            
            # Validate date range
            if from_date > to_date:
                messagebox.showerror("Error", "From Date cannot be greater than To Date.")
                return
            
            if to_date > today:
                messagebox.showerror("Error", "Cannot download bhavcopy for future dates. Please select today or a past date.")
                return
            
            # Check if date range is too large
            date_diff = (to_date - from_date).days
            if date_diff > 365:
                response = messagebox.askyesno(
                    "Warning",
                    f"The selected date range is {date_diff} days (more than a year). "
                    "This may take a long time and some files might not be available. Do you want to continue?"
                )
                if not response:
                    return
            
            # Create save directory if it doesn't exist
            try:
                os.makedirs(save_path, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to create directory: {str(e)}")
                return

            # Disable download button
            self.download_btn.config(state="disabled")
            
            # Clear previous status
            self._update_status("Downloading...", is_success=False)
            
            # Show loading spinner
            self.loading_spinner = LoadingSpinner(self, text="Downloading Bhavcopy Range...")
            
            # Start download thread for date range
            self.download_thread = threading.Thread(
                target=self._download_range_worker,
                args=(from_date, to_date, save_path),
                daemon=True
            )
            self.download_thread.start()
        else:
            # Single date download
            selected_date = self.date_entry.get_date()
            
            # Check if date is in the future
            if selected_date > today:
                messagebox.showerror("Error", "Cannot download bhavcopy for future dates. Please select today or a past date.")
                return

            # Check if date is too old (optional - NSE typically keeps data for a few months)
            old_date_limit = today - timedelta(days=365)
            if selected_date < old_date_limit:
                response = messagebox.askyesno(
                    "Warning",
                    f"The selected date ({selected_date.strftime('%d/%m/%Y')}) is more than a year old. "
                    "The file might not be available. Do you want to continue?"
                )
                if not response:
                    return

            # Create save directory if it doesn't exist
            try:
                os.makedirs(save_path, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to create directory: {str(e)}")
                return

            # Disable download button
            self.download_btn.config(state="disabled")
            
            # Clear previous status
            self._update_status("Downloading...", is_success=False)
            
            # Show loading spinner
            self.loading_spinner = LoadingSpinner(self, text="Downloading Bhavcopy...")
            
            # Start download thread
            self.download_thread = threading.Thread(
                target=self._download_worker,
                args=(selected_date, save_path),
                daemon=True
            )
            self.download_thread.start()

    def _download_worker(self, target_date, save_path):
        """Worker thread for downloading bhavcopy"""
        try:
            # Import heavy libraries here (lazy import)
            import pandas as pd
            import requests
            import zipfile

            def get_user_friendly_error(status_code, failure_reason):
                """Convert technical error messages to user-friendly ones"""
                # Handle HTTP status codes
                if status_code == 404:
                    return "holiday/invalid date"
                elif status_code == 403:
                    return "access denied"
                elif status_code == 500:
                    return "server error"
                elif status_code and status_code != 200:
                    return f"HTTP error ({status_code})"
                
                # Handle failure reason strings
                if failure_reason:
                    failure_lower = failure_reason.lower()
                    if "404" in failure_reason or "not found" in failure_lower:
                        return "holiday/invalid date"
                    elif "403" in failure_reason or "forbidden" in failure_lower:
                        return "access denied"
                    elif "500" in failure_reason or "server error" in failure_lower:
                        return "server error"
                    elif "not a valid zip file" in failure_lower or "invalid zip file" in failure_lower:
                        return "holiday/invalid date"
                    elif "empty or invalid response" in failure_lower:
                        return "holiday/invalid date"
                    elif "http" in failure_lower:
                        # Extract status code from HTTP error message
                        if "404" in failure_reason:
                            return "holiday/invalid date"
                        elif "403" in failure_reason:
                            return "access denied"
                        elif "500" in failure_reason:
                            return "server error"
                
                return failure_reason if failure_reason else "failed to get bhavcopy"

            # Format: BhavCopy_NSE_CM_0_0_0_YYYYMMDD_F_0000.csv.zip
            date_str = target_date.strftime("%Y%m%d")
            filename = f"BhavCopy_NSE_CM_0_0_0_{date_str}_F_0000.csv.zip"
            url = f"https://nsearchives.nseindia.com/content/cm/{filename}"
            
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Accept': '*/*',
                'Accept-Language': 'en-US,en;q=0.9',
                'Referer': 'https://www.nseindia.com/'
            }
            
            download_failed = False
            failure_reason = ""
            
            try:
                # Download ZIP file
                response = requests.get(url, headers=headers, timeout=30)
                
                if response.status_code == 200:
                    # Check if response content is valid (not empty and not an error page)
                    if len(response.content) < 100:  # Very small files are likely errors
                        download_failed = True
                        failure_reason = get_user_friendly_error(200, "empty or invalid response")
                    # Check if it's an HTML error page (common for 404s that return 200)
                    elif b'<html' in response.content[:500].lower() or b'<!doctype' in response.content[:500].lower():
                        download_failed = True
                        failure_reason = get_user_friendly_error(404, "HTTP 404")
                    # Check if content looks like a ZIP file (starts with PK signature)
                    elif not response.content.startswith(b'PK'):
                        download_failed = True
                        failure_reason = get_user_friendly_error(200, "not a valid ZIP file (likely error page)")
                    else:
                        # Valid ZIP file - proceed with extraction
                        # Save ZIP file
                        zip_path = os.path.join(save_path, filename)
                        
                        with open(zip_path, 'wb') as f:
                            f.write(response.content)
                        
                        # Try to extract ZIP file
                        try:
                            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                                zip_ref.extractall(save_path)
                            
                            # Find extracted CSV file
                            extracted_files = [f for f in os.listdir(save_path) 
                                             if f.endswith('.csv') and date_str in f]
                            
                            if extracted_files:
                                csv_path = os.path.join(save_path, extracted_files[0])
                                
                                # Try to read CSV to verify it's valid
                                try:
                                    df = pd.read_csv(csv_path)
                                    if len(df) > 0:  # Check if CSV has data
                                        record_count = len(df)
                                        
                                        # Close loading spinner
                                        if self.loading_spinner:
                                            self.after(0, self.loading_spinner.close)
                                            self.loading_spinner = None
                                        
                                        # Show success message in green
                                        success_msg = f"✅ Downloaded successfully! ({record_count:,} records)"
                                        self.after(0, lambda: self._update_status(success_msg, is_success=True))
                                        # Re-enable download button
                                        self.after(0, lambda: self.download_btn.config(state="normal"))
                                        return  # Success - exit early
                                    else:
                                        download_failed = True
                                        failure_reason = get_user_friendly_error(200, "CSV file is empty")
                                except Exception as csv_error:
                                    download_failed = True
                                    failure_reason = get_user_friendly_error(200, f"CSV read error: {str(csv_error)[:50]}")
                            else:
                                download_failed = True
                                failure_reason = get_user_friendly_error(200, "no CSV file found in ZIP")
                        except zipfile.BadZipFile:
                            download_failed = True
                            failure_reason = get_user_friendly_error(200, "invalid ZIP file")
                        except Exception as zip_error:
                            download_failed = True
                            failure_reason = get_user_friendly_error(200, f"ZIP extraction error: {str(zip_error)[:50]}")
                else:
                    download_failed = True
                    failure_reason = get_user_friendly_error(response.status_code, f"HTTP {response.status_code}")
                    
            except requests.exceptions.Timeout:
                download_failed = True
                failure_reason = get_user_friendly_error(None, "request timeout")
            except requests.exceptions.RequestException as req_error:
                download_failed = True
                error_msg = str(req_error)[:50].lower()
                # Check if it's a 404 or similar in the error message
                if "404" in error_msg or "not found" in error_msg:
                    failure_reason = get_user_friendly_error(404, "HTTP 404")
                else:
                    failure_reason = get_user_friendly_error(None, f"network error: {str(req_error)[:50]}")
            except Exception as e:
                download_failed = True
                error_msg = str(e)[:50].lower()
                if "404" in error_msg or "not found" in error_msg:
                    failure_reason = get_user_friendly_error(404, "HTTP 404")
                else:
                    failure_reason = get_user_friendly_error(None, f"error: {str(e)[:50]}")
            
            # If download failed, show error message
            if download_failed:
                # Close loading spinner
                if self.loading_spinner:
                    self.after(0, self.loading_spinner.close)
                    self.loading_spinner = None
                
                # Get user-friendly error message
                user_friendly_reason = get_user_friendly_error(None, failure_reason) if failure_reason else "holiday/invalid date"
                error_msg = f"❌ Failed to download: {user_friendly_reason}"
                date_str_formatted = target_date.strftime('%d/%m/%Y')
                
                self.after(0, lambda: self._update_status(error_msg, is_success=False))
                self.after(0, lambda: messagebox.showerror("Error", f"Failed to download bhavcopy for {date_str_formatted}.\n{user_friendly_reason}"))
                # Re-enable download button
                self.after(0, lambda: self.download_btn.config(state="normal"))
                
        except Exception as e:
            if self.loading_spinner:
                self.after(0, self.loading_spinner.close)
                self.loading_spinner = None
            error_msg = f"❌ Error: {str(e)[:80]}"
            self.after(0, lambda: self._update_status(error_msg, is_success=False))
            self.after(0, lambda: messagebox.showerror("Error", f"An unexpected error occurred:\n{str(e)}"))
            # Re-enable download button
            self.after(0, lambda: self.download_btn.config(state="normal"))

    def _download_range_worker(self, from_date, to_date, save_path):
        """Worker thread for downloading bhavcopy for a date range"""
        try:
            # Import heavy libraries here (lazy import)
            import pandas as pd
            import requests
            import zipfile
            import time

            def get_user_friendly_error(status_code, failure_reason):
                """Convert technical error messages to user-friendly ones"""
                # Handle HTTP status codes
                if status_code == 404:
                    return "holiday/invalid date"
                elif status_code == 403:
                    return "access denied"
                elif status_code == 500:
                    return "server error"
                elif status_code and status_code != 200:
                    return f"HTTP error ({status_code})"
                
                # Handle failure reason strings
                if failure_reason:
                    failure_lower = failure_reason.lower()
                    if "404" in failure_reason or "not found" in failure_lower:
                        return "holiday/invalid date"
                    elif "403" in failure_reason or "forbidden" in failure_lower:
                        return "access denied"
                    elif "500" in failure_reason or "server error" in failure_lower:
                        return "server error"
                    elif "not a valid zip file" in failure_lower or "invalid zip file" in failure_lower:
                        return "holiday/invalid date"
                    elif "empty or invalid response" in failure_lower:
                        return "holiday/invalid date"
                    elif "http" in failure_lower:
                        # Extract status code from HTTP error message
                        if "404" in failure_reason:
                            return "holiday/invalid date"
                        elif "403" in failure_reason:
                            return "access denied"
                        elif "500" in failure_reason:
                            return "server error"
                
                return failure_reason if failure_reason else "failed to get bhavcopy"

            successful_downloads = 0
            failed_dates = []  # List to store failed dates
            total_records = 0
            
            # Count total weekdays in range
            total_weekdays = 0
            temp_date = from_date
            while temp_date <= to_date:
                if temp_date.weekday() < 5:  # Monday to Friday
                    total_weekdays += 1
                temp_date += timedelta(days=1)
            
            # Iterate through date range
            current_date = from_date
            current_day = 0
            
            while current_date <= to_date:
                # Skip weekends (Monday=0, Friday=4)
                if current_date.weekday() < 5:
                    current_day += 1
                    
                    # Update status
                    progress_msg = f"Downloading {current_date.strftime('%d/%m/%Y')}... ({current_day}/{total_weekdays} files)"
                    self.after(0, lambda msg=progress_msg: self._update_status(msg, is_success=False))
                    
                    # Download for this date
                    date_str = current_date.strftime("%Y%m%d")
                    filename = f"BhavCopy_NSE_CM_0_0_0_{date_str}_F_0000.csv.zip"
                    url = f"https://nsearchives.nseindia.com/content/cm/{filename}"
                    
                    headers = {
                        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                        'Accept': '*/*',
                        'Accept-Language': 'en-US,en;q=0.9',
                        'Referer': 'https://www.nseindia.com/'
                    }
                    
                    download_failed = False
                    failure_reason = ""
                    try:
                        response = requests.get(url, headers=headers, timeout=30)
                        
                        if response.status_code == 200:
                            # Check if response content is valid (not empty and not an error page)
                            if len(response.content) < 100:  # Very small files are likely errors
                                download_failed = True
                                failure_reason = get_user_friendly_error(200, "empty or invalid response")
                            # Check if it's an HTML error page (common for 404s that return 200)
                            elif b'<html' in response.content[:500].lower() or b'<!doctype' in response.content[:500].lower():
                                download_failed = True
                                failure_reason = get_user_friendly_error(404, "HTTP 404")
                            # Check if content looks like a ZIP file (starts with PK signature)
                            elif not response.content.startswith(b'PK'):
                                download_failed = True
                                failure_reason = get_user_friendly_error(200, "not a valid ZIP file (likely error page)")
                            else:
                                # Valid ZIP file - proceed with extraction
                                # Save ZIP file
                                zip_path = os.path.join(save_path, filename)
                                
                                with open(zip_path, 'wb') as f:
                                    f.write(response.content)
                                
                                # Try to extract ZIP file
                                try:
                                    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                                        zip_ref.extractall(save_path)
                                    
                                    # Find extracted CSV file
                                    extracted_files = [f for f in os.listdir(save_path) 
                                                     if f.endswith('.csv') and date_str in f]
                                    
                                    if extracted_files:
                                        csv_path = os.path.join(save_path, extracted_files[0])
                                        # Try to read CSV to verify it's valid
                                        try:
                                            df = pd.read_csv(csv_path)
                                            if len(df) > 0:  # Check if CSV has data
                                                total_records += len(df)
                                                successful_downloads += 1
                                            else:
                                                download_failed = True
                                                failure_reason = get_user_friendly_error(200, "CSV file is empty")
                                        except Exception as csv_error:
                                            download_failed = True
                                            failure_reason = get_user_friendly_error(200, f"CSV read error: {str(csv_error)[:50]}")
                                    else:
                                        download_failed = True
                                        failure_reason = get_user_friendly_error(200, "no CSV file found in ZIP")
                                except zipfile.BadZipFile:
                                    download_failed = True
                                    failure_reason = get_user_friendly_error(200, "invalid ZIP file")
                                except Exception as zip_error:
                                    download_failed = True
                                    failure_reason = get_user_friendly_error(200, f"ZIP extraction error: {str(zip_error)[:50]}")
                        else:
                            download_failed = True
                            failure_reason = get_user_friendly_error(response.status_code, f"HTTP {response.status_code}")
                    except requests.exceptions.Timeout:
                        download_failed = True
                        failure_reason = get_user_friendly_error(None, "request timeout")
                    except requests.exceptions.RequestException as req_error:
                        download_failed = True
                        error_msg = str(req_error)[:50].lower()
                        # Check if it's a 404 or similar in the error message
                        if "404" in error_msg or "not found" in error_msg:
                            failure_reason = get_user_friendly_error(404, "HTTP 404")
                        else:
                            failure_reason = get_user_friendly_error(None, f"network error: {str(req_error)[:50]}")
                    except Exception as e:
                        download_failed = True
                        error_msg = str(e)[:50].lower()
                        if "404" in error_msg or "not found" in error_msg:
                            failure_reason = get_user_friendly_error(404, "HTTP 404")
                        else:
                            failure_reason = get_user_friendly_error(None, f"error: {str(e)[:50]}")
                    
                    # Track failed dates with reason - ALWAYS track if download failed
                    if download_failed:
                        failed_date_str = current_date.strftime('%d/%m/%Y')
                        # Ensure we have a user-friendly reason
                        if not failure_reason:
                            failure_reason = "failed to get bhavcopy"
                        user_friendly_reason = get_user_friendly_error(None, failure_reason)
                        failed_dates.append(f"{failed_date_str} - {user_friendly_reason}")
                    
                    # Small delay between requests
                    time.sleep(0.5)
                
                # Move to next day
                current_date += timedelta(days=1)
            
            # Close loading spinner
            if self.loading_spinner:
                self.after(0, self.loading_spinner.close)
                self.loading_spinner = None
            
            # Show success message with failed dates
            total_processed = successful_downloads + len(failed_dates)
            success_msg = f"✅ Download complete! {successful_downloads} files downloaded (out of {total_processed} weekdays processed)"
            if failed_dates:
                failed_count = len(failed_dates)
                success_msg += f"\n\n❌ {failed_count} failed date(s):\n"
                for failed_date in failed_dates:
                    # failed_date already contains the date and reason
                    success_msg += f"   • {failed_date}\n"
            self.after(0, lambda: self._update_status(success_msg, is_success=True))
            # Re-enable download button
            self.after(0, lambda: self.download_btn.config(state="normal"))
            
        except Exception as e:
            if self.loading_spinner:
                self.after(0, self.loading_spinner.close)
                self.loading_spinner = None
            error_msg = f"❌ Error: {str(e)[:80]}"
            self.after(0, lambda: self._update_status(error_msg, is_success=False))
            self.after(0, lambda: messagebox.showerror("Error", f"An unexpected error occurred:\n{str(e)}"))
            # Re-enable download button
            self.after(0, lambda: self.download_btn.config(state="normal"))
