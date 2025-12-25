import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import json
import zipfile
from datetime import datetime
from decimal import Decimal
import threading

from my_app.pages.loading import LoadingSpinner

# LAZY IMPORTS - Heavy libraries imported only when needed (in methods)
# This speeds up frame opening significantly
# pandas, openpyxl, and CONSTANTS will be imported in _process() method when actually needed


def _safe_decimal(value, precision=15):
    """Safely convert value to Decimal with specified precision, handling None and empty strings.
    
    Args:
        value: Value to convert to Decimal
        precision: Number of decimal places to maintain (default: 15)
    
    Returns:
        Decimal: Decimal value with specified precision (up to 15 decimal places)
    """
    from decimal import getcontext
    # Set precision context to ensure we can handle up to 15 decimal places
    getcontext().prec = 28  # Set precision high enough to handle 15 decimal places accurately
    
    if value in [None, ""]:
        return Decimal("0")
    
    cleaned = str(value).replace(",", "").strip()
    try:
        decimal_val = Decimal(cleaned)
        # Normalize to remove trailing zeros but preserve precision
        return decimal_val.normalize()
    except Exception:
        return Decimal("0")


def _format_date(date_str: str):
    """Format date string from one format to another.
    
    Args:
        date_str: Date string in format like '18Sep25'
    
    Returns:
        tuple: (yyyymmdd, mm-dd-YYYY) formats
    """
    date_obj = datetime.strptime(date_str, '%d%b%y')
    yyyymmdd = date_obj.strftime('%Y%m%d')
    mm_dd_yyyy = date_obj.strftime('%m-%d-%Y')
    return yyyymmdd, mm_dd_yyyy


class ASIOTradeLoaderPage(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg="#ecf0f1")

        # Title
        title = tk.Label(self, text="📑 ASIO Sub Fund 2 Trade Loader FNO", font=("Arial", 20, "bold"), bg="#ecf0f1", fg="#2c3e50")
        title.pack(pady=10)

        # Controls
        controls = tk.Frame(self, bg="#ecf0f1")
        controls.pack(fill="x", padx=20, pady=5)

        self.file_path_var = tk.StringVar()
        self.bhavcopy_path_var = tk.StringVar()

        # Trade File row
        trade_file_row = tk.Frame(controls, bg="#ecf0f1")
        trade_file_row.pack(fill="x", pady=2)
        tk.Label(trade_file_row, text="Trade File:", font=("Arial", 11), bg="#ecf0f1", fg="#2c3e50").pack(side="left")
        tk.Entry(trade_file_row, textvariable=self.file_path_var, width=60).pack(side="left", padx=8)
        tk.Button(trade_file_row, text="Browse", command=self._browse_file, bg="#3498db", fg="white", relief="flat", padx=10, pady=4).pack(side="left")

        # Bhavcopy File row
        bhavcopy_file_row = tk.Frame(controls, bg="#ecf0f1")
        bhavcopy_file_row.pack(fill="x", pady=2)
        tk.Label(bhavcopy_file_row, text="Bhavcopy File:", font=("Arial", 11), bg="#ecf0f1", fg="#2c3e50").pack(side="left")
        tk.Entry(bhavcopy_file_row, textvariable=self.bhavcopy_path_var, width=60).pack(side="left", padx=8)
        tk.Button(bhavcopy_file_row, text="Browse", command=self._browse_bhavcopy, bg="#3498db", fg="white", relief="flat", padx=10, pady=4).pack(side="left")

        # Buttons row
        buttons_row = tk.Frame(controls, bg="#ecf0f1")
        buttons_row.pack(fill="x", pady=5)
        tk.Button(buttons_row, text="Process", command=self._process, bg="#27ae60", fg="white", relief="flat", padx=14, pady=6, font=("Arial", 11, "bold")).pack(side="left", padx=10)
        tk.Button(buttons_row, text="Export Excel", command=self._export_excel, bg="#8e44ad", fg="white", relief="flat", padx=14, pady=6, font=("Arial", 11, "bold")).pack(side="left")
        tk.Button(buttons_row, text="Export to Template", command=self._export_to_template, bg="#d35400", fg="white", relief="flat", padx=14, pady=6, font=("Arial", 11, "bold")).pack(side="left", padx=10)

        # Status
        self.status_var = tk.StringVar(value="Load a file (CSV/XLS/XLSX) and click Process")
        tk.Label(self, textvariable=self.status_var, font=("Arial", 10), bg="#ecf0f1", fg="#7f8c8d").pack(fill="x", padx=20)

        # Single table view
        content = tk.Frame(self, bg="#ecf0f1")
        content.pack(fill="both", expand=True, padx=20, pady=10)

        # Table header with checkbox and search
        header = tk.Frame(content, bg="#ecf0f1")
        header.pack(fill="x", pady=(0, 4))
        tk.Label(header, text="Trade Data", font=("Arial", 13, "bold"), bg="#ecf0f1", fg="#2c3e50").pack(side="left")
        
        # Checkbox to toggle between unique and all data
        self.checkbox_frame = tk.Frame(header, bg="#ecf0f1")
        # Initially hidden - will be shown when data is loaded
        self.unique_checkbox_var = tk.BooleanVar(value=True)
        self.unique_checkbox = tk.Checkbutton(
            self.checkbox_frame, 
            text="Unique trade data based on securities names",
            variable=self.unique_checkbox_var,
            font=("Arial", 10),
            bg="#ecf0f1",
            fg="#2c3e50",
            selectcolor="#ecf0f1",
            command=self._on_checkbox_toggle  # Callback when checkbox is toggled
        )
        self.unique_checkbox.pack(side="left")
        
        # Search box
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._render_table())
        search_box = tk.Frame(header, bg="#ecf0f1")
        search_box.pack(side="right", padx=(0, 18))
        tk.Label(search_box, text="Search:", font=("Arial", 10), bg="#ecf0f1", fg="#2c3e50").pack(side="left")
        tk.Entry(search_box, textvariable=self.search_var, width=24).pack(side="left", padx=(6, 0))
        
        # Table frame
        table_frame = tk.Frame(content, bg="#ecf0f1")
        table_frame.pack(fill="both", expand=True)
        
        # Table columns
        self.table_columns = (
            "Date", "InstrumentType", "BuySell", "Qty", "Price", "TMCode",
            "Securitiy Nmaes", "UnderlyingInvestment", "StrikePrice",
            "Option/Future Type", "PutCallFlag", "ExpireDate"
        )
        self.tree = ttk.Treeview(table_frame, columns=self.table_columns, show="headings", height=12)
        # Column widths - adjust based on content
        column_widths = {
            "Date": 100,
            "InstrumentType": 120,
            "BuySell": 80,
            "Qty": 100,
            "Price": 100,
            "TMCode": 100,
            "Securitiy Nmaes": 180,
            "UnderlyingInvestment": 150,
            "StrikePrice": 100,
            "Option/Future Type": 130,
            "PutCallFlag": 100,
            "ExpireDate": 120
        }
        for col in self.table_columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=column_widths.get(col, 120), anchor="w")
        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        # Grid layout so the bottom scrollbar spans full width
        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, columnspan=2, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        # Data holders
        self.all_table_rows = []  # All trade records
        self.unique_table_rows = []  # Unique trade records (based on security names)
        self._template_data_1 = []  # Placeholder for first template data
        self._template_data_2 = []  # Placeholder for second template data
        self._template_data_3 = []  # Placeholder for third template data

    # ---- UI handlers ----
    def _browse_file(self):
        """Browse for CSV, XLS, or XLSX file."""
        path = filedialog.askopenfilename(
            title="Select Trade File", 
            filetypes=[
                ["All Supported Files", "*.csv *.xls *.xlsx"],
                ["CSV Files", "*.csv"],
                ["Excel Files", "*.xls *.xlsx"],
                ["XLS Files", "*.xls"],
                ["XLSX Files", "*.xlsx"],
                ["All Files", "*.*"]
            ]
        )
        if path:
            self.file_path_var.set(path)

    def _browse_bhavcopy(self):
        """Browse for Bhavcopy file (CSV, XLS, or XLSX)."""
        path = filedialog.askopenfilename(
            title="Select Bhavcopy File", 
            filetypes=[
                ["All Supported Files", "*.csv *.xls *.xlsx"],
                ["CSV Files", "*.csv"],
                ["Excel Files", "*.xls *.xlsx"],
                ["XLS Files", "*.xls"],
                ["XLSX Files", "*.xlsx"],
                ["All Files", "*.*"]
            ]
        )
        if path:
            self.bhavcopy_path_var.set(path)

    # ---- Core processing ----
    def _process(self):
        """Process the file (CSV/XLS/XLSX) and populate table."""
        # Lazy import heavy libraries only when processing (speeds up frame opening)
        from .helper import read_file
        
        # Clear table
        for it in self.tree.get_children():
            self.tree.delete(it)
        self.all_table_rows = []
        self.unique_table_rows = []
        self._template_data_1 = []
        self._template_data_2 = []
        self._template_data_3 = []
        
        # Hide checkbox initially
        self.checkbox_frame.pack_forget()

        # Load consolidated JSON
        from my_app.file_utils import get_app_directory
        app_dir = get_app_directory()
        consolidated_path = os.path.join(app_dir, "consolidated_data.json")

        try:
            if os.path.exists(consolidated_path):
                with open(consolidated_path, "r") as f:
                    consolidated_data = json.load(f)
            
            # Extract configuration from consolidated data
            asio_sf_2_trade_loader = consolidated_data.get("asio_sf_2_trade_loader", {})
            asio_sf_2_option_security = consolidated_data.get("asio_sf_2_option_security", {})
            asio_sf_2_future_security = consolidated_data.get("asio_sf_2_future_security", {})
            fno_tm_code_with_tm_name = consolidated_data.get("fno_tm_code_with_tm_name", {})
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load consolidated_data.json: {e}\nPlease ensure this file exists under my_app/.")
            return

        file_path = self.file_path_var.get().strip()
        if not file_path or not os.path.exists(file_path):
            messagebox.showwarning("File Missing", "Please select a valid Trade file (CSV/XLS/XLSX).")
            return

        # Read bhavcopy file if provided
        df_bhavcopy = None
        bhavcopy_path = self.bhavcopy_path_var.get().strip()
        if bhavcopy_path and os.path.exists(bhavcopy_path):
            try:
                # Read bhavcopy file - adjust parameters based on file format
                df_bhavcopy = read_file(
                    file_path=bhavcopy_path,
                    sheet_name=0,   # Always first sheet
                    start_row=0,    # Start from first row (adjust if needed)
                    header=True,
                    skip_blank_rows=False
                )
                # Clean headers (strip whitespace)
                df_bhavcopy.columns = df_bhavcopy.columns.str.strip()
                
                # Validate bhavcopy: Check if "Sgmt" column contains "FO"
                if not df_bhavcopy.empty:
                    # Check if "Sgmt" column exists (case-insensitive)
                    sgmt_col = None
                    for col in df_bhavcopy.columns:
                        if col.strip().upper() == 'SGMT':
                            sgmt_col = col
                            break
                    
                    if sgmt_col:
                        # Check if any row has "FO" in Sgmt column
                        if df_bhavcopy[sgmt_col].astype(str).str.strip().str.upper().isin(['FO']).any():
                            filename = os.path.basename(bhavcopy_path)
                            messagebox.showwarning(
                                "Wrong Bhavcopy File",
                                f"Warning: The selected bhavcopy file '{filename}' contains 'FO' in the 'Sgmt' column.\n\n"
                                "This is the wrong bhavcopy file. Please attach equity bhavcopy file."
                            )
            except Exception as e:
                messagebox.showwarning("Warning", f"Failed to read Bhavcopy file: {e}\nContinuing without bhavcopy data.")
                df_bhavcopy = None

        try:
            # Use read_file helper function to support CSV, XLS, and XLSX
            # Read from row 11 (0-based, so row 12 in Excel) - no end_row needed, reads until end naturally
            df_data = read_file(
                file_path=file_path,
                sheet_name=0,   # Always first sheet
                start_row=11,   # Row 12 becomes header (0-based index)
                header=True,
                skip_blank_rows=False
            )

            # Keep only B → end (skip column A)
            df_data = df_data.iloc[:, 1:]
            
            # Clean headers (strip whitespace)
            df_data.columns = df_data.columns.str.strip()
            
            # Filter rows where Date exists (not empty/NaN)
            # Try to find Date column (case-insensitive, handle variations)
            date_col = None
            for col in df_data.columns:
                if col.strip().lower() == 'date':
                    date_col = col
                    break
            
            if date_col:
                # Keep only rows where Date is not empty/NaN
                df_data = df_data[df_data[date_col].notna() & (df_data[date_col].astype(str).str.strip() != '')]
            else:
                # If Date column not found, show warning but continue
                messagebox.showwarning("Warning", "Date column not found. Processing all rows.")
            
            # Reset index after filtering
            df_data = df_data.reset_index(drop=True)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to read file: {e}")
            return
        
        # Process data - placeholder function, logic to be implemented
        left_table_data, right_table_data, asio_sub_fund_2_future, asio_sub_fund_2_option, template_data_3 = self._process_data(
            df_data, asio_sf_2_trade_loader, asio_sf_2_option_security, asio_sf_2_future_security, fno_tm_code_with_tm_name, df_bhavcopy
        )

        # Persist in instance and render
        self._df_data = df_data
        self._df_bhavcopy = df_bhavcopy  # Store bhavcopy data for later use
        self.all_table_rows = left_table_data  # All records
        self.unique_table_rows = right_table_data  # Unique records
        self.asio_sub_fund_2_future = asio_sub_fund_2_future
        self.asio_sub_fund_2_option = asio_sub_fund_2_option
        self._template_data_3 = template_data_3

        self._render_table()

        # Show the checkbox when data is loaded
        if len(self.all_table_rows) > 0:
            self.checkbox_frame.pack(side="left", padx=(15, 0))
        else:
            self.checkbox_frame.pack_forget()

        # Update status based on checkbox state
        self._update_status()

    def _process_data(self, df_data, asio_sf_2_trade_loader, asio_sf_2_option_security, asio_sf_2_future_security, fno_tm_code_with_tm_name, df_bhavcopy=None):
        """
        Process file data and generate table data and template data.
        
        Args:
            df_data: DataFrame from file (CSV/XLS/XLSX)
            asio_sf_2_trade_loader: Trade loader configuration
            asio_sf_2_option_security: Option security configuration
            asio_sf_2_future_security: Future security configuration
            fno_tm_code_with_tm_name: TM code to TM name mapping
            df_bhavcopy: Optional DataFrame from bhavcopy file (CSV/XLS/XLSX)
        
        Returns:
            tuple: (left_table_data, right_table_data, template_data_1, template_data_2, template_data_3)
            
        Note:
            - left_table_data: ALL trade records (complete data from file)
            - right_table_data: UNIQUE trade records filtered by "Securitiy Nmaes" (one record per unique security name)
            - Each row should be a list/tuple matching the column order:
              [Date, InstrumentType, BuySell, Qty, Price, TMCode, Securitiy Nmaes, 
               UnderlyingInvestment, StrikePrice, Option/Future Type, PutCallFlag, ExpireDate]
            - df_bhavcopy can be used for additional processing logic
        """
        # Lazy import pandas (used for pd.notna check)
        import pandas as pd
        from CONSTANTS import TM_NAME_HEADERS, ASIO_FUTURE_SF2_HEADER, ASIO_SUB_FUND_2_OPTION_SECURITY_HEADER
        
        left_table_data = []  # All trade records
        right_table_data = []  # Unique trade records based on "Securitiy Nmaes"
        asio_sub_fund_2_future = []
        asio_sub_fund_2_option = []
        template_data_3 = []
        
        # Create ticker to ISIN mapping from bhavcopy file if available
        ticker_isin_dict = {}
        if df_bhavcopy is not None and not df_bhavcopy.empty:
            try:
                # Check if required columns exist
                if 'TckrSymb' in df_bhavcopy.columns and 'ISIN' in df_bhavcopy.columns:
                    ticker_isin_dict = dict(zip(df_bhavcopy['TckrSymb'], df_bhavcopy['ISIN']))
            except Exception as e:
                # If mapping fails, continue without it
                ticker_isin_dict = {}
        
        # Group data by TM code with TM_NAME_HEADERS structure
        data_by_tm_code = {}  # {tm_code: [list of dicts matching TM_NAME_HEADERS]}
        
        # Track unique security names for right table
        seen_security_names = set()
        # Track unique security names for option data
        seen_option_securities = set()
        
        # Process each row in the file
        for idx, row in df_data.iterrows():
            # Extract values for each column (handle missing values)
            def safe_get(col, default=''):
                """Safely get value from row, handling missing columns"""
                if col in df_data.columns:
                    val = row.get(col, default)
                    return str(val).strip() if pd.notna(val) else default
                return default
            
            date_val = safe_get('Date', '')
            instrument_type_val = safe_get('InstrumentType', '')
            buy_sell_val = safe_get('BuySell', '')
            # Handle quantity - convert float string to int
            qty_str = safe_get('Qty', '0')
            try:
                qty_val = int(float(qty_str))
            except:
                qty_val = 0
            # Handle price - convert to Decimal with 15 digit precision
            price_str = safe_get('Price', '')
            price_val = _safe_decimal(price_str, precision=15)
            tm_code_val = safe_get('TMCode', '')
            
            # Look up TM name from fno_tm_code_with_tm_name
            # Keep TM code as string to preserve leading zeros (e.g., "07123")
            # Only convert if it's a float string like "13302.0" -> "13302"
            tm_code_str = str(tm_code_val).strip()
            # If it's a float string, convert to int string (remove .0)
            if '.' in tm_code_str:
                try:
                    tm_code_str = str(int(float(tm_code_str)))
                except:
                    pass  # Keep original if conversion fails
            
            tm_name_val = ''
            if tm_code_str and isinstance(fno_tm_code_with_tm_name, dict):
                # Try direct lookup first
                tm_name_val = str(fno_tm_code_with_tm_name.get(tm_code_str, '')).strip()
                # If not found and it's a number, try with leading zero (e.g., "07536")
                if not tm_name_val and tm_code_str.isdigit():
                    for key, value in fno_tm_code_with_tm_name.items():
                        if str(key).lstrip('0') == tm_code_str or str(key) == tm_code_str:
                            tm_name_val = str(value).strip()
                            break
            
            # If TM name not found, use default pattern
            if not tm_name_val and tm_code_str:
                tm_name_val = f"{tm_code_str}_not_found_tm_name"
            underlying_val = safe_get('UnderlyingCode', '')
            # Handle strike price - convert float string to int
            strike_price_str = safe_get('StrikePrice', '0')
            try:
                strike_price_val = int(float(strike_price_str))
            except:
                strike_price_val = 0
            option_future_val = safe_get('OptionType', '')
            expire_date_val = safe_get('ExpiryDate', '')
            
            from datetime import datetime
            # Parse date from format '28-10-2025' (dd-mm-yyyy)
            try:
                date_obj = datetime.strptime(expire_date_val, "%d-%m-%Y")
                expire_date_formatted = date_obj.strftime("%Y%m%d")
                mm_dd_yyyy = date_obj.strftime("%m-%d-%Y")
                expiry_date = mm_dd_yyyy + ' 23:59:59'
            except:
                expire_date_formatted = ''
                expiry_date = ''
            option_future_first_char = option_future_val[0] if option_future_val else ''

            security_name_val = f"NSE{underlying_val}{expire_date_formatted}{option_future_first_char}{strike_price_val}"
            put_call_flag_val = 1 if option_future_val == "PE" else 0

            # Map ticker symbol to ISIN for display if available
            underlying_display_val = underlying_val
            if ticker_isin_dict and underlying_val in ticker_isin_dict:
                underlying_display_val = ticker_isin_dict[underlying_val]

            # Create row data matching the column order
            # Use tm_code_str (preserves leading zeros) instead of tm_code_val (might be float)
            row_data = [
                date_val,
                instrument_type_val,
                buy_sell_val,
                qty_val,
                price_val,
                tm_code_str,  # Use processed string value (preserves leading zeros)
                security_name_val,
                underlying_display_val,  # Use ISIN if available, otherwise ticker symbol
                strike_price_val,
                option_future_val,
                put_call_flag_val,
                expire_date_val
            ]
            
            # Add ALL records to left table
            left_table_data.append(row_data)
            
            # Add Future security only when not Option (CE/PE)
            if option_future_val != "CE" and option_future_val != "PE":
                future_dict = {}
                for header in ASIO_FUTURE_SF2_HEADER:
                    if isinstance(asio_sf_2_future_security, dict) and asio_sf_2_future_security:
                        if header in asio_sf_2_future_security:
                            future_dict[header] = asio_sf_2_future_security.get(header, '')
                        else:
                            # Fallback default field mappings
                            if header == "Code":
                                future_dict[header] = security_name_val
                            elif header == "KeyValue":
                                future_dict[header] = security_name_val
                            elif header == "Description":
                                future_dict[header] = security_name_val
                            elif header == "ExtendedDescription":
                                future_dict[header] = security_name_val
                            elif header == "UnderlyingInvestment":
                                # Map ticker symbol to ISIN using bhavcopy data if available
                                if ticker_isin_dict and underlying_val in ticker_isin_dict:
                                    future_dict[header] = ticker_isin_dict[underlying_val]
                                else:
                                    future_dict[header] = underlying_val
                            elif header == "ExpireDate":
                                future_dict[header] = expiry_date
                            elif header == "StrikePrice":
                                future_dict[header] = 0

                # Add only unique future securities
                if future_dict and future_dict not in asio_sub_fund_2_future:
                    asio_sub_fund_2_future.append(future_dict)


            # Prepare option data based on ASIO_SUB_FUND_2_OPTION_SECURITY_HEADER and asio_sf_2_option_security
            # Only add unique records based on security_name_val
            if security_name_val and security_name_val not in seen_option_securities:
                if option_future_val == "CE" or option_future_val == "PE":
                    # Check if this is an option (not future)
                    seen_option_securities.add(security_name_val)
                    option_dict = {}
                    for header in ASIO_SUB_FUND_2_OPTION_SECURITY_HEADER:
                        if isinstance(asio_sf_2_option_security, dict) and asio_sf_2_option_security:
                            if header in asio_sf_2_option_security:
                                option_dict[header] = asio_sf_2_option_security.get(header, '')
                            else:
                                # Default mappings
                                if header == "Code":
                                    option_dict[header] = security_name_val
                                elif header == "KeyValue":
                                    option_dict[header] = security_name_val
                                elif header == "Description":
                                    option_dict[header] = security_name_val
                                elif header == "ExtendedDescription":
                                    option_dict[header] = security_name_val
                                elif header == "UnderlyingInvestment":
                                    # Map ticker symbol to ISIN using bhavcopy data if available
                                    if ticker_isin_dict and underlying_val in ticker_isin_dict:
                                        option_dict[header] = ticker_isin_dict[underlying_val]
                                    else:
                                        option_dict[header] = underlying_val
                                elif header == "ExpireDate":
                                    option_dict[header] = expiry_date
                                elif header == "StrikePrice":
                                    option_dict[header] = strike_price_val
                                elif header == "PutCallFlag":
                                    option_dict[header] = put_call_flag_val
                    
                    asio_sub_fund_2_option.append(option_dict)

            # Prepare data based on TM code using TM_NAME_HEADERS and asio_sf_2_trade_loader
            if tm_code_str:
                if tm_code_str not in data_by_tm_code:
                    data_by_tm_code[tm_code_str] = []
                
                # Create row dict based on TM_NAME_HEADERS using asio_sf_2_trade_loader mapping
                row_dict = {}

                if date_val:
                    # Parse the input (DD-MM-YYYY)
                    dt = datetime.strptime(date_val, "%d-%m-%Y")

                    # Convert to MM-DD-YYYY
                    formatted_date = dt.strftime("%m-%d-%Y")
                
                for header in TM_NAME_HEADERS:
                    # Get mapping from asio_sf_2_trade_loader (maps header to file column name)
                    if isinstance(asio_sf_2_trade_loader, dict) and asio_sf_2_trade_loader:
                        if header in asio_sf_2_trade_loader:
                            row_dict[header] = asio_sf_2_trade_loader.get(header, '')
                        else:
                            if header == "RecordType":
                                row_dict[header] = buy_sell_val
                            elif header == "KeyValue":
                                row_dict[header] = security_name_val
                            elif header == "Strategy":
                                row_dict[header] = tm_name_val
                            elif header == "Investment":
                                row_dict[header] = security_name_val
                            elif header == "EventDate":
                                row_dict[header] = formatted_date
                            elif header == "SettleDate":
                                row_dict[header] = formatted_date
                            elif header == "ActualSettleDate":
                                row_dict[header] = formatted_date
                            elif header == "Quantity":
                                row_dict[header] = qty_val
                            elif header == "Price":
                                row_dict[header] = price_val
                            else:                                
                                row_dict[header] = ''

                data_by_tm_code[tm_code_str].append(row_dict)
            
            # Add only UNIQUE records to right table based on "Securitiy Nmaes"
            if security_name_val and security_name_val not in seen_security_names:
                seen_security_names.add(security_name_val)
                right_table_data.append(row_data)
        
        # Store data_by_tm_code in instance for later use

        self.data_by_tm_code = data_by_tm_code
        
        return left_table_data, right_table_data, asio_sub_fund_2_future, asio_sub_fund_2_option, template_data_3

    # ---- Rendering with filtering ----
    def _render_table(self):
        """Render table with search filtering. Shows unique or all data based on checkbox."""
        # Clear table
        for it in self.tree.get_children():
            self.tree.delete(it)
        
        term = (self.search_var.get() if hasattr(self, 'search_var') else "").lower().strip()
        
        # Choose data source based on checkbox state
        if self.unique_checkbox_var.get():
            # Show unique data
            data_source = self.unique_table_rows
        else:
            # Show all data
            data_source = self.all_table_rows
        
        for r in data_source:
            if not term or term in " ".join(map(str, r)).lower():
                self.tree.insert("", "end", values=r)
    
    def _on_checkbox_toggle(self):
        """Callback when checkbox is toggled - refresh table display."""
        self._render_table()
        self._update_status()
    
    def _update_status(self):
        """Update status bar based on current data and checkbox state."""
        if self.unique_checkbox_var.get():
            unique_count = len(self.unique_table_rows)
            self.status_var.set(f"Processed {len(self.all_table_rows)} trade records (all) | Showing {unique_count} unique trade records")
        else:
            self.status_var.set(f"Showing {len(self.all_table_rows)} trade records (all data)")

    def _export_excel(self):
        """Export data to Excel file."""
        # Lazy import heavy libraries only when exporting
        import pandas as pd
        from openpyxl.styles import Font, Border
        
        if not self.all_table_rows and not self.unique_table_rows:
            messagebox.showinfo("Nothing to export", "Process a file first.")
            return

        # Ask output path
        out_path = filedialog.asksaveasfilename(
            title="Save Excel",
            defaultextension=".xlsx",
            filetypes=[["Excel", "*.xlsx"]],
            initialfile="ASIOTradeLoader_Output.xlsx"
        )
        if not out_path:
            return

        # Build DataFrames and write sheets
        try:
            # Get table data based on checkbox state (what's currently displayed)
            if self.unique_checkbox_var.get():
                table_data = self.unique_table_rows
            else:
                table_data = self.all_table_rows
            
            df_table = pd.DataFrame(table_data, columns=self.table_columns) if table_data else pd.DataFrame()

            with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                # Sheet 1: Original Data (no bold headers, no borders)
                df_data = getattr(self, "_df_data", None)
                if df_data is not None:
                    df_data.to_excel(writer, sheet_name="Original_Data", index=False)
                else:
                    pd.DataFrame().to_excel(writer, sheet_name="Original_Data", index=False)
                # Normalize header styling
                ws_orig = writer.book["Original_Data"]
                if ws_orig.max_row >= 1:
                    for cell in ws_orig[1]:
                        cell.font = Font(bold=False)
                        cell.border = Border()

                # Sheet 2: Trade Data (currently displayed - unique or all)
                sheet_name = "Unique_Trade_Data" if self.unique_checkbox_var.get() else "All_Trade_Data"
                df_table.to_excel(writer, sheet_name=sheet_name, index=False)

            messagebox.showinfo("Success", f"Excel exported to:\n{out_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export excel: {e}")

    def _export_to_template(self):
        """Export data to template format (ZIP with Excel files grouped by TM code)."""
        # Lazy import heavy libraries only when exporting
        from CONSTANTS import TM_NAME_HEADERS, ASIO_SUB_FUND_2_OPTION_SECURITY_HEADER
        from .helper import output_save_in_template, multiple_excels_to_zip
        
        if not hasattr(self, 'data_by_tm_code') or not self.data_by_tm_code:
            messagebox.showinfo("Nothing to export", "Process a file first to generate template data.")
            return

        # Ask output path for zip file
        out_path = filedialog.asksaveasfilename(
            title="Save Template Data",
            defaultextension=".zip",
            filetypes=[["ZIP Files", "*.zip"]],
            initialfile="FNO_ASIO_SF2_TradeLoader.zip"
        )
        if not out_path:
            return
        
        import time
        start_time = time.time()
        # Show spinner (non-blocking)
        loader = LoadingSpinner(self, text="Exporting templates...")

        def task():
            try:
                # Create Excel files using helper functions
                excel_files = []
                
                # Create one Excel file for each TM code
                for tm_code, data_list in self.data_by_tm_code.items():
                    if data_list:  # Only create if there's data
                        excel_file, excel_name = output_save_in_template(
                            data_list,  # List of dicts matching TM_NAME_HEADERS
                            TM_NAME_HEADERS,  # Headers from constant
                            f"TM_{tm_code}_template.xlsx"
                        )
                        excel_files.append((excel_file, excel_name))
                
                # Create Excel file for option security data
                if hasattr(self, 'asio_sub_fund_2_option') and self.asio_sub_fund_2_option:
                    option_excel_file, option_excel_name = output_save_in_template(
                        self.asio_sub_fund_2_option,  # List of dicts matching ASIO_SUB_FUND_2_OPTION_SECURITY_HEADER
                        ASIO_SUB_FUND_2_OPTION_SECURITY_HEADER,  # Headers from constant
                        "ASIO_Sub_Fund_2_Option_Security.xlsx"
                    )
                    excel_files.append((option_excel_file, option_excel_name))
                
                if hasattr(self, 'asio_sub_fund_2_future') and self.asio_sub_fund_2_future:
                    future_excel_file, future_excel_name = output_save_in_template(
                        self.asio_sub_fund_2_future,  # List of dicts matching ASIO_SUB_FUND_2_FUTURE_SECURITY_HEADER
                        ASIO_SUB_FUND_2_OPTION_SECURITY_HEADER,  # Headers from constant
                        "ASIO_Sub_Fund_2_Future_Security.xlsx"
                    )
                    excel_files.append((future_excel_file, future_excel_name))
                
                if not excel_files:
                    loader.close()
                    messagebox.showwarning("Warning", "No data to export.")
                    return
                
                # Create zip file
                zip_buffer = multiple_excels_to_zip(excel_files, "FNO_ASIO_SF2_TradeLoader.zip")
                
                # Save zip file
                with open(out_path, 'wb') as f:
                    f.write(zip_buffer.read())
                
                file_list = [name for _, name in excel_files]
                # Close loader + show success
                loader.close()
                messagebox.showinfo("Success", f"Template data exported to:\n{out_path}\n\nContains {len(file_list)} file(s):\n" + "\n".join([f"- {file}" for file in file_list]))
            except Exception as e:
                loader.close()
                messagebox.showerror("Error", f"Failed to export template data: {e}")
            finally:
                end_time = time.time()
        
        # Run heavy work in thread
        threading.Thread(target=task, daemon=True).start()

