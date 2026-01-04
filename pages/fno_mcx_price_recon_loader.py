import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import json
from datetime import datetime, timedelta
from decimal import Decimal
from CONSTANTS import *
import pandas as pd
import threading
from tkcalendar import DateEntry

from my_app.pages.loading import LoadingSpinner
from .helper import output_save_in_template, output_save_in_template_csv, multiple_files_to_zip, read_file


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


class FNOMCXPriceReconLoaderPage(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg="#ecf0f1")

        # Title
        title = tk.Label(self, text="📊 FNO and MCX Price Recon & Loader", font=("Arial", 20, "bold"), bg="#ecf0f1", fg="#2c3e50")
        title.pack(pady=10)

        # Controls Frame
        controls = tk.Frame(self, bg="#ecf0f1")
        controls.pack(fill="x", padx=20, pady=5)

        # LPA File (Local Position Appraisal)
        lpa_file_row = tk.Frame(controls, bg="#ecf0f1")
        lpa_file_row.pack(fill="x", pady=2)
        tk.Label(lpa_file_row, text="LPA File (Local Position Appraisal):", font=("Arial", 11), bg="#ecf0f1", fg="#2c3e50").pack(side="left")
        self.lpa_path_var = tk.StringVar()
        tk.Entry(lpa_file_row, textvariable=self.lpa_path_var, width=60).pack(side="left", padx=8)
        tk.Button(lpa_file_row, text="Browse", command=self._browse_lpa_file, bg="#3498db", fg="white", relief="flat", padx=10, pady=4).pack(side="left")

        # Holding Statement File
        holding_file_row = tk.Frame(controls, bg="#ecf0f1")
        holding_file_row.pack(fill="x", pady=2)
        tk.Label(holding_file_row, text="Holding Statement File:", font=("Arial", 11), bg="#ecf0f1", fg="#2c3e50").pack(side="left")
        self.holding_path_var = tk.StringVar()
        tk.Entry(holding_file_row, textvariable=self.holding_path_var, width=60).pack(side="left", padx=8)
        tk.Button(holding_file_row, text="Browse", command=self._browse_holding_file, bg="#3498db", fg="white", relief="flat", padx=10, pady=4).pack(side="left")

        # Segment Selection
        segment_row = tk.Frame(controls, bg="#ecf0f1")
        segment_row.pack(fill="x", pady=2)
        tk.Label(segment_row, text="Segment:", font=("Arial", 11), bg="#ecf0f1", fg="#2c3e50").pack(side="left")
        self.segment_var = tk.StringVar()
        self.segment_combo = ttk.Combobox(segment_row, textvariable=self.segment_var, width=20, state="readonly")
        self.segment_combo['values'] = ['FNO', 'MCX']
        self.segment_combo.current(0)  # Default to FNO
        self.segment_combo.pack(side="left", padx=8)

        # Price Date
        price_data_row = tk.Frame(controls, bg="#ecf0f1")
        price_data_row.pack(fill="x", pady=2)
        tk.Label(price_data_row, text="Price Date:", font=("Arial", 11), bg="#ecf0f1", fg="#2c3e50").pack(side="left")
        self.price_data_var = tk.StringVar()
        # Set default: if Monday, show previous Friday; otherwise show yesterday
        today = datetime.now().date()
        if today.weekday() == 0:  # Monday (0 = Monday)
            default_price_date = today - timedelta(days=3)  # Previous Friday
        else:
            default_price_date = today - timedelta(days=1)  # Yesterday
        self.price_data_entry = DateEntry(
            price_data_row,
            textvariable=self.price_data_var,
            date_pattern='dd/MM/yyyy',
            width=15,
            font=('Arial', 10)
        )
        self.price_data_entry.set_date(default_price_date)
        self.price_data_entry.pack(side="left", padx=8)
        
        # Exclude Expiry Date (independent field, defaults to same logic as price date)
        date_entry_row = tk.Frame(controls, bg="#ecf0f1")
        date_entry_row.pack(fill="x", pady=2)
        tk.Label(date_entry_row, text="Exclude Expiry Date:", font=("Arial", 11), bg="#ecf0f1", fg="#2c3e50").pack(side="left")
        self.date_var = tk.StringVar()
        # Set default: if Monday, show previous Friday; otherwise show yesterday
        if today.weekday() == 0:  # Monday (0 = Monday)
            exclude_date = today - timedelta(days=3)  # Previous Friday
        else:
            exclude_date = today - timedelta(days=1)  # Yesterday
        self.date_entry = DateEntry(
            date_entry_row,
            textvariable=self.date_var,
            date_pattern='dd/MM/yyyy',
            width=15,
            font=('Arial', 10)
        )
        self.date_entry.set_date(exclude_date)
        self.date_entry.pack(side="left", padx=8)

        # Format selection row
        format_row = tk.Frame(controls, bg="#ecf0f1")
        format_row.pack(fill="x", pady=2)
        tk.Label(format_row, text="Export Format:", font=("Arial", 10), bg="#ecf0f1", fg="#2c3e50").pack(side="left")
        self.csv_format_var = tk.BooleanVar(value=True)  # Default to CSV
        self.xlsx_format_var = tk.BooleanVar(value=False)
        tk.Checkbutton(format_row, text="CSV", variable=self.csv_format_var, font=("Arial", 10), bg="#ecf0f1", fg="#2c3e50", selectcolor="#ecf0f1").pack(side="left", padx=(8, 4))
        tk.Checkbutton(format_row, text="XLSX", variable=self.xlsx_format_var, font=("Arial", 10), bg="#ecf0f1", fg="#2c3e50", selectcolor="#ecf0f1").pack(side="left", padx=4)

        # Buttons row
        buttons_row = tk.Frame(controls, bg="#ecf0f1")
        buttons_row.pack(fill="x", pady=5)
        tk.Button(buttons_row, text="Process", command=self._process, bg="#27ae60", fg="white", relief="flat", padx=14, pady=6, font=("Arial", 11, "bold")).pack(side="left", padx=10)
        tk.Button(buttons_row, text="Export Excel", command=self._export_excel, bg="#8e44ad", fg="white", relief="flat", padx=14, pady=6, font=("Arial", 11, "bold")).pack(side="left")
        tk.Button(buttons_row, text="Export to Template", command=self._export_to_template, bg="#d35400", fg="white", relief="flat", padx=14, pady=6, font=("Arial", 11, "bold")).pack(side="left", padx=10)

        # Status
        self.status_var = tk.StringVar(value="Load LPA file and Holding Statement file, then click Process")
        self.status_label = tk.Label(self, textvariable=self.status_var, font=("Arial", 10), bg="#f8f9fa", fg="#6c757d", anchor="w", relief="flat", bd=0, padx=15, pady=10)
        self.status_label.pack(fill="x", padx=20, pady=(0, 10))

        # Table view
        content = tk.Frame(self, bg="#ecf0f1")
        content.pack(fill="both", expand=True, padx=20, pady=10)

        # Table header with search
        header = tk.Frame(content, bg="#ecf0f1")
        header.pack(fill="x", pady=(0, 4))
        tk.Label(header, text="Price Reconciliation Data", font=("Arial", 13, "bold"), bg="#ecf0f1", fg="#2c3e50").pack(side="left")
        
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
            "Security", "LPA_Quantity", "Holding_Quantity", "Quantity_Difference"
        )
        self.tree = ttk.Treeview(table_frame, columns=self.table_columns, show="headings", height=12)
        
        # Column widths
        column_widths = {
            "Security": 200,
            "LPA_Quantity": 120,
            "Holding_Quantity": 120,
            "Quantity_Difference": 150
        }
        for col in self.table_columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=column_widths.get(col, 120), anchor="w")
        
        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        
        # Grid layout
        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, columnspan=2, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        # Data holders
        self.table_rows = []
        self.lpa_data = None
        self.holding_data = None
        self.holding_data_raw = None  # Store original data before excluding expiry date
        self.processed_data = []
        self._template_data = []
        self.selected_segment = None  # Store selected segment for template export

    # ---- UI handlers ----
    def _browse_lpa_file(self):
        """Browse for LPA file (CSV/XLS/XLSX)."""
        path = filedialog.askopenfilename(
            title="Select LPA File (Local Position Appraisal)", 
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
            self.lpa_path_var.set(path)

    def _browse_holding_file(self):
        """Browse for Holding Statement file (CSV/XLS/XLSX)."""
        path = filedialog.askopenfilename(
            title="Select Holding Statement File", 
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
            self.holding_path_var.set(path)
    
    

    # ---- Core processing ----
    def _process(self):
        """Process the LPA and Holding Statement files."""
        # Clear table
        for it in self.tree.get_children():
            self.tree.delete(it)
        self.table_rows = []
        self.processed_data = []
        self._template_data = []

        # Validate files
        lpa_path = self.lpa_path_var.get().strip()
        holding_path = self.holding_path_var.get().strip()

        if not lpa_path or not os.path.exists(lpa_path):
            messagebox.showwarning("File Missing", "Please select a valid LPA file (CSV/XLS/XLSX).")
            return

        if not holding_path or not os.path.exists(holding_path):
            messagebox.showwarning("File Missing", "Please select a valid Holding Statement file (CSV/XLS/XLSX).")
            return

        # Load consolidated JSON
        from my_app.file_utils import get_app_directory
        app_dir = get_app_directory()
        consolidated_path = os.path.join(app_dir, "consolidated_data.json")

        try:
            consolidated_data = {}
            if os.path.exists(consolidated_path):
                with open(consolidated_path, "r") as f:
                    consolidated_data = json.load(f)
            
            # Extract configuration from consolidated data if needed
            # Add specific configurations here based on requirements
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load consolidated_data.json: {e}\nPlease ensure this file exists under my_app/.")
            return

        try:
            # Load LPA file
            self.status_var.set("Loading LPA file...")
            self.lpa_data = read_file(
                file_path=lpa_path,
                sheet_name=0,
                start_row=0,
                header=True,
                skip_blank_rows=False
            )
            # Clean headers
            if self.lpa_data is not None and not self.lpa_data.empty:
                self.lpa_data.columns = self.lpa_data.columns.str.strip()

            # Load Holding Statement file
            self.status_var.set("Loading Holding Statement file...")
            # Read from row 12 (0-based index 11) and skip first column
            self.holding_data_raw = read_file(
                file_path=holding_path,
                sheet_name=0,
                start_row=11,  # Row 12 becomes header (0-based index)
                header=True,
                skip_blank_rows=False
            )
            # Clean headers and skip first column (column A)
            if self.holding_data_raw is not None and not self.holding_data_raw.empty:
                # Keep only B → end (skip column A)
                self.holding_data_raw = self.holding_data_raw.iloc[:, 1:]
                self.holding_data_raw.columns = self.holding_data_raw.columns.str.strip()
                
                # Exclude expiry date based on selected date from DateEntry
                # Use fixed ExpiryDate column
                expiry_field = 'ExpiryDate'
                
                if expiry_field in self.holding_data_raw.columns:
                    # Get selected date from DateEntry
                    try:
                        selected_date = self.date_entry.get_date() if hasattr(self, 'date_entry') else None
                        if selected_date:
                            # Convert selected date to match the format in the data
                            selected_date_str = selected_date.strftime('%d/%m/%Y')
                            
                            # Filter out rows where expiry date matches the selected date
                            # Try different date formats
                            date_formats = ['%d/%m/%Y', '%d-%m-%Y', '%Y-%m-%d', '%d/%m/%y', '%d-%m-%y']
                            
                            self.holding_data = self.holding_data_raw.copy()
                            for date_format in date_formats:
                                try:
                                    formatted_date = selected_date.strftime(date_format)
                                    # Remove rows where expiry date matches
                                    self.holding_data = self.holding_data[
                                        self.holding_data[expiry_field].astype(str).str.strip() != formatted_date
                                    ]
                                    break
                                except:
                                    continue
                            
                            self.status_var.set(f"Excluded rows with expiry date '{selected_date_str}' from Holding Statement")
                        else:
                            # If no date selected, just exclude the expiry date column
                            self.holding_data = self.holding_data_raw.drop(columns=[expiry_field])
                            self.status_var.set(f"Excluded expiry date column '{expiry_field}' from Holding Statement")
                    except Exception as e:
                        # Fallback: just exclude the expiry date column
                        self.holding_data = self.holding_data_raw.drop(columns=[expiry_field])
                        self.status_var.set(f"Excluded expiry date column '{expiry_field}' from Holding Statement")
                else:
                    # No ExpiryDate column found, keep all data
                    self.holding_data = self.holding_data_raw.copy()
                    self.status_var.set("No ExpiryDate column found in Holding Statement")

            if self.lpa_data is None or self.lpa_data.empty:
                messagebox.showerror("Error", "LPA file is empty or could not be loaded.")
                return

            if self.holding_data is None or self.holding_data.empty:
                messagebox.showerror("Error", "Holding Statement file is empty or could not be loaded.")
                return

            # Get selected segment
            selected_segment = self.segment_var.get().strip() if hasattr(self, 'segment_var') else 'FNO'
            if not selected_segment:
                messagebox.showwarning("Segment Missing", "Please select a segment (FNO or MCX).")
                return
            
            # Process data - placeholder function, logic to be implemented
            self.status_var.set(f"Processing data for {selected_segment}...")
            self.selected_segment = selected_segment  # Store segment for template export
            self.table_rows, self.processed_data, self._template_data = self._process_data(
                self.lpa_data, self.holding_data, consolidated_data, selected_segment
            )

            # Render table
            self._render_table()
            self.status_var.set(f"Processed {len(self.table_rows)} records successfully")
            self.status_label.config(fg="#6c757d")  # Reset to default gray color

        except Exception as e:
            messagebox.showerror("Error", f"Failed to process files: {e}")
            self.status_var.set("Processing failed")
            import traceback
            traceback.print_exc()

    def _process_data(self, lpa_df, holding_df, consolidated_data, segment='FNO'):
        """
        Process LPA and Holding Statement data to create price reconciliation.
        
        Args:
            lpa_df: DataFrame from LPA file
            holding_df: DataFrame from Holding Statement file (expiry date excluded)
            consolidated_data: Configuration data from consolidated_data.json
            segment: Selected segment ('FNO' or 'MCX')
        
        Returns:
            tuple: (table_rows, processed_data, template_data)
        """
        table_rows = []
        processed_data = []
        template_data = []
        
        try:
            # Extract pricing data from consolidated_data
            asio_pricing_fno = consolidated_data.get("asio_pricing_fno", {})
            asio_pricing_mcx = consolidated_data.get("asio_pricing_mcx", {})
            
            # Step 1: Process LPA DataFrame
            # Extract: Invest, Quantity, Group2
            lpa_columns_needed = ['Invest', 'Quantity', 'Group2']
            
            # Check if required columns exist in LPA
            missing_cols = [col for col in lpa_columns_needed if col not in lpa_df.columns]
            if missing_cols:
                raise ValueError(f"LPA file missing required columns: {missing_cols}")
            
            # Filter LPA data based on selected segment
            # Read filters dynamically from consolidated_data.json (configured via Data Config)
            mcx_group2_filters = consolidated_data.get("mcx_group2_filters")
            fno_group2_filters = consolidated_data.get("fno_group2_filters")
            
            # Validate filters exist and are lists
            if not isinstance(mcx_group2_filters, list):
                raise ValueError(
                    "MCX Group2 filters not configured. Please configure 'mcx_group2_filters' "
                    "in Data Config (consolidated_data.json). Expected format: list of filter strings."
                )
            if not isinstance(fno_group2_filters, list):
                raise ValueError(
                    "FNO Group2 filters not configured. Please configure 'fno_group2_filters' "
                    "in Data Config (consolidated_data.json). Expected format: list of filter strings."
                )
            
            # Filter based on selected segment (don't combine - process only selected segment)
            if segment == 'MCX':
                lpa_filtered = lpa_df[lpa_df['Group2'].isin(mcx_group2_filters)].copy()
                lpa_filtered['Type'] = 'MCX'
                exchange_prefix = 'MCX'
            elif segment == 'FNO':
                lpa_filtered = lpa_df[lpa_df['Group2'].isin(fno_group2_filters)].copy()
                lpa_filtered['Type'] = 'FNO'
                exchange_prefix = 'NSE'
            else:
                raise ValueError(f"Invalid segment: {segment}. Must be 'FNO' or 'MCX'")
            
            if lpa_filtered.empty:
                self.status_var.set(f"No matching {segment} data found in LPA file")
                return table_rows, processed_data, template_data
            
            # Step 2: Process Holding Statement DataFrame
            # First, exclude expiry date rows based on selected date from frontend using ExpiryDate column
            # Manual exclusion without pandas-level operations
            holding_processed = holding_df.copy()
            
            # Check if ExpiryDate column exists
            if 'ExpiryDate' not in holding_processed.columns:
                raise ValueError("Holding Statement missing required column: ExpiryDate")
            
            # Get selected expiry date from DateEntry and exclude matching rows
            exclude_date = self.date_entry.get_date() if hasattr(self, 'date_entry') else None
            if exclude_date:
                exclude_date_str = exclude_date.strftime('%d-%m-%Y')  # Format: 25-11-2025
                rows_to_keep = []
                excluded_count = 0
                
                for idx, row in holding_processed.iterrows():
                    expiry_date_val = str(row['ExpiryDate']).strip()
                    if expiry_date_val != exclude_date_str:
                        rows_to_keep.append(idx)
                    else:
                        excluded_count += 1
                
                holding_processed = holding_processed.loc[rows_to_keep].copy()
                if excluded_count > 0:
                    self.status_var.set(f"Excluded {excluded_count} rows with expiry date {exclude_date_str}")
            
            # Step 3: Create concatenated security codes for Holding Statement
            # Required columns: NetBuy, NetSell, UnderlyingCode, ExpiryDate, OptionType, StrikePrice, ContractSettlementPrice
            holding_cols_needed = ['NetBuy', 'NetSell', 'UnderlyingCode', 'ExpiryDate', 'OptionType', 'StrikePrice', 'ContractSettlementPrice']
            
            # Check if required columns exist
            missing_holding_cols = [col for col in holding_cols_needed if col not in holding_processed.columns]
            if missing_holding_cols:
                raise ValueError(f"Holding Statement missing required columns: {missing_holding_cols}")
            
            # Create concatenated security code for each row
            def create_security_code(row, exchange_prefix):
                """Create security code: EXCHANGE + UnderlyingCode + YYYYMMDD + OptionType[0] + StrikePrice"""
                try:
                    underlying = str(row['UnderlyingCode']).strip()
                    
                    # Parse expiry date
                    expiry_date = row['ExpiryDate']
                    if pd.isna(expiry_date) or expiry_date == '':
                        return ''
                    
                    # Try different date formats
                    expiry_formatted = ''
                    date_formats = [
                        ('%d/%m/%Y', '%Y%m%d'),
                        ('%d-%m-%Y', '%Y%m%d'),
                        ('%Y-%m-%d', '%Y%m%d'),
                        ('%d/%m/%y', '%Y%m%d'),
                        ('%d-%m-%y', '%Y%m%d'),
                        ('%Y%m%d', '%Y%m%d'),  # Already in YYYYMMDD format
                    ]
                    
                    expiry_str = str(expiry_date).strip()
                    for read_fmt, write_fmt in date_formats:
                        try:
                            if read_fmt == '%Y%m%d' and len(expiry_str) == 8 and expiry_str.isdigit():
                                expiry_formatted = expiry_str
                                break
                            else:
                                date_obj = datetime.strptime(expiry_str, read_fmt)
                                expiry_formatted = date_obj.strftime(write_fmt)
                                break
                        except:
                            continue
                    
                    if not expiry_formatted:
                        return ''
                    
                    # Get option type first character
                    option_type = str(row['OptionType']).strip()
                    option_first_char = option_type[0] if option_type else ''
                    
                    # Get strike price
                    strike_price = str(row['StrikePrice']).strip()
                    try:
                        strike_price_int = int(float(strike_price))
                    except:
                        strike_price_int = 0
                    
                    # Create security code
                    security_code = f"{exchange_prefix}{underlying}{expiry_formatted}{option_first_char}{strike_price_int}"
                    return security_code
                except Exception as e:
                    return ''
            
            # Create dictionary: Key = SecurityCode, Value = net_qty (NetBuy - NetSell)
            holding_dict = {}
            # Create dictionary: Key = SecurityCode, Value = ContractSettlementPrice
            holding_price_dict = {}
            for _, row in holding_processed.iterrows():
                net_buy= row['NetBuy']
                net_sell=row['NetSell']
                option_type=row['OptionType']
                underlying_code=row['UnderlyingCode']
                strike_price=row['StrikePrice']
                expiry_date=row['ExpiryDate']
                contract_settlement_price = row['ContractSettlementPrice']

                security_code = create_security_code(row, exchange_prefix)
                if not security_code:
                    continue
                net_qty = int(net_buy) - int(net_sell)

                if security_code in holding_dict:
                    holding_dict[security_code] += net_qty
                else:
                    holding_dict[security_code] = net_qty
                
                # if option_type == "FF":
                #     breakpoint()
                
                # Store ContractSettlementPrice (use the last value if multiple rows have same security_code)
                # Convert to Decimal for exact precision
                holding_price_dict[security_code] = _safe_decimal(contract_settlement_price)

            # Match with holding statement using the dictionary
            for _, lpa_row in lpa_filtered.iterrows():
                invest_code = str(lpa_row['Invest']).strip()
                lpa_qty = int(lpa_row['Quantity'])
                lpa_type = lpa_row['Type']
                
                # Lookup net quantity from holding dictionary
                holding_qty = holding_dict.get(invest_code, 0)
                
                # Calculate quantity difference (LPA - Holding)
                qty_difference = lpa_qty - holding_qty
                
                # Create table row: Security, LPA_Quantity, Holding_Quantity, Quantity_Difference
                table_row = [
                    invest_code,      # Security
                    lpa_qty,          # LPA_Quantity
                    holding_qty,      # Holding_Quantity
                    qty_difference    # Quantity_Difference (LPA - Holding)
                ]
                table_rows.append(table_row)
                
                # Store processed data
                processed_data.append({
                    'Security': invest_code,
                    'LPA_Quantity': lpa_qty,
                    'Holding_Quantity': holding_qty,
                    'Quantity_Difference': qty_difference,
                    'Type': lpa_type
                })
                
                # if invest_code == "NSENIFTY20251230F0":
                #     breakpoint()
                
                # Build template_data row with pricing headers
                row = []
                # Access pricing data based on segment
                if segment == 'FNO':
                    pricing_data = asio_pricing_fno
                    for header in FNO_PRICING_HEADER:
                        if header in pricing_data:
                            row.append(pricing_data[header])
                        else:
                            if header == PRICEDATE:
                                # Get price date from frontend DateEntry
                                price_date = self.price_data_entry.get_date() if hasattr(self, 'price_data_entry') else None
                                if price_date:
                                    # Format date as MM-DD-YYYY
                                    price_date_str = price_date.strftime('%m-%d-%Y')
                                    row.append(price_date_str)
                                else:
                                    row.append('')
                            elif header == INVESTMENT:
                                # print("invest_code", invest_code)
                                row.append(invest_code)
                            elif header == PRICE:
                                # Use ContractSettlementPrice from holding_price_dict
                                contract_price = holding_price_dict.get(invest_code, '')
                                row.append(contract_price)
                            else:
                                row.append('')
                
                elif segment == 'MCX':
                    pricing_data = asio_pricing_mcx
                    for header in MCX_PRICING_HEADER:
                        if header in pricing_data:
                            row.append(pricing_data[header])
                        else:
                            if header == PRICEDATE:
                                # Get price date from frontend DateEntry
                                price_date = self.price_data_entry.get_date() if hasattr(self, 'price_data_entry') else None
                                if price_date:
                                    # Format date as MM-DD-YYYY
                                    price_date_str = price_date.strftime('%m-%d-%Y')
                                    row.append(price_date_str)
                                else:
                                    row.append('')
                            elif header == INVESTMENT:
                                row.append(invest_code)
                            elif header == PRICE:
                                # Use ContractSettlementPrice from holding_price_dict
                                contract_price = holding_price_dict.get(invest_code, '')
                                row.append(contract_price)
                            else:
                                row.append('')
                
                template_data.append(row)
            
            self.status_var.set(f"Processed {len(table_rows)} records for {segment}")
            
        except Exception as e:
            import traceback
            error_msg = f"Error in _process_data: {str(e)}\n{traceback.format_exc()}"
            messagebox.showerror("Processing Error", error_msg)
            self.status_var.set(f"Processing error: {str(e)}")
        
        return table_rows, processed_data, template_data

    # ---- Rendering with filtering ----
    def _render_table(self):
        """Render table with search filtering."""
        # Clear table
        for it in self.tree.get_children():
            self.tree.delete(it)
        
        term = (self.search_var.get() if hasattr(self, 'search_var') else "").lower().strip()
        
        for r in self.table_rows:
            if not term or term in " ".join(map(str, r)).lower():
                self.tree.insert("", "end", values=r)

    def _export_excel(self):
        """Export data to Excel and/or CSV file based on format selection."""
        if not self.table_rows and not self.processed_data:
            messagebox.showinfo("Nothing to export", "Process files first.")
            return

        export_csv = self.csv_format_var.get()
        export_xlsx = self.xlsx_format_var.get()
        
        if not export_csv and not export_xlsx:
            messagebox.showwarning("No Format Selected", "Please select at least one export format (CSV or XLSX).")
            return

        exported_files = []

        try:
            if self.table_rows:
                df_table = pd.DataFrame(self.table_rows, columns=self.table_columns)
            else:
                df_table = pd.DataFrame()

            # Export CSV first (default format)
            if export_csv:
                out_path = filedialog.asksaveasfilename(
                    title="Save CSV", defaultextension=".csv",
                    filetypes=[["CSV Files", "*.csv"]],
                    initialfile="FNO_MCX_Price_Recon_Output.csv"
                )
                if out_path:
                    df_table_csv = df_table.copy()
                    for col in df_table_csv.columns:
                        if df_table_csv[col].dtype == 'object' and len(df_table_csv) > 0:
                            non_null_vals = df_table_csv[col].dropna()
                            if not non_null_vals.empty and isinstance(non_null_vals.iloc[0], Decimal):
                                df_table_csv[col] = df_table_csv[col].apply(lambda x: str(x) if pd.notna(x) else '')
                    df_table_csv.to_csv(out_path, index=False, encoding='utf-8-sig')
                    exported_files.append(f"CSV: {out_path}")

            # Export XLSX
            if export_xlsx:
                out_path = filedialog.asksaveasfilename(
                    title="Save Excel", defaultextension=".xlsx",
                    filetypes=[["Excel", "*.xlsx"]],
                    initialfile="FNO_MCX_Price_Recon_Output.xlsx"
                )
                if out_path:
                    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                        # Sheet 1: Processed Data
                        if not df_table.empty:
                            df_table.to_excel(writer, sheet_name="Price_Reconciliation", index=False)
                        
                        # Sheet 2: Original LPA Data
                        if self.lpa_data is not None and not self.lpa_data.empty:
                            self.lpa_data.to_excel(writer, sheet_name="LPA_Original", index=False)
                        
                        # Sheet 3: Original Holding Data (without expiry date)
                        if self.holding_data is not None and not self.holding_data.empty:
                            self.holding_data.to_excel(writer, sheet_name="Holding_Original", index=False)
                    exported_files.append(f"XLSX: {out_path}")

            if exported_files:
                messagebox.showinfo("Success", f"Files exported successfully:\n\n" + "\n".join(exported_files))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export: {e}")

    def _export_to_template(self):
        """Export data to template format (ZIP with Excel/CSV files)."""
        if not self._template_data:
            messagebox.showinfo("Nothing to export", "Process files first to generate template data.")
            return

        # Check format selection
        export_csv = self.csv_format_var.get()
        export_xlsx = self.xlsx_format_var.get()
        
        if not export_csv and not export_xlsx:
            messagebox.showwarning("No Format Selected", "Please select at least one export format (CSV or XLSX).")
            return

        # Determine zip filename based on segment
        if self.selected_segment == 'FNO':
            zip_filename = "FNO_Price_Recon_Template.zip"
        elif self.selected_segment == 'MCX':
            zip_filename = "MCX_Price_Recon_Template.zip"
        else:
            zip_filename = "Price_Recon_Template.zip"
        
        # Ask output path for zip file
        out_path = filedialog.asksaveasfilename(
            title="Save Template Data",
            defaultextension=".zip",
            filetypes=[["ZIP Files", "*.zip"]],
            initialfile=zip_filename
        )
        if not out_path:
            return
        
        # Show spinner (non-blocking)
        loader = LoadingSpinner(self, text="Exporting templates...")

        def task():
            try:
                files = []
                
                # Get the appropriate header based on segment
                if not self.selected_segment:
                    loader.close()
                    messagebox.showwarning("Warning", "No segment selected. Process files first.")
                    return
                
                # Select header based on segment
                if self.selected_segment == 'FNO':
                    pricing_headers = FNO_PRICING_HEADER
                    template_filename_base = "FNO_Pricing_Template"
                    zip_filename = "FNO_Price_Recon_Template.zip"
                elif self.selected_segment == 'MCX':
                    pricing_headers = MCX_PRICING_HEADER
                    template_filename_base = "MCX_Pricing_Template"
                    zip_filename = "MCX_Price_Recon_Template.zip"
                else:
                    loader.close()
                    messagebox.showwarning("Warning", f"Invalid segment: {self.selected_segment}")
                    return
                
                # Convert template_data (list of lists) to list of dicts
                template_data_dicts = []
                for row in self._template_data:
                    row_dict = {}
                    for idx, header in enumerate(pricing_headers):
                        if idx < len(row):
                            row_dict[header] = row[idx]
                        else:
                            row_dict[header] = ''
                    template_data_dicts.append(row_dict)
                
                # Export CSV if selected
                if export_csv and template_data_dicts:
                    file_io, file_name = output_save_in_template_csv(
                        template_data_dicts,
                        pricing_headers,
                        f"{template_filename_base}.csv"
                    )
                    files.append((file_io, file_name))
                
                # Export XLSX if selected
                if export_xlsx and template_data_dicts:
                    file_io, file_name = output_save_in_template(
                        template_data_dicts,
                        pricing_headers,
                        f"{template_filename_base}.xlsx"
                    )
                    files.append((file_io, file_name))
                
                if not files:
                    loader.close()
                    messagebox.showwarning("Warning", "No template data to export.")
                    return
                
                # Create zip file with segment-specific name
                zip_buffer = multiple_files_to_zip(files, zip_filename)
                
                # Save zip file
                with open(out_path, 'wb') as f:
                    f.write(zip_buffer.read())
                
                file_list = [name for _, name in files]
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
        
        # Run heavy work in thread
        threading.Thread(target=task, daemon=True).start()
    
    def _show_email_dialog(self, zip_path, file_list):
        """Show email dialog for sending files via Outlook."""
        try:
            from .email_dialog import EmailDialog
            EmailDialog(self, zip_path, file_list)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open email dialog: {e}")
            import traceback
            traceback.print_exc()