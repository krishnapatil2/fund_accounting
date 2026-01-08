import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import json

# LAZY IMPORTS - Heavy libraries imported only when needed (in methods)
# This speeds up frame opening significantly
# pandas, openpyxl will be imported in _process() method when actually needed


class GTNLoaderPage(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg="#ecf0f1")

        # Title
        title = tk.Label(self, text="📑 GTN Loader", font=("Arial", 20, "bold"), bg="#ecf0f1", fg="#2c3e50")
        title.pack(pady=10)

        # Controls
        controls = tk.Frame(self, bg="#ecf0f1")
        controls.pack(fill="x", padx=20, pady=5)

        self.file_paths = []  # List to store multiple file paths

        # File row
        file_row = tk.Frame(controls, bg="#ecf0f1")
        file_row.pack(fill="x", pady=2)
        tk.Label(file_row, text="Files:", font=("Arial", 11), bg="#ecf0f1", fg="#2c3e50").pack(side="left")
        tk.Button(file_row, text="Browse Files", command=self._browse_files, bg="#3498db", fg="white", relief="flat", padx=10, pady=4).pack(side="left", padx=(0, 5))
        tk.Button(file_row, text="Clear", command=self._clear_files, bg="#e74c3c", fg="white", relief="flat", padx=10, pady=4).pack(side="left")
        
        # Files list display
        files_frame = tk.Frame(controls, bg="#ecf0f1")
        files_frame.pack(fill="both", expand=True, pady=(5, 0))
        tk.Label(files_frame, text="Selected Files:", font=("Arial", 10, "bold"), bg="#ecf0f1", fg="#2c3e50").pack(anchor="w")
        files_list_frame = tk.Frame(files_frame, bg="white", relief="solid", bd=1)
        files_list_frame.pack(fill="both", expand=True, pady=(2, 0))
        
        # Listbox with scrollbar for selected files
        files_scroll = tk.Scrollbar(files_list_frame)
        files_scroll.pack(side="right", fill="y")
        self.files_listbox = tk.Listbox(files_list_frame, yscrollcommand=files_scroll.set, height=4, font=("Arial", 9))
        self.files_listbox.pack(side="left", fill="both", expand=True)
        files_scroll.config(command=self.files_listbox.yview)

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
        self.status_var = tk.StringVar(value="Load a file (CSV/XLS/XLSX) and click Process")
        tk.Label(self, textvariable=self.status_var, font=("Arial", 10), bg="#ecf0f1", fg="#7f8c8d").pack(fill="x", padx=20)

        # Content frame
        content = tk.Frame(self, bg="#ecf0f1")
        content.pack(fill="both", expand=True, padx=20, pady=10)

        # Table header with search
        header = tk.Frame(content, bg="#ecf0f1")
        header.pack(fill="x", pady=(0, 4))
        tk.Label(header, text="GTN Process Data", font=("Arial", 13, "bold"), bg="#ecf0f1", fg="#2c3e50").pack(side="left")
        
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
        
        # Table columns (will be populated based on data)
        self.table_columns = ()
        self.tree = ttk.Treeview(table_frame, columns=self.table_columns, show="headings", height=12)
        
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
        self.processed_data = []
        self._template_data = []
        self.processed_files_data = {}  # Dictionary to store processed data per file: {file_path: {processed_data, template_data, config}}

    # ---- UI handlers ----
    def _browse_files(self):
        """Browse for multiple CSV, XLS, or XLSX files."""
        paths = filedialog.askopenfilenames(
            title="Select Files", 
            filetypes=[
                ["All Supported Files", "*.csv *.xls *.xlsx"],
                ["CSV Files", "*.csv"],
                ["Excel Files", "*.xls *.xlsx"],
                ["XLS Files", "*.xls"],
                ["XLSX Files", "*.xlsx"],
                ["All Files", "*.*"]
            ]
        )
        if paths:
            # Add new files to the list (avoid duplicates)
            for path in paths:
                if path not in self.file_paths:
                    self.file_paths.append(path)
            self._update_files_display()
    
    def _clear_files(self):
        """Clear all selected files."""
        self.file_paths = []
        self._update_files_display()
        # Clear table and data
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.table_rows = []
        self.processed_data = []
        self._template_data = []
        self.processed_files_data = {}
        self.status_var.set("Load file(s) (CSV/XLS/XLSX) and click Process")
    
    def _update_files_display(self):
        """Update the files listbox display."""
        self.files_listbox.delete(0, tk.END)
        for path in self.file_paths:
            # Show just the filename
            filename = os.path.basename(path)
            self.files_listbox.insert(tk.END, filename)
        if self.file_paths:
            self.status_var.set(f"{len(self.file_paths)} file(s) selected. Click Process to generate loaders.")
        else:
            self.status_var.set("Load file(s) (CSV/XLS/XLSX) and click Process")

    @staticmethod
    def month_to_number(month_name: str) -> str:
        """Convert month name to 2-digit number string."""
        months = [
            "January","February","March","April","May","June",
            "July","August","September","October","November","December"
        ]
        try:
            return f"{months.index(month_name) + 1:02d}"
        except ValueError:
            return "01"  # Default to January if not found

    @staticmethod
    def parse_foreign_isin(code: str, gtn_sp_30_call_option: dict, gtn_sp_30_put_option: dict) -> str:
        """
        Parse foreign ISIN code format.
        Example: IBIT\26R18\69.0 -> IBIT 260618P00069000
        Format: {symbol} {YY}{MM}{DD}{C|P}000{strike_5digits}
        """
        if not code or str(code).strip() == "" or str(code).lower() in ["nan", "none", "null"]:
            return ""
        
        try:
            parts = str(code).split("\\")
            if len(parts) < 3:
                return str(code)  # Return original if format is wrong
            
            symbol, date_part, strike_part = parts[0], parts[1], parts[2]

            year = date_part[:2]         # YY
            month_letter = date_part[2]  # month code
            day = date_part[3:]          # DD

            # Determine call / put
            option_type = "C" if month_letter in gtn_sp_30_call_option else "P"

            if option_type == "C":
                month_name = gtn_sp_30_call_option.get(month_letter, "January")
            else:
                month_name = gtn_sp_30_put_option.get(month_letter, "January")

            month_number = GTNLoaderPage.month_to_number(month_name)

            # strike → extract integer part and format as 5 digits (multiply by 1000)
            # Example: 69.0 -> 69 -> 69*1000 = 69000
            try:
                # Extract integer part from strike (handle both "69.0" and "69")
                strike_str = str(strike_part).split(".")[0]
                strike_int = int(strike_str)
                # Multiply by 1000 to get the strike in the correct format
                # Example: 69 -> 69000, 150 -> 150000
                strike_value = strike_int * 1000
                # Convert to string (will be at least 4 digits for strikes >= 1)
                strike_formatted = str(strike_value)
            except (ValueError, IndexError):
                strike_formatted = "00000"

            return f"{symbol} {year}{month_number}{day}{option_type}000{strike_formatted}"
        except Exception as e:
            # Return original code if parsing fails
            return str(code)
    # ---- Core processing ----
    def _process(self):
        """Process all selected files (CSV/XLS/XLSX) and generate loaders."""
        # Lazy import heavy libraries only when processing (speeds up frame opening)
        from .helper import read_file
        
        if not self.file_paths:
            messagebox.showwarning("No Files", "Please select at least one file (CSV/XLS/XLSX).")
            return

        # Load consolidated JSON
        from my_app.file_utils import get_app_directory
        app_dir = get_app_directory()
        consolidated_path = os.path.join(app_dir, "consolidated_data.json")

        try:
            if os.path.exists(consolidated_path):
                with open(consolidated_path, "r") as f:
                    consolidated_data = json.load(f)
            
            # Extract GTN_LOADER configuration from consolidated data
            gtn_loader_config = consolidated_data.get("GTN_LOADER", {})
            gtn_sp_30_call_option = consolidated_data.get("gtn_sp_30_call_option", {})
            gtn_sp_30_put_option = consolidated_data.get("gtn_sp_30_put_option", {})
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load consolidated_data.json: {e}\nPlease ensure this file exists under my_app/.")
            return

        # Clear previous processed data
        self.processed_files_data = {}
        all_processed_rows = []

        # Process each file
        total_files = len(self.file_paths)
        successful_files = 0
        failed_files = []

        for file_idx, file_path in enumerate(self.file_paths, 1):
            if not os.path.exists(file_path):
                failed_files.append(f"{os.path.basename(file_path)} (file not found)")
                continue

            # Update status
            self.status_var.set(f"Processing file {file_idx}/{total_files}: {os.path.basename(file_path)}...")
            self.update()  # Update UI to show progress
            
            try:
                # Import pandas for data handling
                import pandas as pd
                
                # Read the file using helper function (supports CSV, XLS, XLSX)
                # Default: read from row 0, use first row as header
                df_data = read_file(
                    file_path=file_path,
                    sheet_name=0,   # First sheet for Excel files
                    start_row=0,    # Start from first row (0-based)
                    header=True,    # Use first row as column names
                    skip_blank_rows=True
                )
                
                # Clean headers (strip whitespace)
                df_data.columns = df_data.columns.str.strip()
                
                # Process each row and create GTN loader data
                processed_rows = []

                def get_side(text):
                    return text.split(":")[1].strip()

                def format_date(value):
                    # Try DD-MM-YYYY first
                    date = pd.to_datetime(value, errors="coerce", dayfirst=True)

                    if pd.isna(date):
                        # fallback (let pandas guess)
                        date = pd.to_datetime(value, errors="coerce")

                    if pd.isna(date):
                        return ""

                    return date.strftime("%m-%d-%Y")
                
                row_errors = []  # Track errors per row
                for index, row in df_data.iterrows():
                    try:
                        # Get fee amounts, defaulting to 0 if missing
                        B2B_COMM = float(row.get("B2B_COMM", 0) or 0)
                        BROKER_COMM = float(row.get("BROKER_COMM", 0) or 0)
                        OTHER_FEE_AMOUNT = float(row.get("OTHER_FEE_AMOUNT", 0) or 0)
                        VAT_AMOUNT = float(row.get("VAT_AMOUNT", 0) or 0)
                        WHT_AMOUNT = float(row.get("WHT_AMOUNT", 0) or 0)

                        # Calculate non-cap expenses sum
                        noncap_sum = B2B_COMM + BROKER_COMM + OTHER_FEE_AMOUNT + VAT_AMOUNT + WHT_AMOUNT

                        # Start with GTN loader config template
                        row_data = gtn_loader_config.copy()

                        # Parse ISIN code
                        isin_code = row.get("ISINCODE", "")
                        parsed_isin = self.parse_foreign_isin(isin_code, gtn_sp_30_call_option, gtn_sp_30_put_option)
                        
                        ##############################################################
                        # Determine if this is an option trade
                        # Check if parsed ISIN contains 'C' or 'P' (option type indicator)
                        # Format: {symbol} {YY}{MM}{DD}{C|P}000{strike}
                        # So we check if 'C' or 'P' appears in the parsed ISIN
                        # is_option = bool(parsed_isin and ('C' in parsed_isin or 'P' in parsed_isin))
                        # Get UNIT value and apply option logic
                        # unit_value = row.get("UNIT", "")
                        # if is_option and unit_value:
                        #     try:
                        #         # For option trades, multiply UNIT by 100
                        #         unit_float = float(unit_value)
                        #         quantity_value = str(int(unit_float * 100))
                        #     except (ValueError, TypeError):
                        #         # If conversion fails, use original value
                        #         quantity_value = str(unit_value)
                        # else:
                        #     # For non-option trades, use UNIT as is
                        #     quantity_value = str(unit_value)
                        ######################################################

                        trade_date = format_date(row.get("TRADE_DATE"))
                    
                        # Update row data with values from file
                        row_data.update({
                            "RecordType": str(get_side(row.get("SIDE", ""))),
                            "KeyValue": parsed_isin,
                            "Investment": parsed_isin,
                            "EventDate": trade_date,
                            "SettleDate": trade_date,
                            "ActualSettleDate": trade_date,
                            "Quantity": str(row.get("UNIT", "")),
                            "Price": str(row.get("PRICE", "")),
                            "NonCapExpenses.NonCapAmount": str(noncap_sum),
                        })

                        processed_rows.append(row_data)
                    except Exception as row_error:
                        # Track error for this row (Excel row number = index + 2, since index is 0-based and we skip header)
                        excel_row_num = int(index) + 2
                        side_value = row.get("SIDE", "")
                        error_msg = f"Row {excel_row_num} (SIDE='{side_value}'): {str(row_error)}"
                        row_errors.append(error_msg)
                        # Continue processing other rows
                        continue
                
                # Show user-friendly error message if any row errors occurred
                if row_errors:
                    error_count = len(row_errors)
                    error_details = "\n".join(row_errors[:10])  # Show first 10 errors
                    if error_count > 10:
                        error_details += f"\n... and {error_count - 10} more error(s)"
                    
                    messagebox.showerror(
                        "Processing Errors",
                        f"Found {error_count} error(s) while processing {os.path.basename(file_path)}:\n\n"
                        f"{error_details}\n\n"
                        f"Valid rows were processed successfully."
                    )
                
                # Store processed data for this file (after processing all rows)
                self.processed_files_data[file_path] = {
                    'processed_data': processed_rows,
                    'template_data': processed_rows,
                    'config': gtn_loader_config
                }
                
                # Add to combined data for table display
                all_processed_rows.extend(processed_rows)
                successful_files += 1
                    
            except Exception as e:
                failed_files.append(f"{os.path.basename(file_path)}: {str(e)[:50]}")
                import traceback
                print(f"Error processing {file_path}:")
                print(traceback.format_exc())
                continue
        
        # Update table with combined data from all files
        if all_processed_rows:
            # Store combined data for table display
            self.processed_data = all_processed_rows
            self._template_data = all_processed_rows
            self.gtn_loader_config = gtn_loader_config  # Store config for later use
            
            # Prepare table columns based on GTN loader config keys
            if all_processed_rows:
                self.table_columns = tuple(all_processed_rows[0].keys())
            else:
                self.table_columns = tuple(gtn_loader_config.keys())
            
            # Configure table with dynamic columns
            self.tree.configure(columns=self.table_columns)
            
            # Hide the tree column (#0) to prevent stretching issues
            self.tree.column("#0", width=0, stretch=False, minwidth=0)
            
            # Define column widths mapping for GTN loader columns
            col_widths = {
                "RecordType": 100,
                "RecordAction": 120,
                "KeyValue": 200,
                "KeyValue.KeyName": 200,
                "UserTranId1": 120,
                "Portfolio": 120,
                "LocationAccount": 200,
                "Strategy": 100,
                "Investment": 200,
                "Broker": 150,
                "EventDate": 120,
                "SettleDate": 120,
                "ActualSettleDate": 140,
                "Quantity": 100,
                "Price": 100,
                "PriceDenomination": 150,
                "CounterInvestment": 150,
                "NetInvestmentAmount": 150,
                "NetCounterAmount": 150,
                "tradeFX": 100,
                "ContractFxRateNumerator": 180,
                "ContractFxRateDenominator": 190,
                "ContractFxRate": 150,
                "NotionalAmount": 130,
                "FundStructure": 130,
                "SpotDate": 120,
                "PriceDirectly": 120,
                "CounterFXDenomination": 180,
                "CounterTDateFx": 140,
                "AccruedInterest": 130,
                "InvestmentAccruedInterest": 200,
                "Comments": 200,
                "TradeExpenses.ExpenseNumber": 200,
                "TradeExpenses.ExpenseCode": 200,
                "TradeExpenses.ExpenseAmt": 200,
                "TradeExpenses.ExpenseNumber1": 220,
                "TradeExpenses.ExpenseCode1": 220,
                "TradeExpenses.ExpenseAmt1": 220,
                "TradeExpenses.ExpenseNumber2": 220,
                "TradeExpenses.ExpenseCode2": 220,
                "TradeExpenses.ExpenseAmt2": 220,
                "NonCapExpenses.NonCapNumber": 220,
                "NonCapExpenses.NonCapExpenseCode": 250,
                "NonCapExpenses.NonCapAmount": 200,
                "NonCapExpenses.NonCapCurrency": 200,
                "NonCapExpenses.LocationAccount": 250,
                "NonCapExpenses.NonCapLiabilityCode": 250,
                "NonCapExpenses.NonCapPaymentType": 220,
                "NonCapExpenses.NonCapNumber1": 230,
                "NonCapExpenses.NonCapExpenseCode1": 260,
                "NonCapExpenses.NonCapAmount1": 210,
                "NonCapExpenses.NonCapCurrency1": 210,
                "NonCapExpenses.LocationAccount1": 260,
                "NonCapExpenses.NonCapLiabilityCode1": 260,
                "NonCapExpenses.NonCapPaymentType1": 230,
                "NonCapExpenses.NonCapNumber2": 230,
                "NonCapExpenses.NonCapExpenseCode2": 260,
                "NonCapExpenses.NonCapAmount2": 210,
                "NonCapExpenses.NonCapCurrency2": 210,
                "NonCapExpenses.LocationAccount2": 260,
                "NonCapExpenses.NonCapLiabilityCode2": 260,
                "NonCapExpenses.NonCapPaymentType2": 230,
            }
            
            for col in self.table_columns:
                self.tree.heading(col, text=col)
                # Use configured width if available, otherwise calculate based on column name length
                col_width = col_widths.get(col, min(max(len(col) * 10, 100), 300))
                # Set stretch=False and minwidth to prevent columns from stretching or shrinking
                self.tree.column(col, width=col_width, minwidth=col_width, anchor="w", stretch=False)
            
            # Convert processed rows (dictionaries) to list of tuples for table display
            self.table_rows = [tuple(row_data.get(col, "") for col in self.table_columns) for row_data in processed_rows]
            
            # Render the table
            self._render_table()
            
            # Update status
            status_msg = f"Processed {successful_files}/{total_files} file(s). {len(self.table_rows)} total rows loaded."
            if failed_files:
                status_msg += f"\nFailed: {', '.join(failed_files)}"
            self.status_var.set(status_msg)
            
            if successful_files > 0:
                messagebox.showinfo("Success", f"Successfully processed {successful_files} file(s).\n{len(self.table_rows)} total rows loaded.\n\nClick 'Export to Template' to generate loader files.")
            else:
                messagebox.showerror("Error", f"Failed to process all files.\n\nErrors:\n" + "\n".join(failed_files))

    def _render_table(self):
        """Render table with filtered data based on search."""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Get search term
        search_term = self.search_var.get().lower()
        
        # Filter and display rows
        for row in self.table_rows:
            # Apply search filter if search term exists
            if search_term:
                # Check if search term matches any value in the row
                row_str = " ".join(str(val) for val in row).lower()
                if search_term not in row_str:
                    continue
            
            self.tree.insert("", "end", values=row)

    def _prepare_loader_data(self):
        """
        Prepare loader data based on GTN_LOADER headers.
        Uses self.processed_data and returns data formatted with GTN_LOADER keys as headers.
        
        Returns:
            tuple: (headers, data_rows) where headers is list of GTN_LOADER keys and data_rows is list of lists
        """
        if not self.processed_data or not hasattr(self, 'gtn_loader_config'):
            return [], []
        
        # Get headers from GTN_LOADER config (these are the field names)
        headers = list(self.gtn_loader_config.keys())
        
        # Prepare data rows - each row is a list of values in the same order as headers
        data_rows = []
        for row_dict in self.processed_data:
            # Create a row with values in the same order as headers
            row_values = [row_dict.get(header, "") for header in headers]
            data_rows.append(row_values)
        
        return headers, data_rows

    def _export_excel(self):
        """Export processed data to Excel and/or CSV file based on format selection."""
        if not self.processed_data:
            messagebox.showwarning("No Data", "No data to export. Please process file(s) first.")
            return

        export_csv = self.csv_format_var.get()
        export_xlsx = self.xlsx_format_var.get()
        
        if not export_csv and not export_xlsx:
            messagebox.showwarning("No Format Selected", "Please select at least one export format (CSV or XLSX).")
            return
        
        try:
            # Prepare loader data with headers
            headers, data_rows = self._prepare_loader_data()
            
            if not headers or not data_rows:
                messagebox.showwarning("No Data", "No data to export. Please process file(s) first.")
                return
            
            # Import pandas for export
            import pandas as pd
            from decimal import Decimal
            
            # Create DataFrame with headers and data
            df = pd.DataFrame(data_rows, columns=headers)
            exported_files = []

            # Export CSV first (default format)
            if export_csv:
                out_path = filedialog.asksaveasfilename(
                    title="Save CSV", defaultextension=".csv",
                    filetypes=[["CSV Files", "*.csv"]],
                    initialfile="GTN_Loader_Export.csv"
                )
                if out_path:
                    df_csv = df.copy()
                    for col in df_csv.columns:
                        if df_csv[col].dtype == 'object' and len(df_csv) > 0:
                            non_null_vals = df_csv[col].dropna()
                            if not non_null_vals.empty and isinstance(non_null_vals.iloc[0], Decimal):
                                df_csv[col] = df_csv[col].apply(lambda x: str(x) if pd.notna(x) else '')
                    df_csv.to_csv(out_path, index=False, encoding='utf-8-sig')
                    exported_files.append(f"CSV: {out_path}")

            # Export XLSX
            if export_xlsx:
                out_path = filedialog.asksaveasfilename(
                    title="Save Excel", defaultextension=".xlsx",
                    filetypes=[["Excel Files", "*.xlsx"]],
                    initialfile="GTN_Loader_Export.xlsx"
                )
                if out_path:
                    df.to_excel(out_path, index=False, engine='openpyxl')
                    exported_files.append(f"XLSX: {out_path}")

            if exported_files:
                messagebox.showinfo("Success", f"Files exported successfully:\n\n" + "\n".join(exported_files))
                self.status_var.set(f"Exported: {len(data_rows)} rows")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export: {e}")
            import traceback
            print(traceback.format_exc())

    def _export_to_template(self):
        """Export processed data to template format (ZIP with Excel/CSV files) - generates one loader file per input file."""
        if not self.processed_files_data:
            messagebox.showwarning("No Data", "No template data to export. Please process file(s) first.")
            return

        # Check format selection
        export_csv = self.csv_format_var.get()
        export_xlsx = self.xlsx_format_var.get()
        
        if not export_csv and not export_xlsx:
            messagebox.showwarning("No Format Selected", "Please select at least one export format (CSV or XLSX).")
            return
        
        try:
            # Ask for ZIP output path
            from tkinter import filedialog
            out_path = filedialog.asksaveasfilename(
                title="Save Template Data",
                defaultextension=".zip",
                filetypes=[["ZIP Files", "*.zip"]],
                initialfile="GTN_Loader_Template.zip"
            )
            
            if not out_path:
                return
            
            # Use helper functions
            from .helper import output_save_in_template, output_save_in_template_csv, multiple_files_to_zip
            import threading
            from my_app.pages.loading import LoadingSpinner
            
            # Show spinner (non-blocking)
            loader = LoadingSpinner(self, text="Exporting templates...")

            def task():
                try:
                    files = []
                    failed_exports = []
                    
                    # Generate loader file for each processed file
                    for file_path, file_data in self.processed_files_data.items():
                        try:
                            processed_rows = file_data['processed_data']
                            gtn_loader_config = file_data['config']
                            
                            if not processed_rows:
                                continue
                            
                            # Get headers from GTN_LOADER config
                            headers = list(gtn_loader_config.keys())
                            
                            # Convert to list of dicts for helper function
                            data_dicts = []
                            for row_dict in processed_rows:
                                data_dicts.append(row_dict)
                            
                            # Generate output filename base: {input_filename}_loader
                            input_filename = os.path.basename(file_path)
                            base_name = os.path.splitext(input_filename)[0]
                            
                            # Export CSV if selected
                            if export_csv:
                                file_io, file_name = output_save_in_template_csv(
                                    data_dicts,
                                    headers,
                                    f"{base_name}_loader.csv"
                                )
                                files.append((file_io, file_name))
                            
                            # Export XLSX if selected
                            if export_xlsx:
                                file_io, file_name = output_save_in_template(
                                    data_dicts,
                                    headers,
                                    f"{base_name}_loader.xlsx"
                                )
                                files.append((file_io, file_name))
                            
                        except Exception as e:
                            failed_exports.append(f"{os.path.basename(file_path)}: {str(e)[:50]}")
                            import traceback
                            print(f"Error exporting {file_path}:")
                            print(traceback.format_exc())
                            continue
                    
                    if not files:
                        loader.close()
                        messagebox.showwarning("Warning", "No template data to export.")
                        return
                    
                    # Create zip file
                    zip_buffer = multiple_files_to_zip(files, "GTN_Loader_Template.zip")
                    
                    # Save zip file
                    with open(out_path, 'wb') as f:
                        f.write(zip_buffer.read())
                    
                    file_list = [name for _, name in files]
                    loader.close()
                    
                    email_zip_path = out_path
                    email_file_list = file_list.copy()
                    
                    def show_success_and_dialog():
                        success_msg = f"Template data exported to:\n{email_zip_path}\n\nContains {len(email_file_list)} file(s):\n" + "\n".join([f"- {file}" for file in email_file_list])
                        if failed_exports:
                            success_msg += f"\n\nFailed exports:\n" + "\n".join(failed_exports)
                        messagebox.showinfo("Success", success_msg)
                        self.after(100, lambda: self._show_email_dialog(email_zip_path, email_file_list))
                    
                    self.after(0, show_success_and_dialog)
                    
                except Exception as e:
                    loader.close()
                    messagebox.showerror("Error", f"Failed to export template: {e}")
                    import traceback
                    print(traceback.format_exc())
            
            # Run heavy work in thread
            threading.Thread(target=task, daemon=True).start()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export template: {e}")
            import traceback
            print(traceback.format_exc())
    
    def _show_email_dialog(self, zip_path, file_list):
        """Show email dialog for sending files via Outlook."""
        try:
            from .email_dialog import EmailDialog
            EmailDialog(self, zip_path, file_list)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open email dialog: {e}")
            import traceback
            traceback.print_exc()

