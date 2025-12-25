import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
from datetime import datetime

try:
    from my_app.CONSTANTS import fields as DEFAULT_HEADER_FIELDS
except Exception:
    DEFAULT_HEADER_FIELDS = []


class DataConfigPage(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg="#ecf0f1")
        self.default_lotsize_data = {
            "ALUMINI": 1000,
            "ALUMINIUM": 5000,
            "CASTOR": 50,
            "COCUDAKL": 100,
            "COPPER": 2500,
            "CRUDEOIL": 100,
            "CRUDEOILM": 10,
            "DHANIYA": 50,
            "GOLD": 100,
            "GOLDM": 10,
            "GUARGUM5": 50,
            "GUARSEED10": 50,
            "JEERAUNJHA": 30,
            "LEAD": 5000,
            "LEADMINI": 1000,
            "NATGASMINI": 250,
            "NATGAS": 1250,
            "SILVER": 30,
            "SILVERM": 5,
            "SILVERMIC": 1,
            "TMCFGRNZM": 50,
            "ZINC": 5000,
            "ZINCMINI": 1000,
            "JEERAMINI": 10,
            "KAPAS": 200,
            "MCXBULLDEX": 30,
            "WTICRUDE": 100,
            "NATURALGAS": 1250,
            "GOLDTEN": 1,
            "ELECDMBL": 50,
            "ELECMBL": 50,
        }
        self.lotsize_data = self.default_lotsize_data.copy()
        self.default_underlying_code_data = {
            "CRUDEOIL": 100,
            "NATURALGAS": 1250,
        }
        self.underlying_code_data = self.default_underlying_code_data.copy()
        # Discover header datasets dynamically from JSON files
        self.dataset_files = self._discover_datasets()
        # Hold in-memory data per header dataset name
        self.datasets = {}
        # Seed defaults for each dataset
        for dataset_name in self.dataset_files.keys():
            self.datasets[dataset_name] = self._load_default_dataset_data(dataset_name)
        # Track current selection for header datasets (non-lotsize)
        self.current_dataset_name = next(iter(self.dataset_files.keys()), "") or ""
        # Alias for currently selected header dataset data
        self.header_data = self.datasets.get(self.current_dataset_name, {})
        self.mode_var = tk.StringVar(value="LOTSIZE")
        self.setup_ui()
        self.load_saved_data_on_startup()
        self.load_data()

    def setup_ui(self):
        # Light style for tree headers
        style = ttk.Style(self)
        try:
            style.configure("Treeview.Heading", background="#f2f4f7", foreground="#2c3e50", font=("Arial", 11, "bold"))
        except Exception:
            pass

        header_frame = tk.Frame(self, bg="#ecf0f1")
        header_frame.pack(fill="x", padx=20, pady=10)
        tk.Label(header_frame, text="📊 Data Configuration", font=("Arial", 20, "bold"), bg="#ecf0f1", fg="#2c3e50").pack(side="left")
        selector_frame = tk.Frame(header_frame, bg="#ecf0f1")
        selector_frame.pack(side="left", padx=20)
        tk.Label(selector_frame, text="Dataset:", font=("Arial", 11), bg="#ecf0f1").pack(side="left")
        
        # Store all available dataset values for search filtering
        self._all_dataset_values = ["Lotsize", "UnderlyingCode"] + list(self.dataset_files.keys())
        self._base_dataset_values = ["Lotsize", "UnderlyingCode"]
        
        # Search entry (simple, no container needed)
        self.dataset_var = tk.StringVar(value="Lotsize")
        self.search_entry = tk.Entry(selector_frame, textvariable=self.dataset_var, width=28, font=("Arial", 10))
        self.search_entry.pack(side="left", padx=(6, 0))
        
        # Results listbox (positioned absolutely, doesn't affect layout)
        self.results_listbox = tk.Listbox(self, width=28, height=8, font=("Arial", 10), relief="solid", borderwidth=1)
        # Initially hidden, will be positioned using place() when shown
        
        # Real-time search as user types
        def on_search_change(*args):
            search_term = self.dataset_var.get().strip().lower()
            
            if search_term:
                # Filter values that contain the search term
                filtered = [v for v in self._all_dataset_values if search_term in v.lower()]
            else:
                # Show all values if search is empty
                filtered = self._all_dataset_values
            
            # Update listbox
            self.results_listbox.delete(0, tk.END)
            for item in filtered:
                self.results_listbox.insert(tk.END, item)
            
            # Show listbox if there are results
            if filtered:
                self._show_listbox()
            else:
                self._hide_listbox()
        
        # Bind search entry changes
        self.dataset_var.trace_add("write", on_search_change)
        
        # Show listbox when search entry gets focus
        def on_search_focus_in(event):
            # Show all values when clicking into search
            search_term = self.dataset_var.get().strip().lower()
            if not search_term:
                filtered = self._all_dataset_values
                self.results_listbox.delete(0, tk.END)
                for item in filtered:
                    self.results_listbox.insert(tk.END, item)
                self._show_listbox()
        
        self.search_entry.bind('<FocusIn>', on_search_focus_in)
        
        # Track if listbox is visible
        self._listbox_visible = False
        
        # Method to show listbox positioned below search entry
        def _show_listbox():
            try:
                # Get search entry position
                self.update_idletasks()
                x = self.search_entry.winfo_rootx() - self.winfo_rootx()
                y = self.search_entry.winfo_rooty() - self.winfo_rooty() + self.search_entry.winfo_height()
                width = self.search_entry.winfo_width()
                
                # Position listbox below search entry
                self.results_listbox.place(x=x, y=y, width=width)
                self.results_listbox.lift()
                self._listbox_visible = True
            except:
                pass
        
        # Method to hide listbox
        def _hide_listbox():
            self.results_listbox.place_forget()
            self._listbox_visible = False
        
        self._show_listbox = _show_listbox
        self._hide_listbox = _hide_listbox
        
        # When user selects from listbox
        def on_listbox_select(event):
            selection = self.results_listbox.curselection()
            if selection:
                selected_value = self.results_listbox.get(selection[0])
                self.dataset_var.set(selected_value)
                self._hide_listbox()
                on_dataset_change()
        
        self.results_listbox.bind('<Double-Button-1>', on_listbox_select)
        self.results_listbox.bind('<Return>', on_listbox_select)
        
        # When user presses Enter in search entry
        def on_entry_return(event):
            # Get first item from listbox if visible
            if self._listbox_visible:
                items = self.results_listbox.get(0, tk.END)
                if items:
                    self.dataset_var.set(items[0])
                    self._hide_listbox()
                    on_dataset_change()
        
        self.search_entry.bind('<Return>', on_entry_return)
        
        # Hide listbox when clicking outside
        def hide_listbox(event):
            if event.widget != self.search_entry and event.widget != self.results_listbox:
                self._hide_listbox()
        
        self.bind('<Button-1>', hide_listbox)
        
        refresh_btn = tk.Button(selector_frame, text="↻", width=3, command=self.refresh_datasets, bg="#ecf0f1", relief="flat")
        refresh_btn.pack(side="left", padx=(6, 0))

        def on_dataset_change(event=None):
            sel = self.dataset_var.get()
            if sel == "Lotsize":
                self.mode_var.set("LOTSIZE")
            elif sel == "UnderlyingCode":
                self.mode_var.set("UNDERLYINGCODE")
            else:
                # Switch to header mode and point to the selected dataset
                self.mode_var.set("HEADER")
                self.current_dataset_name = sel
                # Ensure dataset exists in memory - try to load from consolidated_data.json first
                if sel not in self.datasets or not self.datasets[sel]:
                    # Try to load from consolidated_data.json
                    try:
                        from my_app.file_utils import get_app_directory
                        app_dir = get_app_directory()
                        consolidated_path = os.path.join(app_dir, "consolidated_data.json")
                        if os.path.exists(consolidated_path):
                            with open(consolidated_path, "r", encoding="utf-8") as f:
                                consolidated_data = json.load(f)
                                if sel in consolidated_data:
                                    self.datasets[sel] = consolidated_data[sel]
                                else:
                                    # Load default if not found
                                    self.datasets[sel] = self._load_default_dataset_data(sel)
                        else:
                            # File doesn't exist, use default
                            self.datasets[sel] = self._load_default_dataset_data(sel)
                    except Exception as e:
                        # On error, use default
                        print(f"Error loading dataset {sel}: {e}")
                        self.datasets[sel] = self._load_default_dataset_data(sel)
                self.header_data = self.datasets[sel]
            self.configure_tree_for_mode()
            self.load_data()

        control_frame = tk.Frame(header_frame, bg="#ecf0f1")
        control_frame.pack(side="right")
        tk.Button(control_frame, text="➕ Add New", bg="#3498db", fg="white", font=("Arial", 10, "bold"), relief="flat", padx=15, pady=5, command=self.add_new_item).pack(side="left", padx=5)
        tk.Button(control_frame, text="💾 Save", bg="#27ae60", fg="white", font=("Arial", 10, "bold"), relief="flat", padx=15, pady=5, command=self.save_data).pack(side="left", padx=5)
        tk.Button(control_frame, text="📁 Load", bg="#f39c12", fg="white", font=("Arial", 10, "bold"), relief="flat", padx=15, pady=5, command=self.load_from_file).pack(side="left", padx=5)
        tk.Button(control_frame, text="🔄 Reset", bg="#e67e22", fg="white", font=("Arial", 10, "bold"), relief="flat", padx=15, pady=5, command=self.reset_to_default).pack(side="left", padx=5)

        search_frame = tk.Frame(self, bg="#ecf0f1")
        search_frame.pack(fill="x", padx=20, pady=5)
        tk.Label(search_frame, text="🔍 Search:", font=("Arial", 12), bg="#ecf0f1", fg="#2c3e50").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self.filter_data)
        tk.Entry(search_frame, textvariable=self.search_var, font=("Arial", 12), width=30, relief="solid", bd=1).pack(side="left", padx=10)

        content_frame = tk.Frame(self, bg="#ecf0f1")
        content_frame.pack(fill="both", expand=True, padx=20, pady=10)
        table_frame = tk.Frame(content_frame, bg="#ecf0f1")
        table_frame.pack(fill="both", expand=True)
        tree_frame = tk.Frame(table_frame, bg="#ecf0f1")
        tree_frame.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(tree_frame, columns=(), show="headings", height=15)
        self.configure_tree_for_mode()
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.actions_frame = tk.Frame(table_frame, bg="#ecf0f1")
        self.actions_frame.pack(fill="x", pady=10)
        self.edit_btn = tk.Button(self.actions_frame, text="✏️ EDIT", bg="#bdc3c7", fg="#7f8c8d", font=("Arial", 12), relief="flat", bd=1, padx=20, pady=8, command=self.edit_selected_item, state="disabled")
        self.edit_btn.pack(side="left", padx=10)
        self.delete_btn = tk.Button(self.actions_frame, text="🗑️ DELETE", bg="#bdc3c7", fg="#7f8c8d", font=("Arial", 12), relief="flat", bd=1, padx=20, pady=8, command=self.delete_selected_item, state="disabled")
        self.delete_btn.pack(side="left", padx=10)
        self.selected_info = tk.Label(self.actions_frame, text="Select a row to edit or delete", font=("Arial", 10), bg="#ecf0f1", fg="#7f8c8d")
        self.selected_info.pack(side="right", padx=10)
        self.tree.bind("<<TreeviewSelect>>", self.on_selection_change)
        self.tree.bind("<Double-1>", self.edit_item)
        self.tree.bind("<Button-3>", self.show_context_menu)
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="Edit", command=self.edit_item)
        self.context_menu.add_command(label="Delete", command=self.delete_item)

    def load_saved_data_on_startup(self):
        try:
            # Load consolidated data
            from my_app.file_utils import get_app_directory
            app_dir = get_app_directory()
            consolidated_path = os.path.join(app_dir, "consolidated_data.json")
            
            # List of all expected datasets
            expected_datasets = ["fund_filename_map", "asio_recon_portfolio_mapping", "asio_recon_format_1_headers", "asio_recon_format_2_headers", 
                                "asio_recon_bhavcopy_headers", "asio_geneva_headers", "trade_headers", 
                                "aafspl_car_future", "option_security", "car_trade_loader", "asio_sf_2_trade_loader", "asio_sf_2_option_security", "asio_sf_2_future_security", "asio_sf_2_mcx_trade_loader", "asio_sf_2_mcx_option_security", "asio_sf_2_mcx_future_security", "geneva_custodian_mapping", "fno_tm_code_with_tm_name", "mcx_tm_code_with_tm_name", "asio_pricing_fno", "asio_pricing_mcx", "asio_sf4_ft", "asio_sf4_trading_code_mapping", "asio_sub_fund4_read_config", "mcx_group2_filters", "fno_group2_filters"]
            
            if os.path.exists(consolidated_path):
                # Load file; if unreadable/empty, recreate with defaults
                consolidated_data = {}
                try:
                    with open(consolidated_path, "r") as f:
                        consolidated_data = json.load(f)
                except Exception:
                    consolidated_data = {}
                if not isinstance(consolidated_data, dict) or not consolidated_data:
                    default_consolidated_data = {
                        "lotsize_data": self.default_lotsize_data.copy(),
                        "underlying_code_data": self.default_underlying_code_data.copy(),
                        "fund_filename_map": {"DBSBK0000033": {"Fund Names": {"Default": "DIF-Class 1 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000036": {"Fund Names": {"Default": "DIF-Class 2 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000038": {"Fund Names": {"Default": "DIF-Class 3 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000042": {"Fund Names": {"Default": "DIF-Class 5 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000044": {"Fund Names": {"Default": "DIF-Class 6 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000043": {"Fund Names": {"Default": "DIF-Class 7 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000049": {"Fund Names": {"Default": "DIF-Class 8 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000050": {"Fund Names": {"Default": "DIF-Class 9 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000051": {"Fund Names": {"Default": "DIF-Class 10 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000052": {"Fund Names": {"Default": "DIF-Class 11 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000074": {"Fund Names": {"Default": "DIF-Class 12 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000179": {"Fund Names": {"Default": "DIF-Class 13 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000189": {"Fund Names": {"Default": "DIF-Class 14 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000192": {"Fund Names": {"Default": "DIF-Class 15 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000214": {"Fund Names": {"Default": "DIF-Class 16 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000216": {"Fund Names": {"Default": "DIF-Class 17 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000217": {"Fund Names": {"Default": "DIF-Class 18_Moon"}, "Password": "AAGCD0792B"}, "DBSBK0000232": {"Fund Names": {"Default": "DIF-Class 19 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000247": {"Fund Names": {"CDS": "DIF-Class 21 CDS Holding", "Default": "DIF-Class 21 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000178": {"Fund Names": {"Default": "DGF-Cell 8"}, "Password": "AAICD1968M"}, "BNPP00000458": {"Fund Names": {"Default": "DGF-Cell 9"}, "Password": "AAICD2891H"}, "DGF-Cell 10": {"Fund Names": {"Default": "DGF-Cell 10"}, "Password": "AAICD3412C"}, "BNPP00000480": {"Fund Names": {"Default": "DGF-Cell 11"}, "Password": "AAICD6359G"}, "BNPP00000488": {"Fund Names": {"Default": "DGF-Cell 13"}, "Password": "AAICD7821M"}, "BNPP00000540": {"Fund Names": {"Default": "DGF-Cell 16"}, "Password": "AAJCD5624K"}, "BNPP00000535": {"Fund Names": {"Default": "DGF-Cell 17"}, "Password": "AAJCD4991K"}, "DBSBK0000229": {"Fund Names": {"Default": "DGF-Cell 18"}, "Password": "AAJCD6205G"}, "DBSBK0000228": {"Fund Names": {"Default": "DGF-Cell 19"}, "Password": "AAJCD6049E"}, "DBSBK0000285": {"Fund Names": {"CDS": "DGF-Cell 23 CDS Holding", "Default": "DGF-Cell 23 Holding"}, "Password": "AAKCD6244Q"}, "DBSBK0000299": {"Fund Names": {"Default": "DGF-Cell 24 Holding"}, "Password": "AAKCD7324B"}, "DBSBK0000353": {"Fund Names": {"Default": "DGF-Cell 28 Holding"}, "Password": "AALCD1140J"}, "DBSBK0000354": {"Fund Names": {"Default": "DGF-Cell 29 Holding"}, "Password": "AALCD1141K"}, "DBSBK0000380": {"Fund Names": {"Default": "DGF-Cell 32 Holding"}, "Password": ""}, "DBSBK0000356": {"Fund Names": {"Default": "GlobalQ_AIF-III"}, "Password": ""}, "DGF-Cell 36": {"Fund Names": {"Default": "DGF-Cell 36"}, "Password": ""}, "DGF-Cell 38": {"Fund Names": {"Default": "DGF-Cell 38"}, "Password": ""}},
                        "asio_sf_2_trade_loader": {
                            "RecordAction": "InsertUpdate",
                            "KeyValue.KeyName": "",
                            "UserTranId1": "",
                            "Portfolio": "ASIO - SF 2_INR",
                            "LocationAccount": "Asio_Sub Fund_2_OHM_FO_KOTBK0001479",
                            "Broker": "",
                            "PriceDenomination": "CALC",
                            "CounterInvestment": "INR",
                            "NetInvestmentAmount": "CALC",
                            "NetCounterAmount": "CALC",
                            "tradeFX": "",
                            "ContractFxRateNumerator": "",
                            "ContractFxRateDenominator": "",
                            "ContractFxRate": "",
                            "NotionalAmount": "",
                            "FundStructure2": "CALC",
                            "SpotDate": "",
                            "PriceDirectly": "",
                            "CounterFXDenomination": "CALC",
                            "CounterTDateFx": "",
                            "AccruedInterest": "",
                            "InvestmentAccruedInterest": "",
                            "TradeExpenses.ExpenseNumber": "",
                            "TradeExpenses.ExpenseCode": "",
                            "TradeExpenses.ExpenseAmt": "",
                            "NonCapExpenses.NonCapNumber": "",
                            "NonCapExpenses.NonCapExpenseCode": "",
                            "NonCapExpenses.NonCapAmount": "",
                            "NonCapExpenses.NonCapCurrency": "",
                            "NonCapExpenses.LocationAccount": "",
                            "NonCapExpenses.NonCapLiabilityCode": "",
                            "NonCapExpenses.NonCapPaymentType": ""
                        },
                        "aafspl_car_future": {
                            "Exchange": "MCX",
                            "Issuer": "",
                            "Ticker": "",
                            "Cusip": "",
                            "Sedol": "",
                            "Isin": "",
                            "AltKey1": "",
                            "AltKey2": "",
                            "BloombergID": "",
                            "BloombergTicker": "",
                            "BloombergUniqueID": "",
                            "BloombergMarketSector": "",
                            "SettleDays": "",
                            "PricingFactor": "",
                            "CurrentPriceDayRange": 1,
                            "AssetType": "FT",
                            "InvestmentType": "CMFT",
                            "PriceDenomination": "INR",
                            "BifurcationCurrency": "INR",
                            "PrincipalCurrency": "INR",
                            "IncomeCurrency": "INR",
                            "RiskCurrency": "INR",
                            "IssueCountry": "IN",
                            "WithholdingTaxType": "Standard",
                            "QDIEligibilityFlag": "Not Eligible",
                            "SharesOutstanding": "",
                            "SubIndustry": "",
                            "SubIndustry2": "",
                            "QuantityPrecision": 3,
                            "InvestmentCrossZero": "Use Accounting Parameters",
                            "PriceByPreference": "Currency",
                            "ForwardPriceInterpolateFlag": 0,
                            "PricingPrecision": 3,
                            "FirstMarkDate": "01-01-2022",
                            "LastMarkDate": "",
                            "PriceList": "",
                            "AutoGenerateMarks": 1,
                            "CashSettlement": 1
                        },
                        "option_security": {
                            "Exchange": "MCX",
                            "Issuer": "",
                            "Ticker": "",
                            "Cusip": "",
                            "Sedol": "",
                            "Isin": "",
                            "Riccode": "",
                            "AltKey1": "",
                            "AltKey2": "",
                            "BloombergID": "",
                            "BloombergTicker": "",
                            "BloombergUniqueID": "",
                            "BloombergMarketSector": "",
                            "SettleDays": 0,
                            "PricingFactor": 1,
                            "CurrentPriceDayRange": 0,
                            "AssetType": "OP",
                            "InvestmentType": "CMFTOP",
                            "PriceDenomination": "",
                            "BifurcationCurrency": "INR",
                            "PrincipalCurrency": "INR",
                            "IncomeCurrency": "INR",
                            "RiskCurrency": "INR",
                            "IssueCountry": "IN",
                            "WithholdingTaxType": "Standard",
                            "QDIEligibilityFlag": "Not Eligible",
                            "SharesOutstanding": "",
                            "Beta": "",
                            "SubIndustry": "",
                            "SubIndustry2": "",
                            "NonDeliverableCurrencyFlag": "",
                            "QuantityPrecision": "",
                            "InvestmentCrossZero": "",
                            "PriceByPreference": "",
                            "ForwardPriceInterpolateFlag": "",
                            "SecFeeSchedule": "",
                            "SecEligibleFlag": "",
                            "WashSalesEligible": "",
                            "ExerciseStyle": "European",
                            "PricingPrecision": 15,
                            "CashSettledFlag": 1
                        },
                        "asio_sf_2_option_security": {
                            "Exchange": "NSE",
                            "Issuer": "",
                            "Ticker": "",
                            "Cusip": "",
                            "Sedol": "",
                            "Isin": "",
                            "Riccode": "",
                            "AltKey1": "",
                            "AltKey2": "",
                            "BloombergID": "",
                            "BloombergTicker": "",
                            "BloombergUniqueID": "",
                            "BloombergMarketSector": "",
                            "SettleDays": "",
                            "PricingFactor": "",
                            "TradingFactor": 1,
                            "CurrentPriceDayRange": 1,
                            "AssetType": "FT",
                            "InvestmentType": "EQFT",
                            "PriceDenomination": "INR",
                            "BifurcationCurrency": "INR",
                            "PrincipalCurrency": "INR",
                            "IncomeCurrency": "INR",
                            "RiskCurrency": "INR",
                            "IssueCountry": "IN",
                            "WithholdingTaxType": "Standard",
                            "QDIEligibilityFlag": "Not Eligible",
                            "SharesOutstanding": "",
                            "SubIndustry": "",
                            "SubIndustry2": "",
                            "QuantityPrecision": 3,
                            "InvestmentCrossZero": "Use Accounting Parameters",
                            "PriceByPreference": "Currency",
                            "ForwardPriceInterpolateFlag": 0,
                            "PricingPrecision": 3,
                            "FirstMarkDate": "01-01-2022",
                            "LastMarkDate": "",
                            "PriceList": "",
                            "AutoGenerateMarks": 1,
                            "CashSettlement": 1
                        },
                        "asio_sf_2_mcx_trade_loader": {
                            "RecordAction": "InsertUpdate",
                            "KeyValue.KeyName": "",
                            "UserTranId1": "",
                            "Portfolio": "ASIO - SF 2_INR",
                            "LocationAccount": "Asio_Sub Fund_2_OHM_MCX_KOTBK0001479",
                            "Broker": "",
                            "PriceDenomination": "CALC",
                            "CounterInvestment": "INR",
                            "NetInvestmentAmount": "CALC",
                            "NetCounterAmount": "CALC",
                            "tradeFX": "",
                            "ContractFxRateNumerator": "",
                            "ContractFxRateDenominator": "",
                            "ContractFxRate": "",
                            "NotionalAmount": "",
                            "FundStructure2": "CALC",
                            "SpotDate": "",
                            "PriceDirectly": "",
                            "CounterFXDenomination": "CALC",
                            "CounterTDateFx": "",
                            "AccruedInterest": "",
                            "InvestmentAccruedInterest": "",
                            "TradeExpenses.ExpenseNumber": "",
                            "TradeExpenses.ExpenseCode": "",
                            "TradeExpenses.ExpenseAmt": "",
                            "NonCapExpenses.NonCapNumber": "",
                            "NonCapExpenses.NonCapExpenseCode": "",
                            "NonCapExpenses.NonCapAmount": "",
                            "NonCapExpenses.NonCapCurrency": "",
                            "NonCapExpenses.LocationAccount": "",
                            "NonCapExpenses.NonCapLiabilityCode": "",
                            "NonCapExpenses.NonCapPaymentType": ""
                        },
                        "asio_sf_2_mcx_option_security": {
                            "Exchange": "MCX",
                            "Issuer": "",
                            "Ticker": "",
                            "Cusip": "",
                            "Sedol": "",
                            "Isin": "",
                            "Riccode": "",
                            "AltKey1": "",
                            "AltKey2": "",
                            "BloombergID": "",
                            "BloombergTicker": "",
                            "BloombergUniqueID": "",
                            "BloombergMarketSector": "",
                            "SettleDays": 0,
                            "PricingFactor": 1,
                            "CurrentPriceDayRange": 1,
                            "AssetType": "OP",
                            "InvestmentType": "CMOP",
                            "PriceDenomination": "",
                            "BifurcationCurrency": "INR",
                            "PrincipalCurrency": "INR",
                            "IncomeCurrency": "INR",
                            "RiskCurrency": "INR",
                            "IssueCountry": "IN",
                            "WithholdingTaxType": "Standard",
                            "QDIEligibilityFlag": "Not Eligible",
                            "SharesOutstanding": "",
                            "Beta": "",
                            "SubIndustry": "",
                            "SubIndustry2": "",
                            "NonDeliverableCurrencyFlag": "",
                            "QuantityPrecision": "",
                            "InvestmentCrossZero": "",
                            "PriceByPreference": "",
                            "ForwardPriceInterpolateFlag": "",
                            "SecFeeSchedule": "",
                            "SecEligibleFlag": "",
                            "WashSalesEligible": "",
                            "ExerciseStyle": "European",
                            "PricingPrecision": 15,
                            "CashSettledFlag": 1
                        },
                        "asio_sf_2_mcx_future_security": {
                            "Exchange": "MCX",
                            "Issuer": "",
                            "Ticker": "",
                            "Cusip": "",
                            "Sedol": "",
                            "Isin": "",
                            "AltKey1": "",
                            "AltKey2": "",
                            "BloombergID": "",
                            "BloombergTicker": "",
                            "BloombergUniqueID": "",
                            "BloombergMarketSector": "",
                            "SettleDays": "",
                            "PricingFactor": "",
                            "CurrentPriceDayRange": 1,
                            "AssetType": "FT",
                            "InvestmentType": "CMFT",
                            "PriceDenomination": "INR",
                            "BifurcationCurrency": "INR",
                            "PrincipalCurrency": "INR",
                            "IncomeCurrency": "INR",
                            "RiskCurrency": "INR",
                            "IssueCountry": "IN",
                            "WithholdingTaxType": "Standard",
                            "QDIEligibilityFlag": "Not Eligible",
                            "SharesOutstanding": "",
                            "SubIndustry": "",
                            "SubIndustry2": "",
                            "QuantityPrecision": 3,
                            "InvestmentCrossZero": "Use Accounting Parameters",
                            "PriceByPreference": "Currency",
                            "ForwardPriceInterpolateFlag": 0,
                            "PricingPrecision": 3,
                            "FirstMarkDate": "01-01-2022",
                            "LastMarkDate": "",
                            "PriceList": "",
                            "AutoGenerateMarks": 1,
                            "CashSettlement": 1
                        },
                        "car_trade_loader": {
                            "RecordAction": "InsertUpdate",
                            "KeyValue.KeyName": "",
                            "UserTranId1": "",
                            "Portfolio": "AAFSPL_CAR",
                            "LocationAccount": "AAFSPL_CAR_Ventura Securities Ltd",
                            "Strategy": "Default",
                            "Broker": "",
                            "PriceDenomination": "CALC",
                            "CounterInvestment": "INR",
                            "NetInvestmentAmount": "CALC",
                            "NetCounterAmount": "CALC",
                            "TradeFX": "",
                            "ContractFxRateNumerator": "",
                            "ContractFxRateDenominator": "",
                            "ContractFxRate": "",
                            "NotionalAmount": "",
                            "FundStructure": "CALC",
                            "SpotDate": "",
                            "PriceDirectly": "",
                            "CounterFXDenomination": "CALC",
                            "CounterTDateFx": "",
                            "AccruedInterest": "",
                            "InvestmentAccruedInterest": "",
                            "TradeExpenses.ExpenseNumber": "",
                            "TradeExpenses.ExpenseCode": "",
                            "TradeExpenses.ExpenseAmt": "",
                            "NonCapExpenses.NonCapNumber": "",
                            "NonCapExpenses.NonCapExpenseCode": "",
                            "NonCapExpenses.NonCapAmount": "",
                            "NonCapExpenses.NonCapCurrency": "",
                            "NonCapExpenses.LocationAccount": "",
                        "NonCapExpenses.NonCapLiabilityCode": "",
                        "NonCapExpenses.NonCapPaymentType": ""
                        },
                        "asio_pricing_fno": {
                            "RecordAction": "InsertUpdate",
                            "PriceDenomination": "INR",
                            "PriceList": "NSE_Equity",
                            "TaxLotID": ""
                        },
                        "asio_pricing_mcx": {
                            "RecordAction": "InsertUpdate",
                            "PriceDenomination": "INR",
                            "PriceList": "NSE_Equity",
                            "TaxLotID": ""
                        },
                        "asio_sf4_trading_code_mapping": {
                            "FT": "Asio_Sub Fund_4_OHM_FO_DBSBK0000289_FT",
                            "FT1": "Asio_Sub Fund_4_OHM_FO_DBSBK0000289_FT1",
                            "FT2": "Asio_Sub Fund_4_OHM_FO_DBSBK0000289_FT2",
                            "FT3": "Asio_Sub Fund_4_OHM_FO_DBSBK0000289_FT3"
                        }
                    }
                    with open(consolidated_path, "w") as fw:
                        json.dump(default_consolidated_data, fw, indent=4)
                    consolidated_data = default_consolidated_data
                # Load lotsize data
                lotsize_data = consolidated_data.get("lotsize_data", {})
                if lotsize_data:
                    self.lotsize_data.update(lotsize_data)
                
                # Check for missing datasets and add defaults
                needs_save = False
                
                # Load underlying code data - always ensure it exists
                underlying_code_data = consolidated_data.get("underlying_code_data")
                if underlying_code_data and isinstance(underlying_code_data, dict) and len(underlying_code_data) > 0:
                    # Update with saved data
                    self.underlying_code_data = underlying_code_data.copy()
                else:
                    # If not found or empty, initialize with defaults
                    self.underlying_code_data = self.default_underlying_code_data.copy()
                    consolidated_data["underlying_code_data"] = self.underlying_code_data.copy()
                    needs_save = True
                for dataset_name in expected_datasets:
                    if dataset_name in consolidated_data:
                        # Normalize keys to strings for TM code mappings
                        data = consolidated_data[dataset_name]
                        # Check if data is empty dict and replace with defaults
                        if isinstance(data, dict) and len(data) == 0:
                            default_data = self._load_default_dataset_data(dataset_name)
                            self.datasets[dataset_name] = default_data
                            consolidated_data[dataset_name] = default_data
                            needs_save = True
                        elif dataset_name in ["fno_tm_code_with_tm_name", "mcx_tm_code_with_tm_name"]:
                            normalized_data = {str(k).strip(): v for k, v in data.items()}
                            self.datasets[dataset_name] = normalized_data
                        else:
                            self.datasets[dataset_name] = data
                    else:
                        # Dataset is missing, add default
                        default_data = self._load_default_dataset_data(dataset_name)
                        self.datasets[dataset_name] = default_data
                        consolidated_data[dataset_name] = default_data
                        needs_save = True
                
                # If we added missing datasets, save the file
                if needs_save:
                    with open(consolidated_path, "w") as fw:
                        json.dump(consolidated_data, fw, indent=4)
            else:
                # Create default consolidated data if file doesn't exist
                default_consolidated_data = {
                    "lotsize_data": self.default_lotsize_data.copy(),
                    "underlying_code_data": self.default_underlying_code_data.copy(),
                    "fund_filename_map": {"DBSBK0000033": {"Fund Names": {"Default": "DIF-Class 1 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000036": {"Fund Names": {"Default": "DIF-Class 2 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000038": {"Fund Names": {"Default": "DIF-Class 3 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000042": {"Fund Names": {"Default": "DIF-Class 5 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000044": {"Fund Names": {"Default": "DIF-Class 6 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000043": {"Fund Names": {"Default": "DIF-Class 7 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000049": {"Fund Names": {"Default": "DIF-Class 8 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000050": {"Fund Names": {"Default": "DIF-Class 9 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000051": {"Fund Names": {"Default": "DIF-Class 10 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000052": {"Fund Names": {"Default": "DIF-Class 11 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000074": {"Fund Names": {"Default": "DIF-Class 12 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000179": {"Fund Names": {"Default": "DIF-Class 13 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000189": {"Fund Names": {"Default": "DIF-Class 14 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000192": {"Fund Names": {"Default": "DIF-Class 15 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000214": {"Fund Names": {"Default": "DIF-Class 16 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000216": {"Fund Names": {"Default": "DIF-Class 17 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000217": {"Fund Names": {"Default": "DIF-Class 18_Moon"}, "Password": "AAGCD0792B"}, "DBSBK0000232": {"Fund Names": {"Default": "DIF-Class 19 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000247": {"Fund Names": {"CDS": "DIF-Class 21 CDS Holding", "Default": "DIF-Class 21 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000178": {"Fund Names": {"Default": "DGF-Cell 8"}, "Password": "AAICD1968M"}, "BNPP00000458": {"Fund Names": {"Default": "DGF-Cell 9"}, "Password": "AAICD2891H"}, "DGF-Cell 10": {"Fund Names": {"Default": "DGF-Cell 10"}, "Password": "AAICD3412C"}, "BNPP00000480": {"Fund Names": {"Default": "DGF-Cell 11"}, "Password": "AAICD6359G"}, "BNPP00000488": {"Fund Names": {"Default": "DGF-Cell 13"}, "Password": "AAICD7821M"}, "BNPP00000540": {"Fund Names": {"Default": "DGF-Cell 16"}, "Password": "AAJCD5624K"}, "BNPP00000535": {"Fund Names": {"Default": "DGF-Cell 17"}, "Password": "AAJCD4991K"}, "DBSBK0000229": {"Fund Names": {"Default": "DGF-Cell 18"}, "Password": "AAJCD6205G"}, "DBSBK0000228": {"Fund Names": {"Default": "DGF-Cell 19"}, "Password": "AAJCD6049E"}, "DBSBK0000285": {"Fund Names": {"CDS": "DGF-Cell 23 CDS Holding", "Default": "DGF-Cell 23 Holding"}, "Password": "AAKCD6244Q"}, "DBSBK0000299": {"Fund Names": {"Default": "DGF-Cell 24 Holding"}, "Password": "AAKCD7324B"}, "DBSBK0000353": {"Fund Names": {"Default": "DGF-Cell 28 Holding"}, "Password": "AALCD1140J"}, "DBSBK0000354": {"Fund Names": {"Default": "DGF-Cell 29 Holding"}, "Password": "AALCD1141K"}, "DBSBK0000380": {"Fund Names": {"Default": "DGF-Cell 32 Holding"}, "Password": ""}, "DBSBK0000356": {"Fund Names": {"Default": "GlobalQ_AIF-III"}, "Password": ""}, "DGF-Cell 36": {"Fund Names": {"Default": "DGF-Cell 36"}, "Password": ""}, "DGF-Cell 38": {"Fund Names": {"Default": "DGF-Cell 38"}, "Password": ""}},
                    "asio_sf_2_trade_loader": {
                        "RecordAction": "InsertUpdate",
                        "KeyValue.KeyName": "",
                        "UserTranId1": "",
                        "Portfolio": "ASIO - SF 2_INR",
                        "LocationAccount": "Asio_Sub Fund_2_OHM_FO_KOTBK0001479",
                        "Broker": "",
                        "PriceDenomination": "CALC",
                        "CounterInvestment": "INR",
                        "NetInvestmentAmount": "CALC",
                        "NetCounterAmount": "CALC",
                        "tradeFX": "",
                        "ContractFxRateNumerator": "",
                        "ContractFxRateDenominator": "",
                        "ContractFxRate": "",
                        "NotionalAmount": "",
                        "FundStructure2": "CALC",
                        "SpotDate": "",
                        "PriceDirectly": "",
                        "CounterFXDenomination": "CALC",
                        "CounterTDateFx": "",
                        "AccruedInterest": "",
                        "InvestmentAccruedInterest": "",
                        "TradeExpenses.ExpenseNumber": "",
                        "TradeExpenses.ExpenseCode": "",
                        "TradeExpenses.ExpenseAmt": "",
                        "NonCapExpenses.NonCapNumber": "",
                        "NonCapExpenses.NonCapExpenseCode": "",
                        "NonCapExpenses.NonCapAmount": "",
                        "NonCapExpenses.NonCapCurrency": "",
                        "NonCapExpenses.LocationAccount": "",
                            "NonCapExpenses.NonCapLiabilityCode": "",
                            "NonCapExpenses.NonCapPaymentType": ""
                        },
                        "asio_sf_2_option_security": {
                            "Exchange": "NSE",
                            "Issuer": "",
                            "Ticker": "",
                            "Cusip": "",
                            "Sedol": "",
                            "Isin": "",
                            "Riccode": "",
                            "AltKey1": "",
                            "AltKey2": "",
                            "BloombergID": "",
                            "BloombergTicker": "",
                            "BloombergUniqueID": "",
                            "BloombergMarketSector": "",
                            "SettleDays": 0,
                            "PricingFactor": 1,
                            "TradingFactor": 1,
                            "CurrentPriceDayRange": 1,
                            "AssetType": "OP",
                            "InvestmentType": "IXOP",
                            "PriceDenomination": "",
                            "BifurcationCurrency": "INR",
                            "PrincipalCurrency": "INR",
                            "IncomeCurrency": "INR",
                            "RiskCurrency": "INR",
                            "IssueCountry": "IN",
                            "WithholdingTaxType": "Standard",
                            "QDIEligibilityFlag": "Not Eligible",
                            "SharesOutstanding": "",
                            "Beta": "",
                            "SubIndustry": "",
                            "SubIndustry2": "",
                            "NonDeliverableCurrencyFlag": "",
                            "QuantityPrecision": "",
                            "InvestmentCrossZero": "",
                            "PriceByPreference": "",
                            "ForwardPriceInterpolateFlag": "",
                            "SecFeeSchedule": "",
                            "SecEligibleFlag": "",
                            "WashSalesEligible": "",
                            "ExerciseStyle": "European",
                            "PricingPrecision": 15,
                            "CashSettledFlag": 1
                        },
                        "asio_sf_2_mcx_future_security": {
                            "Exchange": "MCX",
                            "Issuer": "",
                            "Ticker": "",
                            "Cusip": "",
                            "Sedol": "",
                            "Isin": "",
                            "AltKey1": "",
                            "AltKey2": "",
                            "BloombergID": "",
                            "BloombergTicker": "",
                            "BloombergUniqueID": "",
                            "BloombergMarketSector": "",
                            "SettleDays": "",
                            "PricingFactor": "",
                            "CurrentPriceDayRange": 1,
                            "AssetType": "FT",
                            "InvestmentType": "CMFT",
                            "PriceDenomination": "INR",
                            "BifurcationCurrency": "INR",
                            "PrincipalCurrency": "INR",
                            "IncomeCurrency": "INR",
                            "RiskCurrency": "INR",
                            "IssueCountry": "IN",
                            "WithholdingTaxType": "Standard",
                            "QDIEligibilityFlag": "Not Eligible",
                            "SharesOutstanding": "",
                            "SubIndustry": "",
                            "SubIndustry2": "",
                            "QuantityPrecision": 3,
                            "InvestmentCrossZero": "Use Accounting Parameters",
                            "PriceByPreference": "Currency",
                            "ForwardPriceInterpolateFlag": 0,
                            "PricingPrecision": 3,
                            "FirstMarkDate": "01-01-2022",
                            "LastMarkDate": "",
                            "PriceList": "",
                            "AutoGenerateMarks": 1,
                            "CashSettlement": 1
                        },
                        "asio_sf_2_mcx_trade_loader": {
                            "RecordAction": "InsertUpdate",
                            "KeyValue.KeyName": "",
                            "UserTranId1": "",
                            "Portfolio": "ASIO - SF 2_INR",
                            "LocationAccount": "Asio_Sub Fund_2_OHM_MCX_KOTBK0001479",
                            "Broker": "",
                            "PriceDenomination": "CALC",
                            "CounterInvestment": "INR",
                            "NetInvestmentAmount": "CALC",
                            "NetCounterAmount": "CALC",
                            "tradeFX": "",
                            "ContractFxRateNumerator": "",
                            "ContractFxRateDenominator": "",
                            "ContractFxRate": "",
                            "NotionalAmount": "",
                            "FundStructure2": "CALC",
                            "SpotDate": "",
                            "PriceDirectly": "",
                            "CounterFXDenomination": "CALC",
                            "CounterTDateFx": "",
                            "AccruedInterest": "",
                            "InvestmentAccruedInterest": "",
                            "TradeExpenses.ExpenseNumber": "",
                            "TradeExpenses.ExpenseCode": "",
                            "TradeExpenses.ExpenseAmt": "",
                            "NonCapExpenses.NonCapNumber": "",
                            "NonCapExpenses.NonCapExpenseCode": "",
                            "NonCapExpenses.NonCapAmount": "",
                            "NonCapExpenses.NonCapCurrency": "",
                            "NonCapExpenses.LocationAccount": "",
                            "NonCapExpenses.NonCapLiabilityCode": "",
                            "NonCapExpenses.NonCapPaymentType": ""
                        },
                        "asio_pricing_fno": {
                            "RecordAction": "InsertUpdate",
                            "PriceDenomination": "INR",
                            "PriceList": "NSE_Equity",
                            "TaxLotID": ""
                        },
                        "asio_pricing_mcx": {
                            "RecordAction": "InsertUpdate",
                            "PriceDenomination": "INR",
                            "PriceList": "NSE_Equity",
                            "TaxLotID": ""
                        },
                        "asio_sf_2_mcx_option_security": {
                            "Exchange": "MCX",
                            "Issuer": "",
                            "Ticker": "",
                            "Cusip": "",
                            "Sedol": "",
                            "Isin": "",
                            "Riccode": "",
                            "AltKey1": "",
                            "AltKey2": "",
                            "BloombergID": "",
                            "BloombergTicker": "",
                            "BloombergUniqueID": "",
                            "BloombergMarketSector": "",
                            "SettleDays": 0,
                            "PricingFactor": 1,
                            "CurrentPriceDayRange": 1,
                            "AssetType": "OP",
                            "InvestmentType": "CMOP",
                            "PriceDenomination": "",
                            "BifurcationCurrency": "INR",
                            "PrincipalCurrency": "INR",
                            "IncomeCurrency": "INR",
                            "RiskCurrency": "INR",
                            "IssueCountry": "IN",
                            "WithholdingTaxType": "Standard",
                            "QDIEligibilityFlag": "Not Eligible",
                            "SharesOutstanding": "",
                            "Beta": "",
                            "SubIndustry": "",
                            "SubIndustry2": "",
                            "NonDeliverableCurrencyFlag": "",
                            "QuantityPrecision": "",
                            "InvestmentCrossZero": "",
                            "PriceByPreference": "",
                            "ForwardPriceInterpolateFlag": "",
                            "SecFeeSchedule": "",
                            "SecEligibleFlag": "",
                            "WashSalesEligible": "",
                            "ExerciseStyle": "European",
                            "PricingPrecision": 15,
                            "CashSettledFlag": 1
                        },
                    "asio_recon_portfolio_mapping": {
                        "ASIO - SF 3": "THE ASIO FUND VCC - SUB-FUND 3",
                        "ASIO - SF 8_Golden": "THE ASIO FUND VCC - EMKAY BHARAT FUND - BHARATS GOLDEN DECADE",
                        "ASIO-SF1_": "THE ASIO FUND VCC - FORT PANGEA",
                        "ASIO - SF 8_Capital Builder": "THE ASIO FUND VCC - EMKAY BHARAT FUND",
                        "ASIO - SF 8_Envi": "THE ASIO FUND VCC - EMKAY BHARAT FUND - ENVI",
                        "ASIO - SF 4": "THE ASIO FUND VCC - SUB FUND 4",
                        "ASIO - SF 10": "THE ASIO FUND VCC-HATHI CAPITAL INDIA GROWTH FUND I",
                        "ASIO - SF 6": "THE ASIO FUND VCC - FRAGRANT HARBOUR INDIA OPPORTUNITIES FUND",
                        "ASIO - SF 2": "",
                        "ASIO - SF 7": ""
                    },
                    "asio_recon_format_1_headers": {
                        "CLN_CODE": "Cln Code",
                        "CLN_NAME": "Cln Name",
                        "INSTR_CODE": "Instr Code",
                        "INSTR_ISIN": "Instr ISIN",
                        "INSTR_NAME": "Instr Name",
                        "DEMAT_PHYSICAL": "Demat / Physical",
                        "SETTLED_POSITION": "Settled Position",
                        "PENDING_PURCHASE": "Pending Purchase",
                        "PENDING_SALE": "Pending Sale",
                        "BLOCKED": "Blocked",
                        "BLOCKED_RECEIVABLES": "Blocked Receivables",
                        "BLOCKED_PAYABLES": "Blocked Payables",
                        "PENDING_CA_ENTITLEMENTS": "Pending CA Entitlements",
                        "MARKET_PRICE_DATE": "Market Price Date",
                        "MARKET_PRICE": "Market Price",
                        "SETTLED_BLOCKED_MARKET_PRICE": "(Settled + Blocked * Market Price)",
                        "SALEABLE": "Saleable",
                        "CONTRACTUAL": "Contractual"
                    },
                    "asio_recon_format_2_headers": {
                        "VALUE_DATE_AS_AT": "Value date as at",
                        "SERVICE_LOCATION": "Service location",
                        "SECURITIES_ACCOUNT_NAME": "Securities account name",
                        "SECURITIES_ACCOUNT_NUMBER": "Securities account number",
                        "SECURITY_TYPE": "Security type",
                        "SECURITY_NAME": "Security name",
                        "ISIN": "ISIN",
                        "PLACE_OF_SETTLEMENT": "Place of settlement",
                        "SETTLED_BALANCE": "Settled balance",
                        "ANTICIPATED_DEBITS": "Anticipated debits",
                        "AVAILABLE_BALANCE": "Available balance",
                        "ANTICIPATED_CREDITS": "Anticipated credits",
                        "TRADED_BALANCE": "Traded balance",
                        "SECURITY_CURRENCY": "Security currency",
                        "SECURITY_PRICE": "Security price",
                        "INDICATIVE_SETTLED_VALUE": "Indicative settled value",
                        "INDICATIVE_TRADED_VALUE": "Indicative traded value",
                        "INDICATIVE_SETTLED_VALUE_INR": "Indicative settled value in INR",
                        "INDICATIVE_TRADED_VALUE_INR": "Indicative traded value in INR"
                    },
                    "asio_recon_bhavcopy_headers": {
                        "TRADDT": "TradDt",
                        "BIZDT": "BizDt",
                        "SGMT": "Sgmt",
                        "SRC": "Src",
                        "FININSTRMTP": "FinInstrmTp",
                        "FININSTRMID": "FinInstrmId",
                        "ISIN": "ISIN",
                        "TCKRSYMB": "TckrSymb",
                        "SCTYSRS": "SctySrs",
                        "XPRYDT": "XpryDt",
                        "FININSTRMACTLXPRYDT": "FininstrmActlXpryDt",
                        "STRKPRIC": "StrkPric",
                        "OPTNTP": "OptnTp",
                        "FININSTRMnm": "FinInstrmNm",
                        "OPNPRIC": "OpnPric",
                        "HGHPRIC": "HghPric",
                        "LWPRIC": "LwPric",
                        "CLSPRIC": "ClsPric",
                        "LASTPRIC": "LastPric"
                    },
                    "asio_geneva_headers": {
                        "PORTFOLIO": "Portfolio",
                        "INVESTMENT": "Investment",
                        "INVESTMENT_DESCRIPTION": "Investment Description",
                        "TRADED_QUANTITY": "Traded Quantity",
                        "SETTLED_QUANTITY": "Settled Quantity",
                        "CURRENCY": "Currency",
                        "UNIT_COST": "Unit Cost",
                        "COST_LOCAL": "Cost Local",
                        "COST_BOOK": "Cost Book",
                        "MARKET_PRICE_LOCAL": "Market Price Local",
                        "MARKET_VALUE_LOCAL": "Market Value Local"
                    },
                    "trade_headers": {
                        "CLIENT_CODE": "Client\nCode",
                        "EXCHANGE": "Exchange",
                        "MKT_TYPE": "Mkt\nType",
                        "DATE": "Date",
                        "TRADE_TIME": "TradeTime",
                        "SETL_NO": "SetlNo",
                        "SCRIP_CODE": "Scrip Code",
                        "SCRIP_NAME": "Scrip Name",
                        "ISIN": "ISIN",
                        "BUY_QTY": "Buy Qty",
                        "BUY_RATE": "Buy Rate",
                        "BUY_NET_RATE": "Buy\nNet Rate",
                        "BUY_VALUE": "Buy Value",
                        "SELL_QTY": "Sell Qty",
                        "BALANCE": "Balance",
                        "SELL_RATE": "Sell Rate",
                        "SELL_NET_RATE": "Sell\nNet Rate",
                        "SELL_VALUE": "Sell Value",
                        "BREAK_EVEN": "BreakEven",
                        "MARKET_RATE": "Market Rate",
                        "TERMINAL": "Terminal",
                        "DELY": "Dely"
                    },
                    "aafspl_car_future": {
                        "Exchange": "MCX",
                        "Issuer": "",
                        "Ticker": "",
                        "Cusip": "",
                        "Sedol": "",
                        "Isin": "",
                        "AltKey1": "",
                        "AltKey2": "",
                        "BloombergID": "",
                        "BloombergTicker": "",
                        "BloombergUniqueID": "",
                        "BloombergMarketSector": "",
                        "SettleDays": "",
                        "PricingFactor": "",
                        "CurrentPriceDayRange": 1,
                        "AssetType": "FT",
                        "InvestmentType": "CMFT",
                        "PriceDenomination": "INR",
                        "BifurcationCurrency": "INR",
                        "PrincipalCurrency": "INR",
                        "IncomeCurrency": "INR",
                        "RiskCurrency": "INR",
                        "IssueCountry": "IN",
                        "WithholdingTaxType": "Standard",
                        "QDIEligibilityFlag": "Not Eligible",
                        "SharesOutstanding": "",
                        "SubIndustry": "",
                        "SubIndustry2": "",
                        "QuantityPrecision": 3,
                        "InvestmentCrossZero": "Use Accounting Parameters",
                        "PriceByPreference": "Currency",
                        "ForwardPriceInterpolateFlag": 0,
                        "PricingPrecision": 3,
                        "FirstMarkDate": "01-01-2022",
                        "LastMarkDate": "",
                        "PriceList": "",
                        "AutoGenerateMarks": 1,
                        "CashSettlement": 1
                    },
                    "option_security": {
                        "Exchange": "MCX",
                        "Issuer": "",
                        "Ticker": "",
                        "Cusip": "",
                        "Sedol": "",
                        "Isin": "",
                        "Riccode": "",
                        "AltKey1": "",
                        "AltKey2": "",
                        "BloombergID": "",
                        "BloombergTicker": "",
                        "BloombergUniqueID": "",
                        "BloombergMarketSector": "",
                        "SettleDays": 0,
                        "PricingFactor": 1,
                        "CurrentPriceDayRange": 0,
                        "AssetType": "OP",
                        "InvestmentType": "CMFTOP",
                        "PriceDenomination": "",
                        "BifurcationCurrency": "INR",
                        "PrincipalCurrency": "INR",
                        "IncomeCurrency": "INR",
                        "RiskCurrency": "INR",
                        "IssueCountry": "IN",
                        "WithholdingTaxType": "Standard",
                        "QDIEligibilityFlag": "Not Eligible",
                        "SharesOutstanding": "",
                        "Beta": "",
                        "SubIndustry": "",
                        "SubIndustry2": "",
                        "NonDeliverableCurrencyFlag": "",
                        "QuantityPrecision": "",
                        "InvestmentCrossZero": "",
                        "PriceByPreference": "",
                        "ForwardPriceInterpolateFlag": "",
                        "SecFeeSchedule": "",
                        "SecEligibleFlag": "",
                        "WashSalesEligible": "",
                        "ExerciseStyle": "European",
                        "PricingPrecision": 15,
                        "CashSettledFlag": 1
                    },
                    "asio_sf_2_option_security": {
                        "Exchange": "NSE",
                        "Issuer": "",
                        "Ticker": "",
                        "Cusip": "",
                        "Sedol": "",
                        "Isin": "",
                        "AltKey1": "",
                        "AltKey2": "",
                        "BloombergID": "",
                        "BloombergTicker": "",
                        "BloombergUniqueID": "",
                        "BloombergMarketSector": "",
                        "SettleDays": "",
                        "PricingFactor": "",
                        "TradingFactor": 1,
                        "CurrentPriceDayRange": 1,
                        "AssetType": "FT",
                        "InvestmentType": "EQFT",
                        "PriceDenomination": "INR",
                        "BifurcationCurrency": "INR",
                        "PrincipalCurrency": "INR",
                        "IncomeCurrency": "INR",
                        "RiskCurrency": "INR",
                        "IssueCountry": "IN",
                        "WithholdingTaxType": "Standard",
                        "QDIEligibilityFlag": "Not Eligible",
                        "SharesOutstanding": "",
                        "SubIndustry": "",
                        "SubIndustry2": "",
                        "QuantityPrecision": 3,
                        "InvestmentCrossZero": "Use Accounting Parameters",
                        "PriceByPreference": "Currency",
                        "ForwardPriceInterpolateFlag": 0,
                        "PricingPrecision": 3,
                        "FirstMarkDate": "01-01-2022",
                        "LastMarkDate": "",
                        "PriceList": "",
                        "AutoGenerateMarks": 1,
                        "CashSettlement": 1
                    },
                    "asio_sf_2_future_security": {
                        "Exchange": "NSE",
                        "Issuer": "",
                        "Ticker": "",
                        "Cusip": "",
                        "Sedol": "",
                        "Isin": "",
                        "AltKey1": "",
                        "AltKey2": "",
                        "BloombergID": "",
                        "BloombergTicker": "",
                        "BloombergUniqueID": "",
                        "BloombergMarketSector": "",
                        "SettleDays": "",
                        "PricingFactor": "",
                        "TradingFactor": 1,
                        "CurrentPriceDayRange": 1,
                        "AssetType": "FT",
                        "InvestmentType": "EQFT",
                        "PriceDenomination": "INR",
                        "BifurcationCurrency": "INR",
                        "PrincipalCurrency": "INR",
                        "IncomeCurrency": "INR",
                        "RiskCurrency": "INR",
                        "IssueCountry": "IN",
                        "WithholdingTaxType": "Standard",
                        "QDIEligibilityFlag": "Not Eligible",
                        "SharesOutstanding": "",
                        "SubIndustry": "",
                        "SubIndustry2": "",
                        "QuantityPrecision": 3,
                        "InvestmentCrossZero": "Use Accounting Parameters",
                        "PriceByPreference": "Currency",
                        "ForwardPriceInterpolateFlag": 0,
                        "PricingPrecision": 3,
                        "FirstMarkDate": "01-01-2022",
                        "LastMarkDate": "",
                        "PriceList": "",
                        "AutoGenerateMarks": 1,
                        "CashSettlement": 1
                    },
                    "car_trade_loader": {
                        "RecordAction": "InsertUpdate",
                        "KeyValue.KeyName": "",
                        "UserTranId1": "",
                        "Portfolio": "AAFSPL_CAR",
                        "LocationAccount": "AAFSPL_CAR_Ventura Securities Ltd",
                        "Strategy": "Default",
                        "Broker": "",
                        "PriceDenomination": "CALC",
                        "CounterInvestment": "INR",
                        "NetInvestmentAmount": "CALC",
                        "NetCounterAmount": "CALC",
                        "TradeFX": "",
                        "ContractFxRateNumerator": "",
                        "ContractFxRateDenominator": "",
                        "ContractFxRate": "",
                        "NotionalAmount": "",
                        "FundStructure": "CALC",
                        "SpotDate": "",
                        "PriceDirectly": "",
                        "CounterFXDenomination": "CALC",
                        "CounterTDateFx": "",
                        "AccruedInterest": "",
                        "InvestmentAccruedInterest": "",
                        "TradeExpenses.ExpenseNumber": "",
                        "TradeExpenses.ExpenseCode": "",
                        "TradeExpenses.ExpenseAmt": "",
                        "NonCapExpenses.NonCapNumber": "",
                        "NonCapExpenses.NonCapExpenseCode": "",
                        "NonCapExpenses.NonCapAmount": "",
                        "NonCapExpenses.NonCapCurrency": "",
                        "NonCapExpenses.LocationAccount": "",
                        "NonCapExpenses.NonCapLiabilityCode": "",
                        "NonCapExpenses.NonCapPaymentType": ""
                    },
                    "asio_sf_2_option_security": {
                        "Exchange": "NSE",
                        "Issuer": "",
                        "Ticker": "",
                        "Cusip": "",
                        "Sedol": "",
                        "Isin": "",
                        "Riccode": "",
                        "AltKey1": "",
                        "AltKey2": "",
                        "BloombergID": "",
                        "BloombergTicker": "",
                        "BloombergUniqueID": "",
                        "BloombergMarketSector": "",
                        "SettleDays": 0,
                        "PricingFactor": 1,
                        "TradingFactor": 1,
                        "CurrentPriceDayRange": 1,
                        "AssetType": "OP",
                        "InvestmentType": "IXOP",
                        "PriceDenomination": "",
                        "BifurcationCurrency": "INR",
                        "PrincipalCurrency": "INR",
                        "IncomeCurrency": "INR",
                        "RiskCurrency": "INR",
                        "IssueCountry": "IN",
                        "WithholdingTaxType": "Standard",
                        "QDIEligibilityFlag": "Not Eligible",
                        "SharesOutstanding": "",
                        "Beta": "",
                        "SubIndustry": "",
                        "SubIndustry2": "",
                        "NonDeliverableCurrencyFlag": "",
                        "QuantityPrecision": "",
                        "InvestmentCrossZero": "",
                        "PriceByPreference": "",
                        "ForwardPriceInterpolateFlag": "",
                        "SecFeeSchedule": "",
                        "SecEligibleFlag": "",
                        "WashSalesEligible": "",
                        "ExerciseStyle": "European",
                        "PricingPrecision": 15,
                        "CashSettledFlag": 1
                    },
                    "geneva_custodian_mapping": {
                        "DIF-Class 5_Heer": "DOVETAIL INDIA FUND CLASS 5 SHARES",
                        "DIF-Class 5_Moonrock_F&O": "DOVETAIL INDIA FUND CLASS 5 SHARES",
                        "DIF-Class 3_Saankhya": "DOVETAIL INDIA FUND CLASS 3 Shares",
                        "DIF-Class 11_Eagle": "DOVETAIL INDIA FUND CLASS 11 SHARES",
                        "DGF-Cell 2": "DOVETAIL GLOBAL FUND PCC CELL 2",
                        "DIF-Class 1": "DOVETAIL INDIA FUND",
                        "DGF-Cell 23": "DOVETAIL GLOBAL FUND PCC - CELL 23",
                        "DIF-Class 6_Amal Parikh": "DOVETAIL INDIA FUND CLASS 6 SHARES",
                        "DIF-Class 21": "DOVETAIL INDIA FUND CLASS 21",
                        "DGF-Cell 28": "DOVETAIL GLOBAL FUND PCC - CELL 28",
                        "DGF-Cell 29": "DOVETAIL GLOBAL FUND PCC CELL 29",
                        "DIF-Class 2_Agile": "DOVETAIL INDIA FUND CLASS 2 SHARES",
                        "DIF-Class 19": "DOVETAIL INDIA FUND CLASS 19",
                        "ASIO - SF 2": "THE ASIO FUND VCC - SUB-FUND 2",
                        "ASIO - SF 3": "THE ASIO FUND VCC - SUB-FUND 3",
                        "ASIO - SF 4": "THE ASIO FUND VCC - SUB-FUND 4",
                        "ASIO - SF 7": "THE ASIO FUND VCC FORT CAPITAL ABSOLUTE ALTERNATES",
                        "DIF-Class 18_Moon": "DOVETAIL INDIA FUND CLASS 18",
                        "DIF-Class 12_F&O": "DOVETAIL INDIA FUND CLASS 12",
                        "DIF-Class 12_Moon New F&O": "DOVETAIL INDIA FUND CLASS 12",
                        "DIF-Class 6": "DOVETAIL INDIA FUND CLASS 6 SHARES",
                        "DIF-Class 8_Discipline": "DOVETAIL INDIA FUND CLASS 8 SHARES",
                        "GlobalQ_AIF-III": "GLOBAL Q FUND",
                        "DGF-Cell 32": "DOVETAIL GLOBAL FUND PCC - CELL 32",
                        "DIF-Class 8_Panda": "DOVETAIL INDIA FUND CLASS 8 SHARES",
                        "DGF- Cell 36": "",
                        "DGF- Cell 38": "",
                        "DIF-Class 8_HBE Capital": "DOVETAIL INDIA FUND CLASS 8 SHARES"
                    },
                    "fno_tm_code_with_tm_name": {
                        "13302": "Achintya Securities Pvt. Ltd.",
                        "10412": "Motilal Oswal Financial Services Limited",
                        "07536": "Trustline Securities Limited",
                        "07714": "SMC Global Securities Limited"
                    },
                    "mcx_tm_code_with_tm_name": {
                        "31640": "Achintya Securities Pvt. Ltd.",
                        "10515": "SMC Global Securities Limited"
                    }
                }
                
                # Save the default consolidated data file
                with open(consolidated_path, "w") as f:
                    json.dump(default_consolidated_data, f, indent=4)
                
                # Load the data from the newly created file
                self.lotsize_data.update(default_consolidated_data["lotsize_data"])
                underlying_code_from_file = default_consolidated_data.get("underlying_code_data")
                if underlying_code_from_file and isinstance(underlying_code_from_file, dict) and len(underlying_code_from_file) > 0:
                    self.underlying_code_data = underlying_code_from_file.copy()
                else:
                    self.underlying_code_data = self.default_underlying_code_data.copy()
                for dataset_name in ["fund_filename_map", "asio_recon_portfolio_mapping", "asio_recon_format_1_headers", "asio_recon_format_2_headers", 
                                    "asio_recon_bhavcopy_headers", "asio_geneva_headers", "trade_headers", 
                                    "aafspl_car_future", "option_security", "car_trade_loader", "asio_sf_2_trade_loader", "asio_sf_2_option_security", "asio_sf_2_future_security", "asio_sf_2_mcx_trade_loader", "asio_sf_2_mcx_option_security", "geneva_custodian_mapping", "fno_tm_code_with_tm_name", "mcx_tm_code_with_tm_name", "asio_sf4_ft", "asio_sf4_trading_code_mapping", "asio_sub_fund4_read_config", "mcx_group2_filters", "fno_group2_filters"]:
                    if dataset_name in default_consolidated_data:
                        data = default_consolidated_data[dataset_name]
                        # Normalize keys to strings for TM code mappings
                        if dataset_name in ["fno_tm_code_with_tm_name", "mcx_tm_code_with_tm_name"]:
                            self.datasets[dataset_name] = {str(k).strip(): v for k, v in data.items()}
                        else:
                            self.datasets[dataset_name] = data
            # Keep alias in sync for current dataset
            self.header_data = self.datasets.get(self.current_dataset_name, {})
        except Exception as e:
            print(f"Could not load saved data: {e}")

    def load_data(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        if self.mode_var.get() == "LOTSIZE":
            for symbol, lotsize in self.lotsize_data.items():
                self.tree.insert("", "end", values=(symbol, lotsize))
        elif self.mode_var.get() == "UNDERLYINGCODE":
            for symbol, code in self.underlying_code_data.items():
                self.tree.insert("", "end", values=(symbol, code))
        elif self.current_dataset_name == "fund_filename_map":
            # User-friendly display for fund_filename_map
            for fund_code, fund_data in self.header_data.items():
                if isinstance(fund_data, dict):
                    fund_names = fund_data.get("Fund Names", {})
                    default_name = fund_names.get("Default", "") if isinstance(fund_names, dict) else ""
                    cds_name = fund_names.get("CDS", "") if isinstance(fund_names, dict) else ""
                    password = fund_data.get("Password", "")
                    self.tree.insert("", "end", values=(fund_code, default_name, cds_name, password))
                else:
                    self.tree.insert("", "end", values=(fund_code, "", "", ""))
        elif self.current_dataset_name in ["mcx_group2_filters", "fno_group2_filters"]:
            # Handle list-type datasets (filters)
            if isinstance(self.header_data, list):
                for index, filter_value in enumerate(self.header_data):
                    self.tree.insert("", "end", values=(str(index + 1), filter_value))
            else:
                # Fallback if data is not a list
                self.tree.insert("", "end", values=("1", str(self.header_data)))
        else:
            # Handle dictionary-type datasets
            if isinstance(self.header_data, dict):
                for header, value in self.header_data.items():
                    display_value = json.dumps(value) if isinstance(value, (dict, list)) else value
                    # Ensure header is always inserted as string to preserve leading zeros
                    header_str = str(header) if header is not None else ""
                    self.tree.insert("", "end", values=(header_str, display_value))
            else:
                # Fallback for non-dict, non-list data
                self.tree.insert("", "end", values=("Value", str(self.header_data)))
        self.edit_btn.config(state="disabled", bg="#bdc3c7", fg="#7f8c8d", relief="flat", bd=1, font=("Arial", 12))
        self.delete_btn.config(state="disabled", bg="#bdc3c7", fg="#7f8c8d", relief="flat", bd=1, font=("Arial", 12))
        self.selected_info.config(text="Select a row to edit or delete", fg="#7f8c8d", font=("Arial", 10))

    def filter_data(self, *args):
        term = self.search_var.get().lower()
        for item in self.tree.get_children():
            self.tree.delete(item)
        if self.mode_var.get() == "LOTSIZE":
            for symbol, lotsize in self.lotsize_data.items():
                if term in symbol.lower():
                    self.tree.insert("", "end", values=(symbol, lotsize))
        elif self.mode_var.get() == "UNDERLYINGCODE":
            for symbol, code in self.underlying_code_data.items():
                if term in symbol.lower():
                    self.tree.insert("", "end", values=(symbol, code))
        elif self.current_dataset_name == "fund_filename_map":
            # User-friendly filter for fund_filename_map
            for fund_code, fund_data in self.header_data.items():
                if isinstance(fund_data, dict):
                    fund_names = fund_data.get("Fund Names", {})
                    default_name = fund_names.get("Default", "") if isinstance(fund_names, dict) else ""
                    cds_name = fund_names.get("CDS", "") if isinstance(fund_names, dict) else ""
                    password = fund_data.get("Password", "")
                    # Search in custodian account, fund name, or CDS name
                    if term in fund_code.lower() or term in default_name.lower() or term in cds_name.lower():
                        self.tree.insert("", "end", values=(fund_code, default_name, cds_name, password))
        elif self.current_dataset_name in ["mcx_group2_filters", "fno_group2_filters"]:
            # Handle list-type datasets (filters) - filter by value
            if isinstance(self.header_data, list):
                for index, filter_value in enumerate(self.header_data):
                    if term in str(filter_value).lower():
                        self.tree.insert("", "end", values=(str(index + 1), filter_value))
        else:
            # Handle dictionary-type datasets
            if isinstance(self.header_data, dict):
                for header, value in self.header_data.items():
                    if term in header.lower():
                        display_value = json.dumps(value) if isinstance(value, (dict, list)) else value
                        self.tree.insert("", "end", values=(header, display_value))
        self.edit_btn.config(state="disabled", bg="#bdc3c7", fg="#7f8c8d", relief="flat", bd=1, font=("Arial", 12))
        self.delete_btn.config(state="disabled", bg="#bdc3c7", fg="#7f8c8d", relief="flat", bd=1, font=("Arial", 12))
        self.selected_info.config(text="Select a row to edit or delete", fg="#7f8c8d", font=("Arial", 10))

    def on_selection_change(self, event):
        selection = self.tree.selection()
        if selection:
            self.edit_btn.config(state="normal", bg="#2980b9", fg="white", relief="raised", bd=2, font=("Arial", 12, "bold"))
            self.delete_btn.config(state="normal", bg="#c0392b", fg="white", relief="raised", bd=2, font=("Arial", 12, "bold"))
            item = self.tree.item(selection[0])
            if self.mode_var.get() == "LOTSIZE":
                symbol, lotsize = item["values"]
                self.selected_info.config(text=f"Selected: {symbol} (Lotsize: {lotsize})", fg="#2c3e50", font=("Arial", 10, "bold"))
            elif self.mode_var.get() == "UNDERLYINGCODE":
                symbol, code = item["values"]
                self.selected_info.config(text=f"Selected: {symbol} (Code: {code})", fg="#2c3e50", font=("Arial", 10, "bold"))
            else:
                values = item["values"]
                if self.current_dataset_name == "fund_filename_map":
                    fund_code, fund_name = values[0], values[1] if len(values) > 1 else ""
                    self.selected_info.config(text=f"Selected: {fund_code} - {fund_name}", fg="#2c3e50", font=("Arial", 10, "bold"))
                elif self.current_dataset_name == "asio_recon_portfolio_mapping":
                    header, val = values[0], values[1] if len(values) > 1 else ""
                    self.selected_info.config(text=f"Selected: Portfolio: {header} (Holding Name: {val})", fg="#2c3e50", font=("Arial", 10, "bold"))
                elif self.current_dataset_name == "geneva_custodian_mapping":
                    header, val = values[0], values[1] if len(values) > 1 else ""
                    self.selected_info.config(text=f"Selected: Geneva: {header} | Custodian: {val}", fg="#2c3e50", font=("Arial", 10, "bold"))
                elif self.current_dataset_name == "fno_tm_code_with_tm_name":
                    header, val = values[0], values[1] if len(values) > 1 else ""
                    self.selected_info.config(text=f"Selected: TM Code: {header} | TM Name: {val}", fg="#2c3e50", font=("Arial", 10, "bold"))
                elif self.current_dataset_name == "mcx_tm_code_with_tm_name":
                    header, val = values[0], values[1] if len(values) > 1 else ""
                    self.selected_info.config(text=f"Selected: TM Code: {header} | TM Name: {val}", fg="#2c3e50", font=("Arial", 10, "bold"))
                else:
                    header, val = values[0], values[1] if len(values) > 1 else ""
                    self.selected_info.config(text=f"Selected: Header: {header} (Value: {val})", fg="#2c3e50", font=("Arial", 10, "bold"))
        else:
            self.edit_btn.config(state="disabled", bg="#bdc3c7", fg="#7f8c8d", relief="flat", bd=1, font=("Arial", 12))
            self.delete_btn.config(state="disabled", bg="#bdc3c7", fg="#7f8c8d", relief="flat", bd=1, font=("Arial", 12))
            self.selected_info.config(text="Select a row to edit or delete", fg="#7f8c8d", font=("Arial", 10))

    def edit_selected_item(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select an item to edit")
            return
        item = self.tree.item(selection[0])
        if self.mode_var.get() == "LOTSIZE":
            symbol, lotsize = item["values"]
            self.show_edit_dialog(symbol, lotsize)
        elif self.mode_var.get() == "UNDERLYINGCODE":
            symbol, code = item["values"]
            self.show_edit_dialog(symbol, code, is_underlying_code=True)
        elif self.current_dataset_name == "fund_filename_map":
            values = item["values"]
            fund_code = values[0]
            fund_name = values[1] if len(values) > 1 else ""
            cds_name = values[2] if len(values) > 2 else ""
            password = values[3] if len(values) > 3 else ""
            self.show_fund_edit_dialog(fund_code, fund_name, cds_name, password)
        elif self.current_dataset_name in ["mcx_group2_filters", "fno_group2_filters"]:
            # Handle list-type datasets
            values = item["values"]
            index = int(values[0]) - 1  # Convert back to 0-based index
            filter_value = values[1] if len(values) > 1 else ""
            self.show_filter_edit_dialog(index, filter_value)
        else:
            header, value = item["values"]
            self.show_header_edit_dialog(header, value)

    def delete_selected_item(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select an item to delete")
            return
        item = self.tree.item(selection[0])
        if self.mode_var.get() == "LOTSIZE":
            symbol, _ = item["values"]
            if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete '{symbol}'?"):
                del self.lotsize_data[symbol]
                self.load_data()
                self.auto_save()
        elif self.current_dataset_name in ["mcx_group2_filters", "fno_group2_filters"]:
            # Handle list-type datasets
            values = item["values"]
            if not values or len(values) < 2:
                messagebox.showerror("Error", "Cannot delete: No data found in selected row")
                return
            index = int(values[0]) - 1  # Convert back to 0-based index
            filter_value = values[1]
            
            # Ensure dataset is a list
            if self.current_dataset_name not in self.datasets:
                self.datasets[self.current_dataset_name] = []
            if not isinstance(self.datasets[self.current_dataset_name], list):
                self.datasets[self.current_dataset_name] = []
            
            if 0 <= index < len(self.datasets[self.current_dataset_name]):
                if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete '{filter_value}'?"):
                    self.datasets[self.current_dataset_name].pop(index)
                    self.header_data = self.datasets[self.current_dataset_name]
                    self.load_data()
                    self.auto_save()
            else:
                messagebox.showerror("Error", "Invalid index for deletion")
        else:
            values = item["values"]
            if not values or len(values) == 0:
                messagebox.showerror("Error", "Cannot delete: No data found in selected row")
                return
            # Ensure we're working with the dataset dictionary directly first
            if self.current_dataset_name not in self.datasets:
                self.datasets[self.current_dataset_name] = {}
            dataset_dict = self.datasets[self.current_dataset_name]
            self.header_data = dataset_dict  # Point to the same dictionary
            
            # Get identifier from tree - but tree might convert "07536" to 7536
            # So we'll match by comparing with all dictionary keys
            identifier = values[0]  # First column is always the key
            if identifier is None:
                messagebox.showerror("Error", "Cannot delete: Invalid identifier")
                return
            
            # Convert identifier to string (but this might lose leading zeros if tree converted it)
            identifier_from_tree = str(identifier).strip()
            
            if not identifier_from_tree:
                messagebox.showerror("Error", "Cannot delete: Invalid identifier")
                return
            
            # Find the matching key in dictionary
            # Tree might convert "07536" to 7536, so we match by row position
            # since tree displays items in the same order as dictionary items
            found_key = None
            identifier_str_for_display = identifier_from_tree
            
            # Get the selected row index
            selected_item = selection[0]
            all_items = self.tree.get_children()
            try:
                row_index = all_items.index(selected_item)
            except ValueError:
                row_index = -1
            
            # Try exact match first
            for key in dataset_dict.keys():
                key_str = str(key).strip()
                if key_str == identifier_from_tree:
                    found_key = key
                    identifier_str_for_display = key_str
                    break
            
            # If not found by exact match and we have row index, match by position
            # This handles cases where tree converted "07536" to 7536
            if found_key is None and row_index >= 0:
                keys_list = list(dataset_dict.keys())
                if row_index < len(keys_list):
                    found_key = keys_list[row_index]
                    identifier_str_for_display = str(found_key).strip()
            
            if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete '{identifier_str_for_display}'?"):
                if found_key is not None:
                    del dataset_dict[found_key]
                    self.load_data()
                    self.auto_save()
                else:
                    # Debug: show what's actually in the data
                    available_keys = list(dataset_dict.keys())[:10]  # First 10 keys for debugging
                    messagebox.showerror("Error", f"Item '{identifier_from_tree}' not found in data.\n\nLooking for: '{identifier_from_tree}'\nAvailable keys: {', '.join(str(k) for k in available_keys) if available_keys else 'None'}")

    def add_new_item(self):
        if self.mode_var.get() == "LOTSIZE":
            self.show_edit_dialog()
        elif self.mode_var.get() == "UNDERLYINGCODE":
            self.show_edit_dialog(is_underlying_code=True)
        elif self.current_dataset_name == "fund_filename_map":
            self.show_fund_edit_dialog()
        elif self.current_dataset_name in ["mcx_group2_filters", "fno_group2_filters"]:
            # Handle list-type datasets - add new item
            self.show_filter_edit_dialog()
        else:
            self.show_header_edit_dialog()

    def edit_item(self, event=None):
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select an item to edit")
            return
        item = self.tree.item(selection[0])
        if self.mode_var.get() == "LOTSIZE":
            symbol, lotsize = item["values"]
            self.show_edit_dialog(symbol, lotsize)
        elif self.mode_var.get() == "UNDERLYINGCODE":
            symbol, code = item["values"]
            self.show_edit_dialog(symbol, code, is_underlying_code=True)
        elif self.current_dataset_name == "fund_filename_map":
            values = item["values"]
            fund_code = values[0]
            fund_name = values[1] if len(values) > 1 else ""
            cds_name = values[2] if len(values) > 2 else ""
            password = values[3] if len(values) > 3 else ""
            self.show_fund_edit_dialog(fund_code, fund_name, cds_name, password)
        elif self.current_dataset_name in ["mcx_group2_filters", "fno_group2_filters"]:
            # Handle list-type datasets
            values = item["values"]
            index = int(values[0]) - 1  # Convert back to 0-based index
            filter_value = values[1] if len(values) > 1 else ""
            self.show_filter_edit_dialog(index, filter_value)
        else:
            header, value = item["values"]
            self.show_header_edit_dialog(header, value)

    def show_edit_dialog(self, symbol="", lotsize="", is_underlying_code=False):
        dialog = tk.Toplevel(self)
        dialog.title("Edit Item" if symbol else "Add New Item")
        dialog.geometry("400x200")
        dialog.configure(bg="#ecf0f1")
        dialog.transient(self)
        dialog.grab_set()
        dialog.geometry("+%d+%d" % (self.winfo_rootx() + 50, self.winfo_rooty() + 50))
        symbol_frame = tk.Frame(dialog, bg="#ecf0f1")
        symbol_frame.pack(fill="x", padx=20, pady=10)
        tk.Label(symbol_frame, text="Symbol:", font=("Arial", 12), bg="#ecf0f1").pack(anchor="w")
        symbol_entry = tk.Entry(symbol_frame, font=("Arial", 12), width=30)
        symbol_entry.pack(fill="x", pady=5)
        symbol_entry.insert(0, symbol)
        value_frame = tk.Frame(dialog, bg="#ecf0f1")
        value_frame.pack(fill="x", padx=20, pady=10)
        value_label_text = "Code:" if is_underlying_code else "Lotsize:"
        tk.Label(value_frame, text=value_label_text, font=("Arial", 12), bg="#ecf0f1").pack(anchor="w")
        value_entry = tk.Entry(value_frame, font=("Arial", 12), width=30)
        value_entry.pack(fill="x", pady=5)
        value_entry.insert(0, str(lotsize))
        button_frame = tk.Frame(dialog, bg="#ecf0f1")
        button_frame.pack(fill="x", padx=20, pady=20)

        def save_changes():
            new_symbol = symbol_entry.get().strip().upper()
            new_value = value_entry.get().strip()
            if not new_symbol or not new_value:
                error_msg = "Both Symbol and Code are required" if is_underlying_code else "Both Symbol and Lotsize are required"
                messagebox.showerror("Error", error_msg)
                return
            try:
                value = int(new_value)
            except ValueError:
                error_msg = "Code must be a valid number" if is_underlying_code else "Lotsize must be a valid number"
                messagebox.showerror("Error", error_msg)
                return
            if is_underlying_code:
                if symbol and symbol != new_symbol:
                    del self.underlying_code_data[symbol]
                self.underlying_code_data[new_symbol] = value
            else:
                if symbol and symbol != new_symbol:
                    del self.lotsize_data[symbol]
                self.lotsize_data[new_symbol] = value
            self.load_data()
            self.auto_save()
            dialog.destroy()

        def cancel_changes():
            dialog.destroy()

        tk.Button(button_frame, text="Save", bg="#27ae60", fg="white", font=("Arial", 10, "bold"), relief="flat", padx=20, pady=5, command=save_changes).pack(side="left", padx=5)
        tk.Button(button_frame, text="Cancel", bg="#e74c3c", fg="white", font=("Arial", 10, "bold"), relief="flat", padx=20, pady=5, command=cancel_changes).pack(side="left", padx=5)

    def show_fund_edit_dialog(self, fund_code="", fund_name="", cds_name="", password=""):
        dialog = tk.Toplevel(self)
        dialog.title("Edit Fund" if fund_code else "Add New Fund")
        dialog.geometry("500x350")
        dialog.configure(bg="#ecf0f1")
        dialog.transient(self)
        dialog.grab_set()
        dialog.geometry("+%d+%d" % (self.winfo_rootx() + 50, self.winfo_rooty() + 50))
        
        # Custodian Account
        code_frame = tk.Frame(dialog, bg="#ecf0f1")
        code_frame.pack(fill="x", padx=20, pady=10)
        tk.Label(code_frame, text="Custodian Account:", font=("Arial", 12), bg="#ecf0f1").pack(anchor="w")
        code_entry = tk.Entry(code_frame, font=("Arial", 12), width=40)
        code_entry.pack(fill="x", pady=5)
        code_entry.insert(0, fund_code)
        
        # Fund Name (Default)
        name_frame = tk.Frame(dialog, bg="#ecf0f1")
        name_frame.pack(fill="x", padx=20, pady=10)
        tk.Label(name_frame, text="Fund Name (Default):", font=("Arial", 12), bg="#ecf0f1").pack(anchor="w")
        name_entry = tk.Entry(name_frame, font=("Arial", 12), width=40)
        name_entry.pack(fill="x", pady=5)
        name_entry.insert(0, fund_name)
        
        # Fund CDS Name (Optional)
        cds_frame = tk.Frame(dialog, bg="#ecf0f1")
        cds_frame.pack(fill="x", padx=20, pady=10)
        tk.Label(cds_frame, text="Fund CDS Name (Optional):", font=("Arial", 12), bg="#ecf0f1").pack(anchor="w")
        cds_entry = tk.Entry(cds_frame, font=("Arial", 12), width=40)
        cds_entry.pack(fill="x", pady=5)
        cds_entry.insert(0, cds_name)
        
        # Password
        password_frame = tk.Frame(dialog, bg="#ecf0f1")
        password_frame.pack(fill="x", padx=20, pady=10)
        tk.Label(password_frame, text="Password:", font=("Arial", 12), bg="#ecf0f1").pack(anchor="w")
        password_entry = tk.Entry(password_frame, font=("Arial", 12), width=40)
        password_entry.pack(fill="x", pady=5)
        password_entry.insert(0, password)
        
        button_frame = tk.Frame(dialog, bg="#ecf0f1")
        button_frame.pack(fill="x", padx=20, pady=20)

        def save_changes():
            new_code = code_entry.get().strip()
            new_name = name_entry.get().strip()
            new_cds = cds_entry.get().strip()
            new_password = password_entry.get().strip()
            
            if not new_code:
                messagebox.showerror("Error", "Custodian Account is required")
                return
            if not new_name:
                messagebox.showerror("Error", "Fund Name is required")
                return
            
            # Build fund data structure
            fund_data = {
                "Fund Names": {
                    "Default": new_name
                },
                "Password": new_password
            }
            
            # Add CDS name if provided
            if new_cds:
                fund_data["Fund Names"]["CDS"] = new_cds
            
            # If editing and code changed, delete old entry
            if fund_code and fund_code != new_code:
                if fund_code in self.header_data:
                    del self.header_data[fund_code]
            
            # Save new/updated data
            self.header_data[new_code] = fund_data
            self.load_data()
            self.auto_save()
            dialog.destroy()

        def cancel_changes():
            dialog.destroy()

        tk.Button(button_frame, text="Save", bg="#27ae60", fg="white", font=("Arial", 10, "bold"), relief="flat", padx=20, pady=5, command=save_changes).pack(side="left", padx=5)
        tk.Button(button_frame, text="Cancel", bg="#e74c3c", fg="white", font=("Arial", 10, "bold"), relief="flat", padx=20, pady=5, command=cancel_changes).pack(side="left", padx=5)

    def delete_item(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select an item to delete")
            return
        item = self.tree.item(selection[0])
        if self.mode_var.get() == "LOTSIZE":
            symbol, _ = item["values"]
            if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete '{symbol}'?"):
                del self.lotsize_data[symbol]
                self.load_data()
                self.auto_save()
        else:
            values = item["values"]
            if not values or len(values) == 0:
                messagebox.showerror("Error", "Cannot delete: No data found in selected row")
                return
            # Ensure we're working with the dataset dictionary directly first
            if self.current_dataset_name not in self.datasets:
                self.datasets[self.current_dataset_name] = {}
            dataset_dict = self.datasets[self.current_dataset_name]
            self.header_data = dataset_dict  # Point to the same dictionary
            
            # Get identifier from tree - but tree might convert "07536" to 7536
            # So we'll match by comparing with all dictionary keys
            identifier = values[0]
            if identifier is None:
                messagebox.showerror("Error", "Cannot delete: Invalid identifier")
                return
            
            # Convert identifier to string (but this might lose leading zeros if tree converted it)
            identifier_from_tree = str(identifier).strip()
            
            if not identifier_from_tree:
                messagebox.showerror("Error", "Cannot delete: Invalid identifier")
                return
            
            # Find the matching key in dictionary
            # Tree might convert "07536" to 7536, so we match by row position
            # since tree displays items in the same order as dictionary items
            found_key = None
            identifier_str_for_display = identifier_from_tree
            
            # Get the selected row index
            selected_item = selection[0]
            all_items = self.tree.get_children()
            try:
                row_index = all_items.index(selected_item)
            except ValueError:
                row_index = -1
            
            # Try exact match first
            for key in dataset_dict.keys():
                key_str = str(key).strip()
                if key_str == identifier_from_tree:
                    found_key = key
                    identifier_str_for_display = key_str
                    break
            
            # If not found by exact match and we have row index, match by position
            # This handles cases where tree converted "07536" to 7536
            if found_key is None and row_index >= 0:
                keys_list = list(dataset_dict.keys())
                if row_index < len(keys_list):
                    found_key = keys_list[row_index]
                    identifier_str_for_display = str(found_key).strip()
            
            if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete '{identifier_str_for_display}'?"):
                if found_key is not None:
                    del dataset_dict[found_key]
                    self.load_data()
                    self.auto_save()
                else:
                    # Debug: show what's actually in the data
                    available_keys = list(dataset_dict.keys())[:10]  # First 10 keys for debugging
                    messagebox.showerror("Error", f"Item '{identifier_from_tree}' not found in data.\n\nLooking for: '{identifier_from_tree}'\nAvailable keys: {', '.join(str(k) for k in available_keys) if available_keys else 'None'}")

    def show_context_menu(self, event):
        selection = self.tree.selection()
        if selection:
            self.context_menu.post(event.x_root, event.y_root)

    def auto_save(self):
        try:
            from my_app.file_utils import get_app_directory
            app_dir = get_app_directory()
            consolidated_path = os.path.join(app_dir, "consolidated_data.json")
            
            # Load existing consolidated data or create new
            consolidated_data = {}
            if os.path.exists(consolidated_path):
                try:
                    with open(consolidated_path, "r") as f:
                        consolidated_data = json.load(f)
                except Exception:
                    consolidated_data = {}
            
            # Update the appropriate section
            if self.mode_var.get() == "LOTSIZE":
                consolidated_data["lotsize_data"] = self.lotsize_data
            elif self.mode_var.get() == "UNDERLYINGCODE":
                consolidated_data["underlying_code_data"] = self.underlying_code_data
            else:
                consolidated_data[self.current_dataset_name] = self.header_data
            
            # Always save underlying_code_data and lotsize_data to ensure they're persisted
            consolidated_data["underlying_code_data"] = self.underlying_code_data
            consolidated_data["lotsize_data"] = self.lotsize_data
            
            # Save consolidated data
            with open(consolidated_path, "w") as f:
                json.dump(consolidated_data, f, indent=4)
        except Exception as e:
            print(f"Auto-save failed: {e}")

    def save_data(self):
        try:
            from my_app.file_utils import get_app_directory
            app_dir = get_app_directory()
            consolidated_path = os.path.join(app_dir, "consolidated_data.json")
            
            # Load existing consolidated data or create new
            consolidated_data = {}
            if os.path.exists(consolidated_path):
                try:
                    with open(consolidated_path, "r") as f:
                        consolidated_data = json.load(f)
                except Exception:
                    consolidated_data = {}
            
            # Update the appropriate section
            if self.mode_var.get() == "LOTSIZE":
                consolidated_data["lotsize_data"] = self.lotsize_data
            elif self.mode_var.get() == "UNDERLYINGCODE":
                consolidated_data["underlying_code_data"] = self.underlying_code_data
            else:
                consolidated_data[self.current_dataset_name] = self.header_data
            
            # Always save underlying_code_data and lotsize_data to ensure they're persisted
            consolidated_data["underlying_code_data"] = self.underlying_code_data
            consolidated_data["lotsize_data"] = self.lotsize_data
            
            # Save consolidated data
            with open(consolidated_path, "w") as f:
                json.dump(consolidated_data, f, indent=4)
            messagebox.showinfo("Success", "Data saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save data: {str(e)}")

    def load_from_file(self):
        try:
            from my_app.file_utils import get_app_directory
            app_dir = get_app_directory()
            consolidated_path = os.path.join(app_dir, "consolidated_data.json")
            
            if os.path.exists(consolidated_path):
                with open(consolidated_path, "r") as f:
                    consolidated_data = json.load(f)
                
                if self.mode_var.get() == "LOTSIZE":
                    lotsize_data = consolidated_data.get("lotsize_data", {})
                    if lotsize_data:
                        self.lotsize_data = lotsize_data
                        self.load_data()
                        messagebox.showinfo("Success", "Data loaded successfully!")
                    else:
                        messagebox.showwarning("Warning", "No lotsize data found in consolidated file")
                elif self.mode_var.get() == "UNDERLYINGCODE":
                    underlying_code_data = consolidated_data.get("underlying_code_data", {})
                    if underlying_code_data:
                        self.underlying_code_data = underlying_code_data
                        self.load_data()
                        messagebox.showinfo("Success", "Data loaded successfully!")
                    else:
                        messagebox.showwarning("Warning", "No underlying code data found in consolidated file")
                else:
                    dataset_data = consolidated_data.get(self.current_dataset_name, {})
                    if dataset_data:
                        self.header_data = dataset_data
                        self.datasets[self.current_dataset_name] = dataset_data
                        self.load_data()
                        messagebox.showinfo("Success", "Data loaded successfully!")
                    else:
                        messagebox.showwarning("Warning", f"No {self.current_dataset_name} data found in consolidated file")
            else:
                messagebox.showwarning("Warning", "No consolidated_data.json found")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load data: {str(e)}")

    def reset_to_default(self):
        if messagebox.askyesno("Confirm Reset", "Are you sure you want to reset to default values? This will overwrite current data."):
            if self.mode_var.get() == "LOTSIZE":
                self.lotsize_data = self.default_lotsize_data.copy()
            elif self.mode_var.get() == "UNDERLYINGCODE":
                self.underlying_code_data = self.default_underlying_code_data.copy()
            else:
                # Reset selected header dataset to its defaults
                self.header_data = self._load_default_dataset_data(self.current_dataset_name)
                self.datasets[self.current_dataset_name] = self.header_data
            self.load_data()
            self.auto_save()
            messagebox.showinfo("Success", "Data reset to default values!")

    def _load_default_header_data(self):
        # Backwards compatibility; delegate to dataset-specific loader for first discovered dataset
        default_name = next(iter(self.dataset_files.keys()), "")
        return self._load_default_dataset_data(default_name)

    def _load_default_dataset_data(self, dataset_name):
        # IMPORTANT: This function should ONLY return defaults, never read from consolidated_data.json.
        # Reading the whole file here causes wrong data structures to be injected when a dataset is missing.

        # Default fund filename map (pre-populated mapping)
        if dataset_name == "fund_filename_map":
            return {"DBSBK0000033": {"Fund Names": {"Default": "DIF-Class 1 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000036": {"Fund Names": {"Default": "DIF-Class 2 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000038": {"Fund Names": {"Default": "DIF-Class 3 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000042": {"Fund Names": {"Default": "DIF-Class 5 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000044": {"Fund Names": {"Default": "DIF-Class 6 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000043": {"Fund Names": {"Default": "DIF-Class 7 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000049": {"Fund Names": {"Default": "DIF-Class 8 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000050": {"Fund Names": {"Default": "DIF-Class 9 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000051": {"Fund Names": {"Default": "DIF-Class 10 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000052": {"Fund Names": {"Default": "DIF-Class 11 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000074": {"Fund Names": {"Default": "DIF-Class 12 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000179": {"Fund Names": {"Default": "DIF-Class 13 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000189": {"Fund Names": {"Default": "DIF-Class 14 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000192": {"Fund Names": {"Default": "DIF-Class 15 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000214": {"Fund Names": {"Default": "DIF-Class 16 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000216": {"Fund Names": {"Default": "DIF-Class 17 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000217": {"Fund Names": {"Default": "DIF-Class 18_Moon"}, "Password": "AAGCD0792B"}, "DBSBK0000232": {"Fund Names": {"Default": "DIF-Class 19 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000247": {"Fund Names": {"CDS": "DIF-Class 21 CDS Holding", "Default": "DIF-Class 21 Holding"}, "Password": "AAGCD0792B"}, "DBSBK0000178": {"Fund Names": {"Default": "DGF-Cell 8"}, "Password": "AAICD1968M"}, "BNPP00000458": {"Fund Names": {"Default": "DGF-Cell 9"}, "Password": "AAICD2891H"}, "DGF-Cell 10": {"Fund Names": {"Default": "DGF-Cell 10"}, "Password": "AAICD3412C"}, "BNPP00000480": {"Fund Names": {"Default": "DGF-Cell 11"}, "Password": "AAICD6359G"}, "BNPP00000488": {"Fund Names": {"Default": "DGF-Cell 13"}, "Password": "AAICD7821M"}, "BNPP00000540": {"Fund Names": {"Default": "DGF-Cell 16"}, "Password": "AAJCD5624K"}, "BNPP00000535": {"Fund Names": {"Default": "DGF-Cell 17"}, "Password": "AAJCD4991K"}, "DBSBK0000229": {"Fund Names": {"Default": "DGF-Cell 18"}, "Password": "AAJCD6205G"}, "DBSBK0000228": {"Fund Names": {"Default": "DGF-Cell 19"}, "Password": "AAJCD6049E"}, "DBSBK0000285": {"Fund Names": {"CDS": "DGF-Cell 23 CDS Holding", "Default": "DGF-Cell 23 Holding"}, "Password": "AAKCD6244Q"}, "DBSBK0000299": {"Fund Names": {"Default": "DGF-Cell 24 Holding"}, "Password": "AAKCD7324B"}, "DBSBK0000353": {"Fund Names": {"Default": "DGF-Cell 28 Holding"}, "Password": "AALCD1140J"}, "DBSBK0000354": {"Fund Names": {"Default": "DGF-Cell 29 Holding"}, "Password": "AALCD1141K"}, "DBSBK0000380": {"Fund Names": {"Default": "DGF-Cell 32 Holding"}, "Password": ""}, "DBSBK0000356": {"Fund Names": {"Default": "GlobalQ_AIF-III"}, "Password": ""}, "DGF-Cell 36": {"Fund Names": {"Default": "DGF-Cell 36"}, "Password": ""}, "DGF-Cell 38": {"Fund Names": {"Default": "DGF-Cell 38"}, "Password": ""}}
        
        # Default ASIO portfolio mapping
        if dataset_name == "asio_recon_portfolio_mapping":
            return {
                "ASIO - SF 3": "THE ASIO FUND VCC - SUB-FUND 3",
                "ASIO - SF 8_Golden": "THE ASIO FUND VCC - EMKAY BHARAT FUND - BHARATS GOLDEN DECADE",
                "ASIO-SF1_": "THE ASIO FUND VCC - FORT PANGEA",
                "ASIO - SF 8_Capital Builder": "THE ASIO FUND VCC - EMKAY BHARAT FUND",
                "ASIO - SF 8_Envi": "THE ASIO FUND VCC - EMKAY BHARAT FUND - ENVI",
                "ASIO - SF 4": "THE ASIO FUND VCC - SUB FUND 4",
                "ASIO - SF 10": "THE ASIO FUND VCC-HATHI CAPITAL INDIA GROWTH FUND I",
                "ASIO - SF 6": "THE ASIO FUND VCC - FRAGRANT HARBOUR INDIA OPPORTUNITIES FUND",
                "ASIO - SF 2": "",
                "ASIO - SF 7": ""
            }
        
        # Default Format 1 headers
        if dataset_name == "asio_recon_format_1_headers":
            return {
                "CLN_CODE": "Cln Code",
                "CLN_NAME": "Cln Name",
                "INSTR_CODE": "Instr Code",
                "INSTR_ISIN": "Instr ISIN",
                "INSTR_NAME": "Instr Name",
                "DEMAT_PHYSICAL": "Demat / Physical",
                "SETTLED_POSITION": "Settled Position",
                "PENDING_PURCHASE": "Pending Purchase",
                "PENDING_SALE": "Pending Sale",
                "BLOCKED": "Blocked",
                "BLOCKED_RECEIVABLES": "Blocked Receivables",
                "BLOCKED_PAYABLES": "Blocked Payables",
                "PENDING_CA_ENTITLEMENTS": "Pending CA Entitlements",
                "MARKET_PRICE_DATE": "Market Price Date",
                "MARKET_PRICE": "Market Price",
                "SETTLED_BLOCKED_MARKET_PRICE": "(Settled + Blocked * Market Price)",
                "SALEABLE": "Saleable",
                "CONTRACTUAL": "Contractual"
            }
        
        # Default Format 2 headers
        if dataset_name == "asio_recon_format_2_headers":
            return {
                "VALUE_DATE_AS_AT": "Value date as at",
                "SERVICE_LOCATION": "Service location",
                "SECURITIES_ACCOUNT_NAME": "Securities account name",
                "SECURITIES_ACCOUNT_NUMBER": "Securities account number",
                "SECURITY_TYPE": "Security type",
                "SECURITY_NAME": "Security name",
                "ISIN": "ISIN",
                "PLACE_OF_SETTLEMENT": "Place of settlement",
                "SETTLED_BALANCE": "Settled balance",
                "ANTICIPATED_DEBITS": "Anticipated debits",
                "AVAILABLE_BALANCE": "Available balance",
                "ANTICIPATED_CREDITS": "Anticipated credits",
                "TRADED_BALANCE": "Traded balance",
                "SECURITY_CURRENCY": "Security currency",
                "SECURITY_PRICE": "Security price",
                "INDICATIVE_SETTLED_VALUE": "Indicative settled value",
                "INDICATIVE_TRADED_VALUE": "Indicative traded value",
                "INDICATIVE_SETTLED_VALUE_INR": "Indicative settled value in INR",
                "INDICATIVE_TRADED_VALUE_INR": "Indicative traded value in INR"
            }
        
        # Default BhavCopy headers
        if dataset_name == "asio_recon_bhavcopy_headers":
            return {
                "TRADDT": "TradDt",
                "BIZDT": "BizDt",
                "SGMT": "Sgmt",
                "SRC": "Src",
                "FININSTRMTP": "FinInstrmTp",
                "FININSTRMID": "FinInstrmId",
                "ISIN": "ISIN",
                "TCKRSYMB": "TckrSymb",
                "SCTYSRS": "SctySrs",
                "XPRYDT": "XpryDt",
                "FININSTRMACTLXPRYDT": "FininstrmActlXpryDt",
                "STRKPRIC": "StrkPric",
                "OPTNTP": "OptnTp",
                "FININSTRMnm": "FinInstrmNm",
                "OPNPRIC": "OpnPric",
                "HGHPRIC": "HghPric",
                "LWPRIC": "LwPric",
                "CLSPRIC": "ClsPric",
                "LASTPRIC": "LastPric"
            }
        
        # Default Geneva headers
        if dataset_name == "asio_geneva_headers":
            return {
                "PORTFOLIO": "Portfolio",
                "INVESTMENT": "Investment",
                "INVESTMENT_DESCRIPTION": "Investment Description",
                "TRADED_QUANTITY": "Traded Quantity",
                "SETTLED_QUANTITY": "Settled Quantity",
                "CURRENCY": "Currency",
                "UNIT_COST": "Unit Cost",
                "COST_LOCAL": "Cost Local",
                "COST_BOOK": "Cost Book",
                "MARKET_PRICE_LOCAL": "Market Price Local",
                "MARKET_VALUE_LOCAL": "Market Value Local"
            }
        
        # Default trade headers data
        if dataset_name == "trade_headers":
            return {
                "CLIENT_CODE": "Client\nCode",
                "EXCHANGE": "Exchange",
                "MKT_TYPE": "Mkt\nType",
                "DATE": "Date",
                "TRADE_TIME": "TradeTime",
                "SETL_NO": "SetlNo",
                "SCRIP_CODE": "Scrip Code",
                "SCRIP_NAME": "Scrip Name",
                "ISIN": "ISIN",
                "BUY_QTY": "Buy Qty",
                "BUY_RATE": "Buy Rate",
                "BUY_NET_RATE": "Buy\nNet Rate",
                "BUY_VALUE": "Buy Value",
                "SELL_QTY": "Sell Qty",
                "BALANCE": "Balance",
                "SELL_RATE": "Sell Rate",
                "SELL_NET_RATE": "Sell\nNet Rate",
                "SELL_VALUE": "Sell Value",
                "BREAK_EVEN": "BreakEven",
                "MARKET_RATE": "Market Rate",
                "TERMINAL": "Terminal",
                "DELY": "Dely"
            }
        
        # Default Geneva-Custodian mapping
        if dataset_name == "geneva_custodian_mapping":
            return {
                "DIF-Class 5_Heer": "DOVETAIL INDIA FUND CLASS 5 SHARES",
                "DIF-Class 5_Moonrock_F&O": "DOVETAIL INDIA FUND CLASS 5 SHARES",
                "DIF-Class 3_Saankhya": "DOVETAIL INDIA FUND CLASS 3 Shares",
                "DIF-Class 11_Eagle": "DOVETAIL INDIA FUND CLASS 11 SHARES",
                "DGF-Cell 2": "DOVETAIL GLOBAL FUND PCC CELL 2",
                "DIF-Class 1": "DOVETAIL INDIA FUND",
                "DGF-Cell 23": "DOVETAIL GLOBAL FUND PCC - CELL 23",
                "DIF-Class 6_Amal Parikh": "DOVETAIL INDIA FUND CLASS 6 SHARES",
                "DIF-Class 21": "DOVETAIL INDIA FUND CLASS 21",
                "DGF-Cell 28": "DOVETAIL GLOBAL FUND PCC - CELL 28",
                "DGF-Cell 29": "DOVETAIL GLOBAL FUND PCC CELL 29",
                "DIF-Class 2_Agile": "DOVETAIL INDIA FUND CLASS 2 SHARES",
                "DIF-Class 19": "DOVETAIL INDIA FUND CLASS 19",
                "ASIO - SF 2": "THE ASIO FUND VCC - SUB-FUND 2",
                "ASIO - SF 3": "THE ASIO FUND VCC - SUB-FUND 3",
                "ASIO - SF 4": "THE ASIO FUND VCC - SUB-FUND 4",
                "ASIO - SF 7": "THE ASIO FUND VCC FORT CAPITAL ABSOLUTE ALTERNATES",
                "DIF-Class 18_Moon": "DOVETAIL INDIA FUND CLASS 18",
                "DIF-Class 12_F&O": "DOVETAIL INDIA FUND CLASS 12",
                "DIF-Class 12_Moon New F&O": "DOVETAIL INDIA FUND CLASS 12",
                "DIF-Class 6": "DOVETAIL INDIA FUND CLASS 6 SHARES",
                "DIF-Class 8_Discipline": "DOVETAIL INDIA FUND CLASS 8 SHARES",
                "GlobalQ_AIF-III": "GLOBAL Q FUND",
                "DGF-Cell 32": "DOVETAIL GLOBAL FUND PCC - CELL 32",
                "DIF-Class 8_Panda": "DOVETAIL INDIA FUND CLASS 8 SHARES",
                "DGF- Cell 36": "",
                "DGF- Cell 38": "",
                "DIF-Class 8_HBE Capital": "DOVETAIL INDIA FUND CLASS 8 SHARES"
            }
        
        # Default FNO TM Code to TM Name mapping
        if dataset_name == "fno_tm_code_with_tm_name":
            return {
                "13302": "Achintya Securities Pvt. Ltd.",
                "07536": "Trustline Securities Limited",
                "10412": "Motilal Oswal Financial Services Limited",
                "07714": "SMC Global Securities Limited"
            }
        
        # Default MCX TM Code to TM Name mapping
        if dataset_name == "mcx_tm_code_with_tm_name":
            return {
                    "31640": "Achintya Securities Pvt. Ltd.",
                    "10515": "SMC Global Securities Limited"
            }
        
        # Default ASIO SF2 trade loader
        if dataset_name == "asio_sf_2_trade_loader":
            return {
                "RecordAction": "InsertUpdate",
                "KeyValue.KeyName": "",
                "UserTranId1": "",
                "Portfolio": "ASIO - SF 2_INR",
                "LocationAccount": "Asio_Sub Fund_2_OHM_FO_KOTBK0001479",
                "Broker": "",
                "PriceDenomination": "CALC",
                "CounterInvestment": "INR",
                "NetInvestmentAmount": "CALC",
                "NetCounterAmount": "CALC",
                "tradeFX": "",
                "ContractFxRateNumerator": "",
                "ContractFxRateDenominator": "",
                "ContractFxRate": "",
                "NotionalAmount": "",
                "FundStructure2": "CALC",
                "SpotDate": "",
                "PriceDirectly": "",
                "CounterFXDenomination": "CALC",
                "CounterTDateFx": "",
                "AccruedInterest": "",
                "InvestmentAccruedInterest": "",
                "TradeExpenses.ExpenseNumber": "",
                "TradeExpenses.ExpenseCode": "",
                "TradeExpenses.ExpenseAmt": "",
                "NonCapExpenses.NonCapNumber": "",
                "NonCapExpenses.NonCapExpenseCode": "",
                "NonCapExpenses.NonCapAmount": "",
                "NonCapExpenses.NonCapCurrency": "",
                "NonCapExpenses.LocationAccount": "",
                "NonCapExpenses.NonCapLiabilityCode": "",
                "NonCapExpenses.NonCapPaymentType": ""
            }

        # Default ASIO SF2 option security (FT/EQFT style)
        if dataset_name == "asio_sf_2_option_security":
            return {
                "Exchange": "NSE",
                "Issuer": "",
                "Ticker": "",
                "Cusip": "",
                "Sedol": "",
                "Isin": "",
                "AltKey1": "",
                "AltKey2": "",
                "BloombergID": "",
                "BloombergTicker": "",
                "BloombergUniqueID": "",
                "BloombergMarketSector": "",
                "SettleDays": "",
                "PricingFactor": "",
                "TradingFactor": 1,
                "CurrentPriceDayRange": 1,
                "AssetType": "FT",
                "InvestmentType": "EQFT",
                "PriceDenomination": "INR",
                "BifurcationCurrency": "INR",
                "PrincipalCurrency": "INR",
                "IncomeCurrency": "INR",
                "RiskCurrency": "INR",
                "IssueCountry": "IN",
                "WithholdingTaxType": "Standard",
                "QDIEligibilityFlag": "Not Eligible",
                "SharesOutstanding": "",
                "SubIndustry": "",
                "SubIndustry2": "",
                "QuantityPrecision": 3,
                "InvestmentCrossZero": "Use Accounting Parameters",
                "PriceByPreference": "Currency",
                "ForwardPriceInterpolateFlag": 0,
                "PricingPrecision": 3,
                "FirstMarkDate": "01-01-2022",
                "LastMarkDate": "",
                "PriceList": "",
                "AutoGenerateMarks": 1,
                "CashSettlement": 1
            }

        if dataset_name == "asio_sf_2_mcx_trade_loader":
            return {
                "RecordAction": "InsertUpdate",
                "KeyValue.KeyName": "",
                "UserTranId1": "",
                "Portfolio": "ASIO - SF 2_INR",
                "LocationAccount": "Asio_Sub Fund_2_OHM_MCX_KOTBK0001479",
                "Broker": "",
                "PriceDenomination": "CALC",
                "CounterInvestment": "INR",
                "NetInvestmentAmount": "CALC",
                "NetCounterAmount": "CALC",
                "tradeFX": "",
                "ContractFxRateNumerator": "",
                "ContractFxRateDenominator": "",
                "ContractFxRate": "",
                "NotionalAmount": "",
                "FundStructure2": "CALC",
                "SpotDate": "",
                "PriceDirectly": "",
                "CounterFXDenomination": "CALC",
                "CounterTDateFx": "",
                "AccruedInterest": "",
                "InvestmentAccruedInterest": "",
                "TradeExpenses.ExpenseNumber": "",
                "TradeExpenses.ExpenseCode": "",
                "TradeExpenses.ExpenseAmt": "",
                "NonCapExpenses.NonCapNumber": "",
                "NonCapExpenses.NonCapExpenseCode": "",
                "NonCapExpenses.NonCapAmount": "",
                "NonCapExpenses.NonCapCurrency": "",
                "NonCapExpenses.LocationAccount": "",
                "NonCapExpenses.NonCapLiabilityCode": "",
                "NonCapExpenses.NonCapPaymentType": ""
            }

        if dataset_name == "asio_sf_2_mcx_option_security":
            return {
                "Exchange": "MCX",
                "Issuer": "",
                "Ticker": "",
                "Cusip": "",
                "Sedol": "",
                "Isin": "",
                "Riccode": "",
                "AltKey1": "",
                "AltKey2": "",
                "BloombergID": "",
                "BloombergTicker": "",
                "BloombergUniqueID": "",
                "BloombergMarketSector": "",
                "SettleDays": 0,
                "PricingFactor": 1,
                "CurrentPriceDayRange": 1,
                "AssetType": "OP",
                "InvestmentType": "CMOP",
                "PriceDenomination": "",
                "BifurcationCurrency": "INR",
                "PrincipalCurrency": "INR",
                "IncomeCurrency": "INR",
                "RiskCurrency": "INR",
                "IssueCountry": "IN",
                "WithholdingTaxType": "Standard",
                "QDIEligibilityFlag": "Not Eligible",
                "SharesOutstanding": "",
                "Beta": "",
                "SubIndustry": "",
                "SubIndustry2": "",
                "NonDeliverableCurrencyFlag": "",
                "QuantityPrecision": "",
                "InvestmentCrossZero": "",
                "PriceByPreference": "",
                "ForwardPriceInterpolateFlag": "",
                "SecFeeSchedule": "",
                "SecEligibleFlag": "",
                "WashSalesEligible": "",
                "ExerciseStyle": "European",
                "PricingPrecision": 15,
                "CashSettledFlag": 1
            }

        # Default ASIO SF2 future security (FT/EQFT style)
        if dataset_name == "asio_sf_2_future_security":
            return {
                "Exchange": "NSE",
                "Issuer": "",
                "Ticker": "",
                "Cusip": "",
                "Sedol": "",
                "Isin": "",
                "AltKey1": "",
                "AltKey2": "",
                "BloombergID": "",
                "BloombergTicker": "",
                "BloombergUniqueID": "",
                "BloombergMarketSector": "",
                "SettleDays": "",
                "PricingFactor": "",
                "TradingFactor": 1,
                "CurrentPriceDayRange": 1,
                "AssetType": "FT",
                "InvestmentType": "EQFT",
                "PriceDenomination": "INR",
                "BifurcationCurrency": "INR",
                "PrincipalCurrency": "INR",
                "IncomeCurrency": "INR",
                "RiskCurrency": "INR",
                "IssueCountry": "IN",
                "WithholdingTaxType": "Standard",
                "QDIEligibilityFlag": "Not Eligible",
                "SharesOutstanding": "",
                "SubIndustry": "",
                "SubIndustry2": "",
                "QuantityPrecision": 3,
                "InvestmentCrossZero": "Use Accounting Parameters",
                "StrikePrice":0,
                "PriceByPreference": "Currency",
                "ForwardPriceInterpolateFlag": 0,
                "PricingPrecision": 3,
                "FirstMarkDate": "01-01-2022",
                "LastMarkDate": "",
                "PriceList": "",
                "AutoGenerateMarks": 1,
                "CashSettlement": 1
            }
        
        # Default ASIO SF2 MCX future security (FT/CMFT style)
        if dataset_name == "asio_sf_2_mcx_future_security":
            return {
                "Exchange": "MCX",
                "Issuer": "",
                "Ticker": "",
                "Cusip": "",
                "Sedol": "",
                "Isin": "",
                "AltKey1": "",
                "AltKey2": "",
                "BloombergID": "",
                "BloombergTicker": "",
                "BloombergUniqueID": "",
                "BloombergMarketSector": "",
                "SettleDays": "",
                "PricingFactor": "",
                "CurrentPriceDayRange": 1,
                "AssetType": "FT",
                "InvestmentType": "CMFT",
                "PriceDenomination": "INR",
                "BifurcationCurrency": "INR",
                "PrincipalCurrency": "INR",
                "IncomeCurrency": "INR",
                "RiskCurrency": "INR",
                "IssueCountry": "IN",
                "WithholdingTaxType": "Standard",
                "QDIEligibilityFlag": "Not Eligible",
                "SharesOutstanding": "",
                "SubIndustry": "",
                "SubIndustry2": "",
                "QuantityPrecision": 3,
                "InvestmentCrossZero": "Use Accounting Parameters",
                "PriceByPreference": "Currency",
                "ForwardPriceInterpolateFlag": 0,
                "PricingPrecision": 3,
                "FirstMarkDate": "01-01-2022",
                "LastMarkDate": "",
                "PriceList": "",
                "AutoGenerateMarks": 1,
                "CashSettlement": 1
            }
        
        # If default header fields are provided and this is the first dataset, seed from it
        first_name = next(iter(self.dataset_files.keys()), None)
        if dataset_name == first_name and DEFAULT_HEADER_FIELDS:
            return {field: "" for field in DEFAULT_HEADER_FIELDS}
        # Otherwise default to empty mapping
        # Default ASIO Pricing FNO
        if dataset_name == "asio_pricing_fno":
            return {
                "RecordAction": "InsertUpdate",
                "PriceDenomination": "INR",
                "PriceList": "NSE_Equity",
                "TaxLotID": ""
            }
        
        # Default ASIO Pricing MCX
        if dataset_name == "asio_pricing_mcx":
            return {
                "RecordAction": "InsertUpdate",
                "PriceDenomination": "INR",
                "PriceList": "NSE_Equity",
                "TaxLotID": ""
            }
        
        # Default ASIO SF4 Trading Code to Location Account Mapping
        # Note: Use _get_location_account_from_trading_code() function in asio_sub_fund4.py to generate values dynamically
        if dataset_name == "asio_sf4_trading_code_mapping":
            return {
                "FT": "Asio_Sub Fund_4_OHM_FO_DBSBK0000289_FT",
                "FT1": "Asio_Sub Fund_4_OHM_FO_DBSBK0000289_FT1",
                "FT2": "Asio_Sub Fund_4_OHM_FO_DBSBK0000289_FT2",
                "FT3": "Asio_Sub Fund_4_OHM_FO_DBSBK0000289_FT3"
            }
        
        # Default ASIO Sub Fund 4 read configuration
        if dataset_name == "asio_sub_fund4_read_config":
            return {
                "read_from_row": 1,
                "read_from_column": "A"
            }
        
        # Default MCX Group2 Filters
        if dataset_name == "mcx_group2_filters":
            return ['Commodity Future Option', 'Commodity Option', 'Commodity Future']
        
        # Default FNO Group2 Filters
        if dataset_name == "fno_group2_filters":
            return ['Equity Option', 'Index Option', 'Index Future', 'Equity future']
        
        return {}

    def _discover_datasets(self):
        datasets = {}
        try:
            # First check for consolidated file
            from my_app.file_utils import get_app_directory
            app_dir = get_app_directory()
            consolidated_path = os.path.join(app_dir, "consolidated_data.json")
            if os.path.exists(consolidated_path):
                with open(consolidated_path, "r") as f:
                    consolidated_data = json.load(f)
                    # Add datasets from consolidated file
                for dataset_name in ["fund_filename_map", "asio_recon_portfolio_mapping", "asio_recon_format_1_headers", "asio_recon_format_2_headers", 
                                    "asio_recon_bhavcopy_headers", "asio_geneva_headers", "trade_headers", 
                                    "aafspl_car_future", "option_security", "car_trade_loader", "asio_sf_2_trade_loader", "asio_sf_2_option_security", "asio_sf_2_future_security", "asio_sf_2_mcx_trade_loader", "asio_sf_2_mcx_option_security", "geneva_custodian_mapping", "fno_tm_code_with_tm_name", "mcx_tm_code_with_tm_name", "asio_pricing_fno", "asio_pricing_mcx", "asio_sf4_ft", "asio_sf4_trading_code_mapping", "asio_sub_fund4_read_config", "mcx_group2_filters", "fno_group2_filters"]:
                        if dataset_name in consolidated_data:
                            datasets[dataset_name] = consolidated_path
            
            # Also check for individual files as fallback
            for fname in os.listdir(app_dir):
                if not fname.lower().endswith(".json"):
                    continue
                if fname == "lotsize_data.json" or fname == "consolidated_data.json":
                    continue
                # Use file stem as dataset name
                dataset_name = os.path.splitext(os.path.basename(fname))[0]
                if dataset_name not in datasets:  # Don't override consolidated data
                    datasets[dataset_name] = fname
        except Exception as e:
            print(f"Dataset discovery failed: {e}")
        
        # Always include these datasets (loaded from consolidated_data.json or defaults)
        for dataset_name in ["fund_filename_map", "asio_recon_portfolio_mapping", "asio_recon_format_1_headers", "asio_recon_format_2_headers", 
                            "asio_recon_bhavcopy_headers", "asio_geneva_headers", "trade_headers", 
                            "aafspl_car_future", "option_security", "car_trade_loader", "asio_sf_2_trade_loader", "asio_sf_2_option_security", "asio_sf_2_future_security", "asio_sf_2_mcx_trade_loader", "asio_sf_2_mcx_option_security", "asio_sf_2_mcx_future_security", "geneva_custodian_mapping", "asio_sf4_ft", "asio_sf4_trading_code_mapping", "asio_sub_fund4_read_config", "mcx_group2_filters", "fno_group2_filters"]:
            if dataset_name not in datasets:
                datasets[dataset_name] = consolidated_path  # All come from consolidated file
        
        return datasets

    def refresh_datasets(self):
        # Re-scan JSON files and update combobox
        self.dataset_files = self._discover_datasets()
        # Ensure current selection remains valid
        if self.current_dataset_name and self.current_dataset_name not in self.dataset_files:
            self.current_dataset_name = next(iter(self.dataset_files.keys()), "")
        # Load or initialize datasets map
        for dataset_name in self.dataset_files.keys():
            if dataset_name not in self.datasets:
                self.datasets[dataset_name] = self._load_default_dataset_data(dataset_name)
        self.header_data = self.datasets.get(self.current_dataset_name, {})
        # Update search values
        values = ["Lotsize", "UnderlyingCode"] + list(self.dataset_files.keys())
        self._all_dataset_values = values
        # Keep selection consistent in UI
        if self.mode_var.get() == "LOTSIZE":
            self.dataset_var.set("Lotsize")
        elif self.mode_var.get() == "UNDERLYINGCODE":
            self.dataset_var.set("UnderlyingCode")
        else:
            self.dataset_var.set(self.current_dataset_name)

    def _infer_type(self, value):
        if isinstance(value, bool):
            return "bool"
        if isinstance(value, int):
            return "int"
        if isinstance(value, float):
            return "float"
        if isinstance(value, str):
            for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y"):
                try:
                    datetime.strptime(value, fmt)
                    return "date"
                except Exception:
                    pass
            return "str"
        return type(value).__name__

    def _cast_from_string(self, text):
        s = str(text).strip()
        if s.lower() in ("true", "false"):
            return s.lower() == "true"
        try:
            if s.isdigit() or (s.startswith("-") and s[1:].isdigit()):
                return int(s)
        except Exception:
            pass
        try:
            return float(s)
        except Exception:
            pass
        return s

    def show_header_edit_dialog(self, field_name="", value=""):
        # Customize labels for different datasets
        if self.current_dataset_name == "asio_recon_portfolio_mapping":
            field_label = "Portfolio:"
            value_label = "Holding Name:"
            dialog_title = "Edit Portfolio Mapping" if field_name else "Add Portfolio Mapping"
        elif self.current_dataset_name == "geneva_custodian_mapping":
            field_label = "Fund name in Geneva:"
            value_label = "Fund name in Cust Holding:"
            dialog_title = "Edit Geneva-Custodian Mapping" if field_name else "Add Geneva-Custodian Mapping"
        elif self.current_dataset_name == "fno_tm_code_with_tm_name":
            field_label = "TM Code:"
            value_label = "TM Name:"
            dialog_title = "Edit FNO TM Code Mapping" if field_name else "Add FNO TM Code Mapping"
        elif self.current_dataset_name == "mcx_tm_code_with_tm_name":
            field_label = "TM Code:"
            value_label = "TM Name:"
            dialog_title = "Edit MCX TM Code Mapping" if field_name else "Add MCX TM Code Mapping"
        elif self.current_dataset_name in ["asio_recon_format_1_headers", "asio_recon_format_2_headers", "asio_recon_bhavcopy_headers", "asio_geneva_headers"]:
            field_label = "Header Key:"
            value_label = "Header Name:"
            dialog_title = "Edit Header" if field_name else "Add Header"
        else:
            field_label = "Header:"
            value_label = "Value:"
            dialog_title = "Edit Fixed Value" if field_name else "Add Fixed Value"
        
        dialog = tk.Toplevel(self)
        dialog.title(dialog_title)
        dialog.geometry("420x240")
        dialog.configure(bg="#ecf0f1")
        dialog.transient(self)
        dialog.grab_set()
        dialog.geometry("+%d+%d" % (self.winfo_rootx() + 50, self.winfo_rooty() + 50))
        
        # Show both field and value inputs
        field_frame = tk.Frame(dialog, bg="#ecf0f1")
        field_frame.pack(fill="x", padx=20, pady=10)
        tk.Label(field_frame, text=field_label, font=("Arial", 12), bg="#ecf0f1").pack(anchor="w")
        field_entry = tk.Entry(field_frame, font=("Arial", 12), width=30)
        field_entry.pack(fill="x", pady=5)
        field_entry.insert(0, field_name)
        
        value_frame = tk.Frame(dialog, bg="#ecf0f1")
        value_frame.pack(fill="x", padx=20, pady=10)
        tk.Label(value_frame, text=value_label, font=("Arial", 12), bg="#ecf0f1").pack(anchor="w")
        value_entry = tk.Entry(value_frame, font=("Arial", 12), width=30)
        value_entry.pack(fill="x", pady=5)
        try:
            if isinstance(value, (dict, list)):
                value_entry.insert(0, json.dumps(value))
            else:
                value_entry.insert(0, str(value))
        except Exception:
            value_entry.insert(0, str(value))
        
        button_frame = tk.Frame(dialog, bg="#ecf0f1")
        button_frame.pack(fill="x", padx=20, pady=20)

        def save_changes():
            new_field = field_entry.get().strip()
            raw = value_entry.get()
            if not new_field:
                if self.current_dataset_name == "asio_recon_portfolio_mapping":
                    error_msg = "Portfolio is required"
                elif self.current_dataset_name == "geneva_custodian_mapping":
                    error_msg = "Fund name in Geneva is required"
                elif self.current_dataset_name == "fno_tm_code_with_tm_name":
                    error_msg = "TM Code is required"
                elif self.current_dataset_name == "mcx_tm_code_with_tm_name":
                    error_msg = "TM Code is required"
                elif self.current_dataset_name in ["asio_recon_format_1_headers", "asio_recon_format_2_headers", "asio_recon_bhavcopy_headers", "asio_geneva_headers"]:
                    error_msg = "Header Key is required"
                else:
                    error_msg = "Field is required"
                messagebox.showerror("Error", error_msg)
                return
            try:
                new_value = json.loads(raw)
            except Exception:
                new_value = self._cast_from_string(raw)
            # Ensure we're working with the dataset dictionary directly
            if self.current_dataset_name not in self.datasets:
                self.datasets[self.current_dataset_name] = {}
            dataset_dict = self.datasets[self.current_dataset_name]
            self.header_data = dataset_dict  # Point to the same dictionary
            
            # Convert field name to string for consistent storage
            new_field_str = str(new_field).strip()
            
            # If editing existing field, delete old key (try both string and original type)
            if field_name:
                old_field_str = str(field_name).strip()
                # Try to find and delete old key regardless of type
                for key in list(dataset_dict.keys()):
                    if str(key).strip() == old_field_str:
                        del dataset_dict[key]
                        break
            
            # Always store with string key
            dataset_dict[new_field_str] = new_value
            self.load_data()
            self.auto_save()
            dialog.destroy()

        def cancel_changes():
            dialog.destroy()

        tk.Button(button_frame, text="Save", bg="#27ae60", fg="white", font=("Arial", 10, "bold"), relief="flat", padx=20, pady=5, command=save_changes).pack(side="left", padx=5)
        tk.Button(button_frame, text="Cancel", bg="#e74c3c", fg="white", font=("Arial", 10, "bold"), relief="flat", padx=20, pady=5, command=cancel_changes).pack(side="left", padx=5)

    def show_filter_edit_dialog(self, index=None, filter_value=""):
        """Edit dialog for list-type filter datasets"""
        is_edit = index is not None
        dialog_title = f"Edit {self.current_dataset_name.replace('_', ' ').title()}" if is_edit else f"Add {self.current_dataset_name.replace('_', ' ').title()}"
        
        dialog = tk.Toplevel(self)
        dialog.title(dialog_title)
        dialog.geometry("420x180")
        dialog.configure(bg="#ecf0f1")
        dialog.transient(self)
        dialog.grab_set()
        dialog.geometry("+%d+%d" % (self.winfo_rootx() + 50, self.winfo_rooty() + 50))
        
        # Filter value input
        value_frame = tk.Frame(dialog, bg="#ecf0f1")
        value_frame.pack(fill="x", padx=20, pady=20)
        tk.Label(value_frame, text="Filter Value:", font=("Arial", 12), bg="#ecf0f1").pack(anchor="w")
        value_entry = tk.Entry(value_frame, font=("Arial", 12), width=30)
        value_entry.pack(fill="x", pady=5)
        value_entry.insert(0, str(filter_value))
        value_entry.focus()
        
        button_frame = tk.Frame(dialog, bg="#ecf0f1")
        button_frame.pack(fill="x", padx=20, pady=20)
        
        def save_changes():
            new_value = value_entry.get().strip()
            if not new_value:
                messagebox.showerror("Error", "Filter value is required")
                return
            
            # Ensure dataset is a list
            if self.current_dataset_name not in self.datasets:
                self.datasets[self.current_dataset_name] = []
            if not isinstance(self.datasets[self.current_dataset_name], list):
                self.datasets[self.current_dataset_name] = []
            
            if is_edit and index is not None:
                # Edit existing item
                if 0 <= index < len(self.datasets[self.current_dataset_name]):
                    self.datasets[self.current_dataset_name][index] = new_value
                else:
                    messagebox.showerror("Error", "Invalid index for editing")
                    return
            else:
                # Add new item
                self.datasets[self.current_dataset_name].append(new_value)
            
            self.header_data = self.datasets[self.current_dataset_name]
            self.load_data()
            self.auto_save()
            dialog.destroy()
        
        def cancel_changes():
            dialog.destroy()
        
        tk.Button(button_frame, text="Save", bg="#27ae60", fg="white", font=("Arial", 10, "bold"), relief="flat", padx=20, pady=5, command=save_changes).pack(side="left", padx=5)
        tk.Button(button_frame, text="Cancel", bg="#e74c3c", fg="white", font=("Arial", 10, "bold"), relief="flat", padx=20, pady=5, command=cancel_changes).pack(side="left", padx=5)

    def configure_tree_for_mode(self):
        for col in self.tree["columns"]:
            try:
                self.tree.heading(col, text="")
                self.tree.column(col, width=0)
            except Exception:
                pass
        if self.mode_var.get() == "LOTSIZE":
            cols = ("Symbol", "Lotsize")
            self.tree.configure(columns=cols)
            self.tree.heading("Symbol", text="Symbol")
            self.tree.heading("Lotsize", text="Lotsize")
            self.tree.column("Symbol", width=200, anchor="w")
            self.tree.column("Lotsize", width=100, anchor="center")
        elif self.mode_var.get() == "UNDERLYINGCODE":
            cols = ("Symbol", "Code")
            self.tree.configure(columns=cols)
            self.tree.heading("Symbol", text="Symbol")
            self.tree.heading("Code", text="Code")
            self.tree.column("Symbol", width=200, anchor="w")
            self.tree.column("Code", width=100, anchor="center")
        elif self.current_dataset_name == "fund_filename_map":
            # Special 4-column layout for fund_filename_map
            cols = ("Custodian Account", "Fund Name", "Fund CDS Name", "Password")
            self.tree.configure(columns=cols)
            self.tree.heading("Custodian Account", text="Custodian Account")
            self.tree.heading("Fund Name", text="Fund Name")
            self.tree.heading("Fund CDS Name", text="Fund CDS Name")
            self.tree.heading("Password", text="Password")
            self.tree.column("Custodian Account", width=150, anchor="w")
            self.tree.column("Fund Name", width=200, anchor="w")
            self.tree.column("Fund CDS Name", width=200, anchor="w")
            self.tree.column("Password", width=120, anchor="w")
        elif self.current_dataset_name in ["mcx_group2_filters", "fno_group2_filters"]:
            # Special 2-column layout for list-type filter datasets
            cols = ("Index", "Filter Value")
            self.tree.configure(columns=cols)
            self.tree.heading("Index", text="Sr No")
            self.tree.heading("Filter Value", text="Filter Value")
            self.tree.column("Index", width=80, anchor="center")
            self.tree.column("Filter Value", width=400, anchor="w")
        else:
            cols = ("Header", "Value")
            self.tree.configure(columns=cols)
            
            # Customize column headers for different datasets
            if self.current_dataset_name == "asio_recon_portfolio_mapping":
                self.tree.heading("Header", text="Portfolio")
                self.tree.heading("Value", text="Holding Name")
            elif self.current_dataset_name == "geneva_custodian_mapping":
                self.tree.heading("Header", text="Fund name in Geneva")
                self.tree.heading("Value", text="Fund name in Cust Holding")
            elif self.current_dataset_name == "fno_tm_code_with_tm_name":
                self.tree.heading("Header", text="TM Code")
                self.tree.heading("Value", text="TM Name")
            elif self.current_dataset_name == "mcx_tm_code_with_tm_name":
                self.tree.heading("Header", text="TM Code")
                self.tree.heading("Value", text="TM Name")
            else:
                self.tree.heading("Header", text="Header")
                self.tree.heading("Value", text="Value")
            
            self.tree.column("Header", width=220, anchor="w")
            self.tree.column("Value", width=260, anchor="w")