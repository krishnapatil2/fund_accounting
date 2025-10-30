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
        self.dataset_var = tk.StringVar(value="Lotsize")
        self.dataset_combo = ttk.Combobox(
            selector_frame,
            state="readonly",
            width=28,
            values=["Lotsize"] + list(self.dataset_files.keys()),
            textvariable=self.dataset_var,
        )
        self.dataset_combo.pack(side="left", padx=(6, 0))
        refresh_btn = tk.Button(selector_frame, text="↻", width=3, command=self.refresh_datasets, bg="#ecf0f1", relief="flat")
        refresh_btn.pack(side="left", padx=(6, 0))

        def on_dataset_change(event=None):
            sel = self.dataset_combo.get()
            if sel == "Lotsize":
                self.mode_var.set("LOTSIZE")
            else:
                # Switch to header mode and point to the selected dataset
                self.mode_var.set("HEADER")
                self.current_dataset_name = sel
                # Ensure dataset exists in memory
                if sel not in self.datasets:
                    self.datasets[sel] = self._load_default_dataset_data(sel)
                self.header_data = self.datasets[sel]
            self.configure_tree_for_mode()
            self.load_data()

        self.dataset_combo.bind("<<ComboboxSelected>>", on_dataset_change)

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
            expected_datasets = ["fund_filename_map", "asio_portfolio_mapping", "asio_format_1_headers", "asio_format_2_headers", 
                                "asio_bhavcopy_headers", "asio_geneva_headers", "trade_headers", 
                                "aafspl_car_future", "option_security", "car_trade_loader", "asio_sub_fund_2_trade", "geneva_custodian_mapping"]
            
            if os.path.exists(consolidated_path):
                with open(consolidated_path, "r") as f:
                    consolidated_data = json.load(f)
                    # Load lotsize data
                    lotsize_data = consolidated_data.get("lotsize_data", {})
                    if lotsize_data:
                        self.lotsize_data.update(lotsize_data)
                    
                    # Check for missing datasets and add defaults
                    needs_save = False
                    for dataset_name in expected_datasets:
                        if dataset_name in consolidated_data:
                            self.datasets[dataset_name] = consolidated_data[dataset_name]
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
                    "asio_portfolio_mapping": {
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
                    "asio_format_1_headers": {
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
                    "asio_format_2_headers": {
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
                    "asio_bhavcopy_headers": {
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
                    "asio_sub_fund_2_trade": {
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
                    "asio_sub_fund_2_option_security": {
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
                    }
                }
                
                # Save the default consolidated data file
                with open(consolidated_path, "w") as f:
                    json.dump(default_consolidated_data, f, indent=4)
                
                # Load the data from the newly created file
                self.lotsize_data.update(default_consolidated_data["lotsize_data"])
                for dataset_name in ["fund_filename_map", "asio_portfolio_mapping", "asio_format_1_headers", "asio_format_2_headers", 
                                    "asio_bhavcopy_headers", "asio_geneva_headers", "trade_headers", 
                                    "aafspl_car_future", "option_security", "car_trade_loader", "asio_sub_fund_2_trade", "geneva_custodian_mapping"]:
                    if dataset_name in default_consolidated_data:
                        self.datasets[dataset_name] = default_consolidated_data[dataset_name]
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
        else:
            for header, value in self.header_data.items():
                display_value = json.dumps(value) if isinstance(value, (dict, list)) else value
                self.tree.insert("", "end", values=(header, display_value))
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
        else:
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
            else:
                values = item["values"]
                if self.current_dataset_name == "fund_filename_map":
                    fund_code, fund_name = values[0], values[1] if len(values) > 1 else ""
                    self.selected_info.config(text=f"Selected: {fund_code} - {fund_name}", fg="#2c3e50", font=("Arial", 10, "bold"))
                elif self.current_dataset_name == "asio_portfolio_mapping":
                    header, val = values[0], values[1] if len(values) > 1 else ""
                    self.selected_info.config(text=f"Selected: Portfolio: {header} (Holding Name: {val})", fg="#2c3e50", font=("Arial", 10, "bold"))
                elif self.current_dataset_name == "geneva_custodian_mapping":
                    header, val = values[0], values[1] if len(values) > 1 else ""
                    self.selected_info.config(text=f"Selected: Geneva: {header} | Custodian: {val}", fg="#2c3e50", font=("Arial", 10, "bold"))
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
        elif self.current_dataset_name == "fund_filename_map":
            values = item["values"]
            fund_code = values[0]
            fund_name = values[1] if len(values) > 1 else ""
            cds_name = values[2] if len(values) > 2 else ""
            password = values[3] if len(values) > 3 else ""
            self.show_fund_edit_dialog(fund_code, fund_name, cds_name, password)
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
        else:
            values = item["values"]
            identifier = values[0]  # First column is always the key
            if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete '{identifier}'?"):
                if identifier in self.header_data:
                    del self.header_data[identifier]
                self.load_data()
                self.auto_save()

    def add_new_item(self):
        if self.mode_var.get() == "LOTSIZE":
            self.show_edit_dialog()
        elif self.current_dataset_name == "fund_filename_map":
            self.show_fund_edit_dialog()
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
        elif self.current_dataset_name == "fund_filename_map":
            values = item["values"]
            fund_code = values[0]
            fund_name = values[1] if len(values) > 1 else ""
            cds_name = values[2] if len(values) > 2 else ""
            password = values[3] if len(values) > 3 else ""
            self.show_fund_edit_dialog(fund_code, fund_name, cds_name, password)
        else:
            header, value = item["values"]
            self.show_header_edit_dialog(header, value)

    def show_edit_dialog(self, symbol="", lotsize=""):
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
        lotsize_frame = tk.Frame(dialog, bg="#ecf0f1")
        lotsize_frame.pack(fill="x", padx=20, pady=10)
        tk.Label(lotsize_frame, text="Lotsize:", font=("Arial", 12), bg="#ecf0f1").pack(anchor="w")
        lotsize_entry = tk.Entry(lotsize_frame, font=("Arial", 12), width=30)
        lotsize_entry.pack(fill="x", pady=5)
        lotsize_entry.insert(0, str(lotsize))
        button_frame = tk.Frame(dialog, bg="#ecf0f1")
        button_frame.pack(fill="x", padx=20, pady=20)

        def save_changes():
            new_symbol = symbol_entry.get().strip().upper()
            new_lotsize = lotsize_entry.get().strip()
            if not new_symbol or not new_lotsize:
                messagebox.showerror("Error", "Both Symbol and Lotsize are required")
                return
            try:
                lotsize_value = int(new_lotsize)
            except ValueError:
                messagebox.showerror("Error", "Lotsize must be a valid number")
                return
            if symbol and symbol != new_symbol:
                del self.lotsize_data[symbol]
            self.lotsize_data[new_symbol] = lotsize_value
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
            identifier = values[0]
            if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete '{identifier}'?"):
                if identifier in self.header_data:
                    del self.header_data[identifier]
                self.load_data()
                self.auto_save()

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
            else:
                consolidated_data[self.current_dataset_name] = self.header_data
            
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
            else:
                consolidated_data[self.current_dataset_name] = self.header_data
            
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
        try:
            file_name = self.dataset_files.get(dataset_name, None)
            if file_name and os.path.exists(file_name):
                with open(file_name, "r") as f:
                    return json.load(f)
        except Exception as e:
            print(f"Could not read {dataset_name}: {e}")
        
        # Default fund filename map (empty by default, will be populated from file_utils)
        if dataset_name == "fund_filename_map":
            return {}
        
        # Default ASIO portfolio mapping
        if dataset_name == "asio_portfolio_mapping":
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
        if dataset_name == "asio_format_1_headers":
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
        if dataset_name == "asio_format_2_headers":
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
        if dataset_name == "asio_bhavcopy_headers":
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
        
        # If default header fields are provided and this is the first dataset, seed from it
        first_name = next(iter(self.dataset_files.keys()), None)
        if dataset_name == first_name and DEFAULT_HEADER_FIELDS:
            return {field: "" for field in DEFAULT_HEADER_FIELDS}
        # Otherwise default to empty mapping
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
                    for dataset_name in ["fund_filename_map", "asio_portfolio_mapping", "asio_format_1_headers", "asio_format_2_headers", 
                                        "asio_bhavcopy_headers", "asio_geneva_headers", "trade_headers", 
                                        "aafspl_car_future", "option_security", "car_trade_loader", "asio_sub_fund_2_trade", "geneva_custodian_mapping"]:
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
        for dataset_name in ["fund_filename_map", "asio_portfolio_mapping", "asio_format_1_headers", "asio_format_2_headers", 
                            "asio_bhavcopy_headers", "asio_geneva_headers", "trade_headers", 
                            "aafspl_car_future", "option_security", "car_trade_loader", "asio_sub_fund_2_trade", "geneva_custodian_mapping"]:
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
        # Update combobox options
        values = ["Lotsize"] + list(self.dataset_files.keys())
        self.dataset_combo.configure(values=values)
        # Keep selection consistent in UI
        self.dataset_var.set("Lotsize" if self.mode_var.get() == "LOTSIZE" else self.current_dataset_name)

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
        if self.current_dataset_name == "asio_portfolio_mapping":
            field_label = "Portfolio:"
            value_label = "Holding Name:"
            dialog_title = "Edit Portfolio Mapping" if field_name else "Add Portfolio Mapping"
        elif self.current_dataset_name == "geneva_custodian_mapping":
            field_label = "Fund name in Geneva:"
            value_label = "Fund name in Cust Holding:"
            dialog_title = "Edit Geneva-Custodian Mapping" if field_name else "Add Geneva-Custodian Mapping"
        elif self.current_dataset_name in ["asio_format_1_headers", "asio_format_2_headers", "asio_bhavcopy_headers", "asio_geneva_headers"]:
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
                if self.current_dataset_name == "asio_portfolio_mapping":
                    error_msg = "Portfolio is required"
                elif self.current_dataset_name == "geneva_custodian_mapping":
                    error_msg = "Fund name in Geneva is required"
                elif self.current_dataset_name in ["asio_format_1_headers", "asio_format_2_headers", "asio_bhavcopy_headers", "asio_geneva_headers"]:
                    error_msg = "Header Key is required"
                else:
                    error_msg = "Field is required"
                messagebox.showerror("Error", error_msg)
                return
            try:
                new_value = json.loads(raw)
            except Exception:
                new_value = self._cast_from_string(raw)
            if field_name and field_name != new_field and field_name in self.header_data:
                del self.header_data[field_name]
            self.header_data[new_field] = new_value
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
        else:
            cols = ("Header", "Value")
            self.tree.configure(columns=cols)
            
            # Customize column headers for different datasets
            if self.current_dataset_name == "asio_portfolio_mapping":
                self.tree.heading("Header", text="Portfolio")
                self.tree.heading("Value", text="Holding Name")
            elif self.current_dataset_name == "geneva_custodian_mapping":
                self.tree.heading("Header", text="Fund name in Geneva")
                self.tree.heading("Value", text="Fund name in Cust Holding")
            else:
                self.tree.heading("Header", text="Header")
                self.tree.heading("Value", text="Value")
            
            self.tree.column("Header", width=220, anchor="w")
            self.tree.column("Value", width=260, anchor="w")