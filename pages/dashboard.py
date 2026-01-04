import tkinter as tk
from tkinter import ttk
import os
import json
from datetime import datetime

class DashboardPage(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg="#ecf0f1")
        self.parent = parent
        # Try to get main app reference for navigation
        self.main_app = self._get_main_app()
        
        # Create scrollable canvas with scrollbar
        # Create a frame to hold canvas and scrollbar
        canvas_frame = tk.Frame(self, bg="#ecf0f1")
        canvas_frame.pack(fill="both", expand=True)
        
        self.canvas = tk.Canvas(canvas_frame, bg="#ecf0f1", highlightthickness=0)
        self.scrollable_frame = tk.Frame(self.canvas, bg="#ecf0f1")
        
        # Create vertical scrollbar
        self.scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Throttle scroll region updates to prevent hanging
        self._scroll_update_pending = False
        
        def update_scroll_region(event=None):
            """Update scroll region with throttling to prevent UI freezing"""
            if not self._scroll_update_pending:
                self._scroll_update_pending = True
                # Schedule update after a short delay to batch multiple updates
                self.after(10, self._do_scroll_update)
        
        self.scrollable_frame.bind("<Configure>", update_scroll_region)
        
        # Create window in canvas for scrollable_frame
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        
        # Configure canvas scrolling with scrollbar
        
        # Make canvas window width match canvas width (throttled)
        self._width_update_pending = False
        
        def configure_canvas_width(event):
            """Update canvas window width with throttling"""
            if not self._width_update_pending:
                self._width_update_pending = True
                canvas_width = event.width
                self.after(10, lambda w=canvas_width: self._do_width_update(w))
        
        self.canvas.bind("<Configure>", configure_canvas_width)
        
        # Bind mouse wheel events will be called after all widgets are created
        
        # Title Section
        title_frame = tk.Frame(self.scrollable_frame, bg="#ecf0f1")
        title_frame.pack(fill="x", padx=20, pady=(20, 10))
        
        title_label = tk.Label(title_frame, text="📊 Dashboard", 
                              font=("Arial", 28, "bold"), bg="#ecf0f1", fg="#2c3e50")
        title_label.pack(anchor="w")
        
        subtitle_label = tk.Label(title_frame, text="Fund Accounting System Overview", 
                                 font=("Arial", 11), bg="#ecf0f1", fg="#7f8c8d")
        subtitle_label.pack(anchor="w", pady=(5, 0))
        
        # Get data from consolidated_data.json
        config_data = self.load_config_data()
        
        # Quick Access Section
        quick_access_frame = tk.LabelFrame(self.scrollable_frame, text="🚀 Quick Access", 
                                          font=("Arial", 14, "bold"),
                                          bg="#ecf0f1", fg="#2c3e50", padx=20, pady=15)
        quick_access_frame.pack(fill="x", padx=20, pady=(0, 15))
        self.create_quick_access_section(quick_access_frame)
        
        # Dataset Overview Section
        dataset_frame = tk.LabelFrame(self.scrollable_frame, text="📋 Dataset Overview", 
                                     font=("Arial", 14, "bold"),
                                     bg="#ecf0f1", fg="#2c3e50", padx=20, pady=15)
        dataset_frame.pack(fill="x", padx=20, pady=(0, 15))
        self.create_dataset_overview(dataset_frame, config_data)
        
        # Pack canvas and scrollbar side by side
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # Bind keyboard events for scrolling
        self._bind_keyboard_scrolling()
        
        # Bind mousewheel events after all widgets are created
        self.after(300, self._bind_mousewheel_global)
        
        # Update scroll region after all widgets are packed (deferred for smooth loading)
        self.after(100, self._finalize_scroll_region)
    
    def _do_scroll_update(self):
        """Actually perform the scroll region update"""
        try:
            # Only update if canvas still exists
            if self.canvas.winfo_exists():
                bbox = self.canvas.bbox("all")
                if bbox:
                    self.canvas.configure(scrollregion=bbox)
        finally:
            self._scroll_update_pending = False
    
    def _do_width_update(self, width):
        """Actually update the canvas window width"""
        try:
            if self.canvas.winfo_exists():
                self.canvas.itemconfig(self.canvas_window, width=width)
        finally:
            self._width_update_pending = False
    
    def _finalize_scroll_region(self):
        """Finalize scroll region after all content is loaded"""
        try:
            if self.canvas.winfo_exists():
                bbox = self.canvas.bbox("all")
                if bbox:
                    self.canvas.configure(scrollregion=bbox)
        except Exception:
            pass
    
    def load_config_data(self):
        """Load data from consolidated_data.json"""
        try:
            from my_app.file_utils import get_app_directory
            app_dir = get_app_directory()
            consolidated_path = os.path.join(app_dir, "consolidated_data.json")
            
            if os.path.exists(consolidated_path):
                with open(consolidated_path, "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception as e:
            print(f"Error loading config: {e}")
        return {}
    
    def _discover_datasets(self, config_data):
        """Dynamically discover all datasets from consolidated_data.json"""
        datasets = set()
        
        # Get ALL keys from consolidated_data.json (these are actual datasets)
        # This is the source of truth - includes all datasets that exist
        if isinstance(config_data, dict):
            for key in config_data.keys():
                # Exclude non-dataset keys (these are configuration data, not datasets)
                if key not in ["lotsize_data", "underlying_code_data"]:
                    datasets.add(key)
        
        # Also check for individual JSON files in app directory (additional datasets)
        try:
            from my_app.file_utils import get_app_directory
            app_dir = get_app_directory()
            if os.path.exists(app_dir):
                for fname in os.listdir(app_dir):
                    if fname.lower().endswith(".json"):
                        if fname not in ["lotsize_data.json", "consolidated_data.json"]:
                            dataset_name = os.path.splitext(os.path.basename(fname))[0]
                            datasets.add(dataset_name)
        except Exception:
            pass
        
        return sorted(list(datasets))
    
    def create_statistics_cards(self, parent, config_data):
        """Create statistics cards showing datasets with mapping counts and reports"""
        # Dynamically discover all datasets from consolidated_data.json
        all_datasets = self._discover_datasets(config_data)
        
        # Count total datasets - show ALL datasets that exist
        total_datasets = len(all_datasets)
        
        # Count configured datasets (datasets that have data/content)
        configured_count = sum(
            1 for ds in all_datasets 
            if ds in config_data and config_data.get(ds) and 
            (isinstance(config_data[ds], dict) and len(config_data[ds]) > 0 or 
             isinstance(config_data[ds], list) and len(config_data[ds]) > 0 or
             not isinstance(config_data[ds], (dict, list)))
        )
        
        # Count reports/modules
        report_count = 8  # Based on MENU_STRUCTURE
        
        # Card color - using #367a58 for reports card
        bg_color = "#367a58"
        border_color = "#2d6346"  # Slightly darker for border
        
        # Create card container (single card now)
        cards_container = tk.Frame(parent, bg="#ecf0f1")
        cards_container.pack(fill="x")
        
        # Create single reports card (small, centered)
        card = tk.Frame(cards_container, bg=border_color, relief="flat", bd=2, width=250)
        card.pack(padx=10, pady=5)
        card.pack_propagate(False)  # Prevent card from resizing to content
        
        inner = tk.Frame(card, bg=bg_color, padx=20, pady=15)
        inner.pack(fill="both", expand=True)
        
        title_label = tk.Label(inner, text="📄 Reports", font=("Arial", 11, "bold"), 
                              bg=bg_color, fg="white")
        title_label.pack(anchor="w")
        
        value_label = tk.Label(inner, text=str(report_count), font=("Arial", 24, "bold"), 
                              bg=bg_color, fg="white")
        value_label.pack(anchor="w", pady=(5, 0))
        
        subtitle_label = tk.Label(inner, text="Available", font=("Arial", 9), 
                                  bg=bg_color, fg="white")
        subtitle_label.pack(anchor="w", pady=(2, 0))
    
    def create_quick_access_section(self, parent):
        """Create quick access buttons for reports"""
        content = tk.Frame(parent, bg="#ecf0f1")
        content.pack(fill="both", expand=True)
        
        # Report modules with their icons and descriptions
        reports = [
            ("📈 Alpha Report", "Process trade data and generate security reports"),
            ("🔄 ASIO Reconciliation", "Reconcile ASIO portfolio data"),
            ("📊 ASIO SF2 Trade Loader FNO", "Process FNO trade files"),
            ("📊 ASIO SF2 Trade Loader MCX", "Process MCX trade files"),
            ("📋 ASIO Sub Fund 4", "Process Sub Fund 4 trade files"),
            ("📈 Daily F&O Reconciliation", "Daily F&O reconciliation process"),
            ("💰 FNO/MCX Price Recon", "Price reconciliation reports"),
            ("📑 GTN Loader", "Process GTN trade files"),
            ("📑 Excel Merger", "Merge multiple Excel files"),
            ("📥 Download NSE Bhavcopy", "Download NSE equity bhavcopy for a specific date"),
        ]
        
        # Create buttons in a grid
        buttons_frame = tk.Frame(content, bg="#ecf0f1")
        buttons_frame.pack(fill="x", padx=10, pady=10)
        
        for i, (title, desc) in enumerate(reports):
            row = i // 2
            col = i % 2
            
            btn_frame = tk.Frame(buttons_frame, bg="#ecf0f1", relief="flat", bd=1)
            btn_frame.grid(row=row, column=col, padx=10, pady=8, sticky="ew")
            buttons_frame.grid_columnconfigure(col, weight=1)
            
            btn = tk.Button(btn_frame, text=title, font=("Arial", 11, "bold"),
                           bg="#668871", fg="white", relief="flat", padx=20, pady=12,
                           cursor="hand2", anchor="w",
                           activebackground="#5a7579", activeforeground="white",
                           command=lambda t=title: self._navigate_to_report(t))
            btn.pack(fill="x")
            
            desc_label = tk.Label(btn_frame, text=desc, font=("Arial", 9),
                                 bg="#ecf0f1", fg="#7f8c8d", anchor="w")
            desc_label.pack(fill="x", padx=(20, 10), pady=(5, 10))
    
    def _bind_mousewheel_global(self):
        """Bind mousewheel events for smooth scrolling anywhere on page"""
        # Bind to canvas for Windows/Mac - primary binding
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        # Bind to canvas for Linux
        self.canvas.bind("<Button-4>", lambda e: self._on_mousewheel_linux(-1))
        self.canvas.bind("<Button-5>", lambda e: self._on_mousewheel_linux(1))
        
        # Bind to scrollable_frame so scrolling works when mouse is over content
        self.scrollable_frame.bind("<MouseWheel>", self._on_mousewheel)
        self.scrollable_frame.bind("<Button-4>", lambda e: self._on_mousewheel_linux(-1))
        self.scrollable_frame.bind("<Button-5>", lambda e: self._on_mousewheel_linux(1))
        
        # Bind to self (dashboard frame) to catch events anywhere in dashboard
        self.bind("<MouseWheel>", self._on_mousewheel)
        self.bind("<Button-4>", lambda e: self._on_mousewheel_linux(-1))
        self.bind("<Button-5>", lambda e: self._on_mousewheel_linux(1))
        
        # Recursively bind to all existing child widgets - this ensures scrolling works
        # on all widgets including labels, frames, buttons, etc.
        self._bind_mousewheel_recursive(self.scrollable_frame)
        
        # Use a more direct approach - bind to the root window when mouse is over dashboard
        root = self.winfo_toplevel()
        if root:
            def dashboard_mousewheel(event):
                """Handle mousewheel when over dashboard"""
                try:
                    # Get widget under cursor
                    x, y = root.winfo_pointerxy()
                    widget = root.winfo_containing(x, y)
                    # Check if widget is part of this dashboard
                    current = widget
                    while current:
                        if current == self or current == self.canvas or current == self.scrollable_frame:
                            self._on_mousewheel(event)
                            return "break"
                        if hasattr(current, 'master'):
                            current = current.master
                        else:
                            break
                except Exception:
                    pass
            
            # Store reference to avoid garbage collection
            self._dashboard_mousewheel_handler = dashboard_mousewheel
            root.bind_all("<MouseWheel>", dashboard_mousewheel)
            root.bind_all("<Button-4>", lambda e: dashboard_mousewheel(e) if e else None)
            root.bind_all("<Button-5>", lambda e: dashboard_mousewheel(e) if e else None)
    
    def _bind_mousewheel_recursive(self, widget):
        """Recursively bind mousewheel to widget and its children"""
        try:
            # Bind to current widget
            widget.bind("<MouseWheel>", self._on_mousewheel, add="+")
            widget.bind("<Button-4>", lambda e: self._on_mousewheel_linux(-1), add="+")
            widget.bind("<Button-5>", lambda e: self._on_mousewheel_linux(1), add="+")
            
            # Bind to all child widgets recursively
            for child in widget.winfo_children():
                self._bind_mousewheel_recursive(child)
        except Exception:
            pass
    
    def _on_mousewheel(self, event):
        """Handle mousewheel scrolling with smooth scrolling"""
        # Check if mouse is over the canvas or its contents
        try:
            # Calculate scroll amount - use smaller increments for smoother scrolling
            delta = int(-1 * (event.delta / 120))
            # Use "units" for smooth scrolling (can also use "pages" for faster)
            self.canvas.yview_scroll(delta, "units")
            # Prevent event propagation
            return "break"
        except Exception:
            pass
    
    def _on_mousewheel_linux(self, direction):
        """Handle Linux mousewheel scrolling"""
        try:
            self.canvas.yview_scroll(direction, "units")
            return "break"
        except Exception:
            pass
    

    def _bind_keyboard_scrolling(self):
        """Bind keyboard events for scrolling with arrow keys"""
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
    
    def _get_main_app(self):
        """Try to get main app reference by traversing widget hierarchy"""
        widget = self
        # Traverse up to find the root window (MainApp)
        while widget:
            if hasattr(widget, 'show_page'):
                return widget
            widget = widget.master if hasattr(widget, 'master') else None
            # Stop at root window
            if isinstance(widget, tk.Tk):
                if hasattr(widget, 'show_page'):
                    return widget
                break
        return None
    
    def _navigate_to_report(self, report_name):
        """Navigate to report page - optimized with lazy imports"""
        # Lazy import only the specific module needed (not all at once)
        try:
            # Map report names to (module_name, class_name) for lazy loading
            report_to_module = {
                "📈 Alpha Report": ("alpha_report", "AlphaReportPage"),
                "🔄 ASIO Reconciliation": ("asio_reconciliation", "ASIOReconciliationPage"),
                "📊 ASIO SF2 Trade Loader FNO": ("asio_trade_loader", "ASIOTradeLoaderPage"),
                "📊 ASIO SF2 Trade Loader MCX": ("asio_trade_loader_mcx", "ASIOTradeLoaderMCXPage"),
                "📋 ASIO Sub Fund 4": ("asio_sub_fund4", "ASIOSubFund4Page"),
                "📈 Daily F&O Reconciliation": ("fo_reconciliation", "FOReconciliationPage"),
                "💰 FNO/MCX Price Recon": ("fno_mcx_price_recon_loader", "FNOMCXPriceReconLoaderPage"),
                "📑 GTN Loader": ("gtn_loader", "GTNLoaderPage"),
                "📑 Excel Merger": ("excel_merger", "ExcelMergerPage"),
                "📥 Download NSE Bhavcopy": ("bhavcopy_downloader", "BhavcopyDownloaderPage"),
            }
            
            if report_name not in report_to_module:
                return
            
            # Lazy import only the specific module needed
            module_name, class_name = report_to_module[report_name]
            module = __import__(f"pages.{module_name}", fromlist=[class_name])
            page_class = getattr(module, class_name)
            
            if page_class:
                # Try to find main app by traversing to root
                root = self.winfo_toplevel()
                if hasattr(root, 'show_page'):
                    root.show_page(page_class)
                    return
                # Alternative: try parent traversal
                if self.main_app and hasattr(self.main_app, 'show_page'):
                    self.main_app.show_page(page_class)
                    return
        except Exception as e:
            pass
        
        # Fallback: show informative message
        from tkinter import messagebox
        clean_name = report_name.replace("📈 ", "").replace("🔄 ", "").replace("📊 ", "").replace("📋 ", "").replace("💰 ", "").replace("📑 ", "").replace("📥 ", "")
        messagebox.showinfo("Quick Access", 
                          f"To access '{clean_name}', please use the 'Process' menu in the navigation bar.")
    
    def _get_dataset_to_report_mapping(self):
        """Get mapping of datasets to reports/processes"""
        return {
            "trade_headers": ["Alpha Report"],
            "lotsize_data": ["Alpha Report"],
            "aafspl_car_future": ["Alpha Report"],
            "option_security": ["Alpha Report"],
            "car_trade_loader": ["Alpha Report"],
            "asio_recon_format_1_headers": ["ASIO Reconciliation"],
            "asio_recon_format_2_headers": ["ASIO Reconciliation"],
            "asio_recon_bhavcopy_headers": ["ASIO Reconciliation"],
            "asio_recon_portfolio_mapping": ["ASIO Reconciliation"],
            "asio_geneva_headers": ["ASIO Reconciliation"],
            "geneva_custodian_mapping": ["ASIO Reconciliation"],
            "fund_filename_map": ["Alpha Report", "ASIO Reconciliation"],
            "asio_sf_2_trade_loader": ["ASIO Sub Fund 2 Trade Loader FNO"],
            "asio_sf_2_option_security": ["ASIO Sub Fund 2 Trade Loader FNO"],
            "asio_sf_2_future_security": ["ASIO Sub Fund 2 Trade Loader FNO"],
            "fno_tm_code_with_tm_name": ["ASIO Sub Fund 2 Trade Loader FNO"],
            "asio_sf_2_mcx_trade_loader": ["ASIO Sub Fund 2 Trade Loader MCX"],
            "asio_sf_2_mcx_option_security": ["ASIO Sub Fund 2 Trade Loader MCX"],
            "asio_sf_2_mcx_future_security": ["ASIO Sub Fund 2 Trade Loader MCX"],
            "mcx_tm_code_with_tm_name": ["ASIO Sub Fund 2 Trade Loader MCX"],
            "underlying_code_data": ["ASIO Sub Fund 2 Trade Loader MCX"],
            "asio_sf4_ft": ["ASIO Sub Fund 4"],
            "asio_sf4_trading_code_mapping": ["ASIO Sub Fund 4"],
            "asio_sub_fund4_read_config": ["ASIO Sub Fund 4"],
            "asio_pricing_fno": ["FNO and MCX Price Recon & Loader"],
            "asio_pricing_mcx": ["FNO and MCX Price Recon & Loader"],
            "mcx_group2_filters": ["FNO and MCX Price Recon & Loader"],
            "fno_group2_filters": ["FNO and MCX Price Recon & Loader"],
            "GTN_LOADER": ["GTN Loader"],
            "gtn_sp_30_call_option": ["GTN Loader"],
            "gtn_sp_30_put_option": ["GTN Loader"],
        }
    
    def create_dataset_overview(self, parent, config_data):
        """Create dataset overview showing process-wise dataset usage (grouped by reports)"""
        content = tk.Frame(parent, bg="#ecf0f1")
        content.pack(fill="both", expand=True)
        
        # Dynamically discover all datasets
        all_datasets = self._discover_datasets(config_data)
        
        # Get dataset to report mapping
        dataset_to_reports = self._get_dataset_to_report_mapping()
        
        # Group datasets by report/process
        report_to_datasets = {}
        for dataset, reports in dataset_to_reports.items():
            for report in reports:
                if report not in report_to_datasets:
                    report_to_datasets[report] = []
                report_to_datasets[report].append(dataset)
        
        # Also add datasets that don't have a mapping
        unmapped_datasets = [ds for ds in all_datasets if ds not in dataset_to_reports]
        if unmapped_datasets:
            report_to_datasets["Other Datasets"] = unmapped_datasets
        
        # Create sections for each report/process
        for report_name, datasets_list in sorted(report_to_datasets.items()):
            # Report section header
            report_frame = tk.Frame(content, bg="#ecf0f1")
            report_frame.pack(fill="x", padx=10, pady=(10, 5))
            
            report_label = tk.Label(report_frame, 
                                   text=f"📋 {report_name}",
                                   font=("Arial", 12, "bold"), 
                                   bg="#ecf0f1", fg="#2c3e50")
            report_label.pack(anchor="w", pady=(0, 5))
            
            # Datasets for this report
            datasets_frame = tk.Frame(content, bg="#ecf0f1")
            datasets_frame.pack(fill="x", padx=30, pady=(0, 10))
            
            for dataset in sorted(datasets_list):
                ds_frame = tk.Frame(datasets_frame, bg="#ecf0f1")
                ds_frame.pack(fill="x", pady=2)
                
                # Check if dataset exists and get mapping count
                dataset_data = config_data.get(dataset, {})
                if dataset_data:
                    if isinstance(dataset_data, dict):
                        mapping_count = len(dataset_data)
                    elif isinstance(dataset_data, list):
                        mapping_count = len(dataset_data)
                    else:
                        mapping_count = 1 if dataset_data else 0
                    status = "✓"
                    status_color = "#27ae60"
                else:
                    mapping_count = 0
                    status = "○"
                    status_color = "#95a5a6"
                
                # Dataset name with status and mapping count
                ds_text = f"  {status} {dataset}"
                if mapping_count > 0:
                    ds_text += f" ({mapping_count} mappings)"
                
                ds_label = tk.Label(ds_frame, 
                                   text=ds_text,
                                   font=("Arial", 9), 
                                   bg="#ecf0f1", fg=status_color,
                                   anchor="w")
                ds_label.pack(anchor="w")
    
    def create_system_status_section(self, parent, config_data):
        """Create system status indicators"""
        content = tk.Frame(parent, bg="#ecf0f1")
        content.pack(fill="both", expand=True)
        
        # Check consolidated_data.json
        try:
            from my_app.file_utils import get_app_directory
            app_dir = get_app_directory()
            consolidated_path = os.path.join(app_dir, "consolidated_data.json")
            config_exists = os.path.exists(consolidated_path)
            
            if config_exists:
                file_size = os.path.getsize(consolidated_path)
                file_size_kb = file_size / 1024
                modified_time = datetime.fromtimestamp(os.path.getmtime(consolidated_path))
                modified_str = modified_time.strftime("%Y-%m-%d %H:%M:%S")
            else:
                file_size_kb = 0
                modified_str = "N/A"
        except:
            config_exists = False
            file_size_kb = 0
            modified_str = "N/A"
        
        # Status items
        status_items = [
            ("Configuration File", "✓ Ready" if config_exists else "⚠ Missing", 
             "#27ae60" if config_exists else "#e74c3c"),
            ("File Size", f"{file_size_kb:.2f} KB" if config_exists else "N/A", "#34495e"),
            ("Last Modified", modified_str, "#34495e"),
            ("System Status", "✓ Operational", "#27ae60"),
        ]
        
        for label_text, value_text, color in status_items:
            item_frame = tk.Frame(content, bg="#ecf0f1")
            item_frame.pack(fill="x", padx=10, pady=5)
            
            label = tk.Label(item_frame, text=f"{label_text}:", 
                           font=("Arial", 10, "bold"), bg="#ecf0f1", fg="#2c3e50",
                           width=15, anchor="w")
            label.pack(side="left")
            
            value = tk.Label(item_frame, text=value_text, 
                           font=("Arial", 10), bg="#ecf0f1", fg=color,
                           anchor="w")
            value.pack(side="left", padx=(10, 0))
    
    def check_system_status(self):
        """Check if consolidated_data.json exists"""
        try:
            from my_app.file_utils import get_app_directory
            app_dir = get_app_directory()
            consolidated_path = os.path.join(app_dir, "consolidated_data.json")
            
            if os.path.exists(consolidated_path):
                return "✓ Ready - All configurations loaded"
            else:
                return "⚠ Configuration file will be created on first use"
        except:
            return "System Ready"
