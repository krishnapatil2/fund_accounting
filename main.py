import tkinter as tk
import sys
import os
import threading
# LAZY IMPORTS - Only import when needed for faster startup
from file_utils import ensure_consolidated_data_file

# Company Theme
THEME_COLOR = "#307356"
HOVER_COLOR = "#23533E"
BG_COLOR = "#F9F9F9"

# Lazy page loader - pages imported only when accessed
_page_cache = {}

def _get_page_class(page_name):
    """Lazy load page classes to speed up startup"""
    if page_name in _page_cache:
        return _page_cache[page_name]
    
    # Import only when needed
    if page_name == "dashboard":
        from pages import dashboard
        _page_cache[page_name] = dashboard.DashboardPage
        return dashboard.DashboardPage
    elif page_name == "dataconfig":
        from pages import dataconfig
        _page_cache[page_name] = dataconfig.DataConfigPage
        return dataconfig.DataConfigPage
    elif page_name == "settings":
        from pages import settings
        _page_cache[page_name] = settings.SettingsPage
        return settings.SettingsPage
    elif page_name == "alpha_report":
        from pages import alpha_report
        _page_cache[page_name] = alpha_report.AlphaReportPage
        return alpha_report.AlphaReportPage
    elif page_name == "asio_reconciliation":
        from pages import asio_reconciliation
        _page_cache[page_name] = asio_reconciliation.ASIOReconciliationPage
        return asio_reconciliation.ASIOReconciliationPage
    elif page_name == "asio_trade_loader":
        from pages import asio_trade_loader
        _page_cache[page_name] = asio_trade_loader.ASIOTradeLoaderPage
        return asio_trade_loader.ASIOTradeLoaderPage
    elif page_name == "asio_trade_loader_mcx":
        from pages import asio_trade_loader_mcx
        _page_cache[page_name] = asio_trade_loader_mcx.ASIOTradeLoaderMCXPage
        return asio_trade_loader_mcx.ASIOTradeLoaderMCXPage
    elif page_name == "asio_sub_fund4":
        from pages import asio_sub_fund4
        _page_cache[page_name] = asio_sub_fund4.ASIOSubFund4Page
        return asio_sub_fund4.ASIOSubFund4Page
    elif page_name == "fo_reconciliation":
        from pages import fo_reconciliation
        _page_cache[page_name] = fo_reconciliation.FOReconciliationPage
        return fo_reconciliation.FOReconciliationPage
    elif page_name == "fno_mcx_price_recon_loader":
        from pages import fno_mcx_price_recon_loader
        _page_cache[page_name] = fno_mcx_price_recon_loader.FNOMCXPriceReconLoaderPage
        return fno_mcx_price_recon_loader.FNOMCXPriceReconLoaderPage
    elif page_name == "excel_merger":
        from pages import excel_merger
        _page_cache[page_name] = excel_merger.ExcelMergerPage
        return excel_merger.ExcelMergerPage
    return None

# Menu structure with lazy loading
MENU_STRUCTURE = {
    "Dashboard": "dashboard",
    "Data Config": "dataconfig",
    "Settings": "settings",
    "Process": {  # Dropdown Menu
        "Alpha Report": "alpha_report",
        "ASIO Reconciliation": "asio_reconciliation",
        "ASIO Sub Fund 2 Trade Loader FNO": "asio_trade_loader",
        "ASIO Sub Fund 2 Trade Loader MCX": "asio_trade_loader_mcx",
        "ASIO Sub Fund 4": "asio_sub_fund4",
        "Daily F&O Reconciliation": "fo_reconciliation",
        "FNO and MCX Price Recon & Loader": "fno_mcx_price_recon_loader",
        "Excel Merger": "excel_merger",
    }
}


class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        
        # OPTIMIZATION: Update window immediately before heavy operations
        self.update_idletasks()
        
        # Show UI immediately for faster perceived startup
        self.title("Fund Accounting App")
        self.geometry("1000x600")
        self.configure(bg=BG_COLOR)

        # Top Nav
        self.navbar = tk.Frame(self, bg=THEME_COLOR, height=50)
        self.navbar.pack(side="top", fill="x")

        # Content
        self.content = tk.Frame(self, bg=BG_COLOR)
        self.content.pack(fill="both", expand=True)

        self.open_menu = None

        # Generate menu immediately
        self.generate_menu(MENU_STRUCTURE)
        
        # Force window to appear immediately
        self.update()
        
        # Load icon and data file in background (non-blocking)
        self.after(100, self._load_resources)
        
        # Show dashboard immediately (lazy loaded)
        self.after(50, lambda: self.show_page("dashboard"))
    
    def _load_resources(self):
        """Load resources in background to not block UI"""
        try:
            # Load icon (non-blocking)
            self._load_icon()
        except:
            pass
        
        # Ensure consolidated data file exists (in background)
        try:
            ensure_consolidated_data_file()
        except:
            pass

    def _get_icon_path(self):
        """Get icon path for the application"""
        if getattr(sys, 'frozen', False):
            return os.path.join(sys._MEIPASS, "favicon.ico")
        else:
            # Try multiple possible locations for the icon
            possible_paths = [
                "favicon.ico",  # Same directory (my_app)
                "../logo.ico",  # Logo in parent directory
                "logo.ico"  # Same directory
            ]
            for path in possible_paths:
                if os.path.exists(path):
                    return os.path.abspath(path)
            # Fallback to favicon in assets
            return os.path.abspath("../favicon.ico")
    
    def _load_icon(self):
        """Load icon asynchronously"""
        try:
            icon_path = self._get_icon_path()
            # Try iconbitmap first
            try:
                self.iconbitmap(icon_path)
            except:
                # If iconbitmap fails, try iconphoto (lazy import PIL)
                try:
                    from PIL import Image, ImageTk
                    icon_image = Image.open(icon_path)
                    icon_photo = ImageTk.PhotoImage(icon_image)
                    self.iconphoto(False, icon_photo)
                except Exception:
                    pass
        except Exception:
            pass

    def generate_menu(self, menu_dict):
        for name, target in menu_dict.items():
            # Create button with common styling
            btn = tk.Button(
                self.navbar, text=name,
                bg=THEME_COLOR, fg="white",
                activebackground=HOVER_COLOR, activeforeground="white",
                relief="flat", padx=15, pady=10,
                command=lambda t=target, n=name: self.handle_menu_click(t, n)
            )
            btn.pack(side="left", padx=5)

    def handle_menu_click(self, target, name):
        """Handle menu button clicks - either show page or toggle dropdown"""
        if isinstance(target, dict):  # Dropdown menu
            self.toggle_dropdown(target, name)
        else:  # Direct page (string name for lazy loading)
            self.show_page(target)

    def toggle_dropdown(self, submenu_dict, btn_name):
        """Show/hide dropdown menu on click"""
        if self.open_menu:  # already open → close
            self.open_menu.unpost()
            self.open_menu = None
            return

        # Create dropdown menu
        dropdown = tk.Menu(
            self, tearoff=0,
            bg=HOVER_COLOR, fg="white",
            activebackground="#4E7D66", activeforeground="white"
        )

        # Add menu items (handle lazy loading)
        for sub_name, sub_target in submenu_dict.items():
            dropdown.add_command(
                label=sub_name,
                command=lambda t=sub_target: self.show_page(t)
            )

        # Position and show dropdown
        btn_widget = self.find_button_by_name(btn_name)
        if btn_widget:
            x = btn_widget.winfo_rootx()
            y = btn_widget.winfo_rooty() + btn_widget.winfo_height()
            dropdown.post(x, y)
            self.open_menu = dropdown

    def find_button_by_name(self, name):
        """Find button widget by name"""
        for widget in self.navbar.winfo_children():
            if widget.cget("text") == name:
                return widget
        return None

    def clear_content(self):
        """Fast widget destruction - optimized for speed"""
        widgets = list(self.content.winfo_children())
        for widget in widgets:
            widget_type = widget.winfo_class()
            # Fast clear complex widgets before destroying
            if widget_type == "Treeview":
                try:
                    items = list(widget.get_children())
                    if items:
                        widget.delete(*items)  # Batch delete
                except:
                    pass
            elif widget_type == "Canvas":
                try:
                    widget.delete("all")
                except:
                    pass
            widget.destroy()

    def show_page(self, page_identifier):
        """Show a page - optimized for speed"""
        if self.open_menu:
            self.open_menu.unpost()
            self.open_menu = None
        
        # Fast clear
        self.clear_content()
        
        # Handle lazy loading - if string, get the class
        if isinstance(page_identifier, str):
            page_class = _get_page_class(page_identifier)
        else:
            page_class = page_identifier
        
        if page_class:
            # Create and pack immediately
            page = page_class(self.content)
            page.pack(fill="both", expand=True)
            # Force UI update for instant visual feedback
            self.update_idletasks()


if __name__ == "__main__":
    app = MainApp()
    app.mainloop()


# i need to develope code like this 
# multiple files 
# prepare dict

# actualy cds holdings have 3 files then create one dataframe...
# CDS_HOLDINGS : datafreame
# REGULAR_HOLDINGS : dataframe
# GENEVA_HOLDINGS : dataframe
# NSE_F_AND_O_BHAVCOPY : datafrmae