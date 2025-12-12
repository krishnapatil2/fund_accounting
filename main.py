import tkinter as tk
import sys
import os
from pages import dashboard, settings, alpha_report, dataconfig, asio_reconciliation, fo_reconciliation, excel_merger, asio_trade_loader, asio_trade_loader_mcx, fno_mcx_price_recon_loader, asio_sub_fund4
from PIL import Image, ImageTk  # for better image support
from file_utils import ensure_consolidated_data_file


# Company Theme
THEME_COLOR = "#307356"
HOVER_COLOR = "#23533E"
BG_COLOR = "#F9F9F9"

MENU_STRUCTURE = {
    "Dashboard": dashboard.DashboardPage,
    "Data Config": dataconfig.DataConfigPage,
    "Settings": settings.SettingsPage,
    "Process": {  # Dropdown Menu
        "Alpha Report": alpha_report.AlphaReportPage,
        "ASIO Reconciliation": asio_reconciliation.ASIOReconciliationPage,
        "ASIO Sub Fund 2 Trade Loader FNO": asio_trade_loader.ASIOTradeLoaderPage,
        "ASIO Sub Fund 2 Trade Loader MCX": asio_trade_loader_mcx.ASIOTradeLoaderMCXPage,
        "ASIO Sub Fund 4": asio_sub_fund4.ASIOSubFund4Page,
        "Daily F&O Reconciliation": fo_reconciliation.FOReconciliationPage,
        "FNO and MCX Price Recon & Loader": fno_mcx_price_recon_loader.FNOMCXPriceReconLoaderPage,
        "Excel Merger": excel_merger.ExcelMergerPage,
    }
}


class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        
        # Ensure consolidated data file exists
        ensure_consolidated_data_file()
        
        self.title("Fund Accounting App")
        self.geometry("1000x600")
        self.configure(bg=BG_COLOR)

        # Top Nav
        self.navbar = tk.Frame(self, bg=THEME_COLOR, height=50)
        self.navbar.pack(side="top", fill="x")

        # Try to set icon, but don't fail if it's not found
        try:
            icon_path = self._get_icon_path()            
            # Try iconbitmap first
            try:
                self.iconbitmap(icon_path)
            except:
                # If iconbitmap fails, try iconphoto
                try:
                    icon_image = Image.open(icon_path)
                    icon_photo = ImageTk.PhotoImage(icon_image)
                    self.iconphoto(False, icon_photo)
                except Exception as photo_error:
                    raise photo_error
                    
        except Exception as e:
            # Icon not found, continue without it
            pass

        # Content
        self.content = tk.Frame(self, bg=BG_COLOR)
        self.content.pack(fill="both", expand=True)

        self.open_menu = None

        self.generate_menu(MENU_STRUCTURE)
        self.show_page(dashboard.DashboardPage)

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
        else:  # Direct page
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

        # Add menu items
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
        for widget in self.content.winfo_children():
            widget.destroy()

    def show_page(self, page_class):
        if self.open_menu:
            self.open_menu.unpost()
            self.open_menu = None
        self.clear_content()
        page = page_class(self.content)
        page.pack(fill="both", expand=True)


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