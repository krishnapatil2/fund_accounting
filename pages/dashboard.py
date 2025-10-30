import tkinter as tk
from tkinter import ttk
import os
import json

class DashboardPage(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg="#ecf0f1")
        
        # Create scrollable canvas
        canvas = tk.Canvas(self, bg="#ecf0f1", highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="#ecf0f1")
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Title
        title_frame = tk.Frame(scrollable_frame, bg="#ecf0f1")
        title_frame.pack(fill="x", padx=20, pady=20)
        
        lbl = tk.Label(title_frame, text="📊 Dashboard", font=("Arial", 24, "bold"), bg="#ecf0f1", fg="#2c3e50")
        lbl.pack(anchor="w")
        
        # Get data from consolidated_data.json
        config_data = self.load_config_data()
        
        # Single Overview Section
        overview_frame = tk.LabelFrame(scrollable_frame, text="📋 Quick Overview", 
                                      font=("Arial", 14, "bold"),
                                      bg="#ecf0f1", fg="#2c3e50", padx=20, pady=15)
        overview_frame.pack(fill="x", padx=20, pady=(0, 15))
        
        self.create_overview_section(overview_frame, config_data)
        
        # System Status
        status_frame = tk.Frame(scrollable_frame, bg="#ecf0f1")
        status_frame.pack(fill="x", padx=20, pady=20)
        
        status_text = self.check_system_status()
        tk.Label(status_frame, text=f"System Status: {status_text}", 
                font=("Arial", 10, "bold"), bg="#ecf0f1", fg="#27ae60").pack(anchor="w")
        
        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
    def load_config_data(self):
        """Load data from consolidated_data.json"""
        try:
            from my_app.file_utils import get_app_directory
            app_dir = get_app_directory()
            consolidated_path = os.path.join(app_dir, "consolidated_data.json")
            
            if os.path.exists(consolidated_path):
                with open(consolidated_path, "r") as f:
                    return json.load(f)
        except:
            pass
        return {}
    
    def create_overview_section(self, parent, config_data):
        """Create simple overview section"""
        content = tk.Frame(parent, bg="#ecf0f1")
        content.pack(fill="both")
        
        # Data Config Section
        tk.Label(content, text="Data Config:", 
                font=("Arial", 10, "bold"), bg="#ecf0f1", fg="#2c3e50").pack(anchor="w", pady=(0, 5))
        
        tk.Label(content, text="  • Manage portfolio mappings and header configurations for all formats", 
                font=("Arial", 9), bg="#ecf0f1", fg="#34495e", justify="left").pack(anchor="w", pady=1)
        
        # Separator
        tk.Label(content, text="", bg="#ecf0f1").pack(pady=5)
        
        # Reports Section
        tk.Label(content, text="Reports:", 
                font=("Arial", 10, "bold"), bg="#ecf0f1", fg="#2c3e50").pack(anchor="w", pady=(0, 5))
        
        # Alpha Report
        tk.Label(content, text="Alpha Report:", 
                font=("Arial", 9, "bold"), bg="#ecf0f1", fg="#34495e").pack(anchor="w", pady=(3, 2))
        alpha_reports = [
            "    • CAR Trade Loader - Processes trade data",
            "    • Option Security Loader - Processes option securities",
            "    • Future Security Loader - Processes future contracts"
        ]
        for report in alpha_reports:
            tk.Label(content, text=report, font=("Arial", 9), 
                    bg="#ecf0f1", fg="#34495e", justify="left").pack(anchor="w", pady=1)
        
        # Reconciliation Reports
        tk.Label(content, text="Asio Reconciliation Reports:", 
                font=("Arial", 9, "bold"), bg="#ecf0f1", fg="#34495e").pack(anchor="w", pady=(8, 2))
        recon_reports = [
            "    • DBS_QTY_RECON - DBS Quantity Reconciliation",
            "    • BNP_QTY_RECON - BNP Quantity Reconciliation",
            "    • Price_Recon - Price Reconciliation"
        ]
        for report in recon_reports:
            tk.Label(content, text=report, font=("Arial", 9), 
                    bg="#ecf0f1", fg="#34495e", justify="left").pack(anchor="w", pady=1)
    
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
