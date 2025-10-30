import tkinter as tk

class SettingsPage(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg="#ecf0f1")
        lbl = tk.Label(self, text="⚙️ Settings", font=("Arial", 20), bg="#ecf0f1", fg="#2c3e50")
        lbl.pack(pady=20)
