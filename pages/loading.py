import tkinter as tk
import math

class LoadingSpinner(tk.Toplevel):
    def __init__(self, parent, text="Loading...", dot_count=8, radius=40, speed=50):
        super().__init__(parent)
        self.title("Please wait")
        self.geometry("220x220")
        self.resizable(False, False)
        self.configure(bg="white")
        self.transient(parent)   # stay on top of parent
        self.grab_set()          # modal

        # Center window relative to parent
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() // 2) - 110
        y = parent.winfo_y() + (parent.winfo_height() // 2) - 110
        self.geometry(f"+{x}+{y}")

        # Spinner canvas
        self.canvas = tk.Canvas(self, width=220, height=150, bg="white", highlightthickness=0)
        self.canvas.pack(pady=20)

        # Loading text
        self.label = tk.Label(self, text=text, font=("Arial", 12), bg="white")
        self.label.pack()

        # Spinner config
        self.dot_count = dot_count
        self.radius = radius
        self.speed = speed
        self.angle_offset = 0
        self.dots = []

        # Create dots
        for _ in range(dot_count):
            dot = self.canvas.create_oval(0, 0, 12, 12, fill="blue", outline="")
            self.dots.append(dot)

        self._running = True
        self.animate()

    def animate(self):
        if not self._running:
            return
        try:
            if not hasattr(self, "canvas") or not self.winfo_exists():
                return
            self.angle_offset += 25  # how much it rotates per frame
            for i, dot in enumerate(self.dots):
                angle = math.radians(self.angle_offset + (360 / self.dot_count) * i)
                x = 110 + self.radius * math.cos(angle)
                y = 75 + self.radius * math.sin(angle)
                self.canvas.coords(dot, x-6, y-6, x+6, y+6)

                # Optional fade effect
                brightness = int(150 + 100 * (1 + math.cos(angle)) / 2)
                color = f"#{brightness:02x}{brightness:02x}ff"
                self.canvas.itemconfig(dot, fill=color)
        except Exception:
            return
        self.after(self.speed, self.animate)

    def close(self):
        self._running = False
        try:
            self.destroy()
        except Exception:
            pass


# # Example usage
# if __name__ == "__main__":
#     root = tk.Tk()
#     root.geometry("400x300")

#     def show_loader():
#         loader = LoadingSpinner(root, text="Processing...")
#         root.after(5000, loader.close)  # auto-close after 5 sec

#     btn = tk.Button(root, text="Start Task", command=show_loader)
#     btn.pack(pady=50)

#     root.mainloop()
