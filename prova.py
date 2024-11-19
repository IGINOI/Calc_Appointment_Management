import tkinter as tk
from tkinter import ttk

class DaysToggleApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Selezione giorni")
        self.root.geometry("300x300")

        # Label for instructions
        label = ttk.Label(root, text="Selezione i giorni degli appuntamenti:", font=("Arial", 14))
        label.pack(pady=10)

        # Container for toggle buttons
        self.toggle_frame = ttk.Frame(root)
        self.toggle_frame.pack(pady=10)

        # Days of the week
        self.days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        self.day_states = {day: tk.BooleanVar() for day in self.days}

        # Create toggle buttons
        for day in self.days:
            btn = ttk.Checkbutton(
                self.toggle_frame,
                text=day,
                variable=self.day_states[day],
                command=lambda day=day: self.toggle_day(day)
            )
            btn.pack(anchor="w", padx=10)

        # Selected days display
        self.selected_days_label = ttk.Label(root, text="Selected Days: None", font=("Arial", 12))
        self.selected_days_label.pack(pady=10)

    def toggle_day(self, day):
        """Update the selected days label when a day is toggled."""
        selected_days = [d for d, var in self.day_states.items() if var.get()]
        self.selected_days_label.config(
            text=f"Selected Days: {', '.join(selected_days) if selected_days else 'None'}"
        )

if __name__ == "__main__":
    root = tk.Tk()
    app = DaysToggleApp(root)
    root.mainloop()