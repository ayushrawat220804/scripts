import tkinter as tk
from system_utilities import SystemUtilities

# Create root window
root = tk.Tk()
root.title("Windows System Utilities")

# Create app instance
app = SystemUtilities(root)

# Start main loop
print("Starting application...")
root.mainloop()
print("Application closed.") 