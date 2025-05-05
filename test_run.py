import sys
import traceback

try:
    print("Attempting to import system_utilities...")
    import system_utilities
    print("Import successful")
    
    print("Checking if main execution code exists...")
    with open('system_utilities.py', 'r') as f:
        code = f.read()
        if "if __name__ == \"__main__\":" in code:
            print("Main execution code found")
        else:
            print("No main execution code found!")
    
    print("Attempting to create and run the application manually...")
    import tkinter as tk
    root = tk.Tk()
    app = system_utilities.SystemUtilities(root)
    print("Application started successfully!")
    root.mainloop()
    
except Exception as e:
    print(f"Error: {str(e)}")
    print("Traceback:")
    traceback.print_exc() 