import traceback
import sys

try:
    print("Importing Tkinter...")
    import tkinter as tk
    from tkinter import ttk
    print("Tkinter imported successfully!")
    
    print("Trying to create a root window...")
    root = tk.Tk()
    print("Root window created!")
    
    print("Importing SystemUtilities...")
    from system_utilities import SystemUtilities
    print("SystemUtilities imported successfully!")
    
    print("Creating application instance...")
    app = SystemUtilities(root)
    print("Application instance created!")
    
    # Add a callback to print when the window is fully initialized
    def ready():
        print("Application is fully initialized and ready!")
    
    root.after(1000, ready)
    
    # Add a button to close the app
    close_button = tk.Button(root, text="Close Application", command=root.destroy)
    close_button.pack(pady=10)
    
    print("Starting main loop...")
    root.mainloop()
    print("Application closed.")
    
except Exception as e:
    print(f"ERROR: {type(e).__name__}: {str(e)}")
    print("Traceback:")
    traceback.print_exc()
    sys.exit(1) 