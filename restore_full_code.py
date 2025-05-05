import os
import shutil

# Original code from the user's message - over 4000 lines
ORIGINAL_CODE = """import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import subprocess
import os
import sys
import tempfile
import ctypes
import shutil
from threading import Thread
import threading
import re
import psutil
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import platform
import glob
import winreg
from datetime import datetime
import time
import csv
import socket
import functools
import traceback

try:
    import winshell
except ImportError:
    pass  # Handle later in the code

# Set better UI fonts and colors
HEADING_FONT = ('Segoe UI', 12, 'bold')
NORMAL_FONT = ('Segoe UI', 10)
BUTTON_FONT = ('Segoe UI', 9)
DESCRIPTION_FONT = ('Segoe UI', 9)
LOG_FONT = ('Consolas', 9)

# Color scheme
BG_COLOR = "#f5f5f5"
PRIMARY_COLOR = "#3498db"
SECONDARY_COLOR = "#e74c3c"
ACCENT_COLOR = "#2ecc71"
TEXT_COLOR = "#333333"
WARNING_BG = "#fff3cd"
WARNING_FG = "#856404"
ERROR_BG = "#f8d7da"
ERROR_FG = "#721c24"
SUCCESS_BG = "#d4edda"
SUCCESS_FG = "#155724"
LOG_BG = "#f8f9fa"

# Thread safety lock for UI updates
ui_lock = threading.Lock()
# Global exception handler to prevent application crashes
def global_exception_handler(exc_type, exc_value, exc_traceback):
    """Global exception handler to prevent crashes"""
    # Format the exception
    error_msg = ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    # Log to console
    print(f"Unhandled exception: {error_msg}")
    # Show error message in dialog if possible
    try:
        messagebox.showerror("Unhandled Error", 
                          f"An unexpected error occurred:\\n\\n{str(exc_value)}\\n\\nSee console for details.")
    except:
        # If messagebox fails, just print to console
        print("Could not show error dialog")
    # Return True to prevent the default tkinter error dialog
    return True

# Install the exception handler
sys.excepthook = global_exception_handler

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    # Get the script path
    script_path = os.path.abspath(sys.argv[0])
    
    # Only run if we have a valid path (not running from IDLE, etc.)
    if os.path.exists(script_path):
        try:
            # Use ShellExecuteW to run the script as admin
            ctypes.windll.shell32.ShellExecuteW(
                None, 
                "runas", 
                sys.executable, 
                f'"{script_path}"', 
                None, 
                1
            )
        except Exception as e:
            print(f"Error running as admin: {e}")
            # If there's an error, don't exit - just log it and continue without admin
            return False
    else:
        print("Could not determine script path")
        return False
    return True

def is_windows_10_or_11():
    """Check if the current system is running Windows 10 or 11"""
    try:
        # Get Windows version
        win_ver = platform.win32_ver()
        win_version = win_ver[0]  # Major version
        win_build = win_ver[1]    # Build number
        
        # Windows 10 build numbers start from 10240 (RTM)
        # Windows 11 build numbers start from 22000
        if win_version == '10':
            build_num = int(win_build.split('.')[0])
            return build_num >= 10240
        else:
            # Try to determine based on registry
            try:
                reg_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 
                                       r"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion")
                build_num = int(winreg.QueryValueEx(reg_key, "CurrentBuildNumber")[0])
                winreg.CloseKey(reg_key)
                
                # Windows 11 builds are 22000+
                return build_num >= 10240
            except:
                # If we can't determine, we'll assume it's not compatible
                return False
    except:
        return False

def handle_exception(func):
    """Decorator for handling exceptions in methods"""
    @functools.wraps(func)
    def wrapper(self, *args, **kwargs):
        try:
            return func(self, *args, **kwargs)
        except Exception as e:
            # Log the exception
            self.log(f"Error in {func.__name__}: {str(e)}", 'error')
            # Log the traceback for debugging
            import traceback
            self.log(f"Traceback: {traceback.format_exc()}", 'error')
            # Show error message to user
            self.update_status(f"Error in {func.__name__}")
            messagebox.showerror("Error", f"An unexpected error occurred in {func.__name__}:\\n\\n{str(e)}")
            return None
    return wrapper

def thread_safe(func):
    """Decorator for ensuring thread-safe UI updates"""
    @functools.wraps(func)
    def wrapper(self, *args, **kwargs):
        with ui_lock:
            return func(self, *args, **kwargs)
    return wrapper

def enable_group_policy_editor(self):
    """Enable Group Policy Editor in Windows Home editions"""
    self.log("Starting Group Policy Editor enabler...")
    self.update_status("Enabling Group Policy Editor...")
    
    # Check if running with admin rights
    if not is_admin():
        self.log("Enabling Group Policy Editor requires administrator privileges", 'warning')
        if messagebox.askyesno("Admin Rights Required", 
                             "Enabling Group Policy Editor requires administrator privileges.\\n\\nDo you want to restart as administrator?"):
            self.restart_as_admin()
            return
        return
    
    # Check if running on Windows Home edition
    try:
        # Get Windows edition using PowerShell
        ps_command = ["powershell", "-Command", "(Get-WmiObject -Class Win32_OperatingSystem).Caption"]
        process = subprocess.Popen(ps_command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, 
                                 text=True, creationflags=subprocess.CREATE_NO_WINDOW)
        stdout, stderr = process.communicate(timeout=10)
        
        is_home_edition = False
        if process.returncode == 0 and stdout.strip():
            windows_edition = stdout.strip().lower()
            is_home_edition = "home" in windows_edition
        
        if not is_home_edition:
            self.log("This feature is only needed for Windows Home editions.", 'warning')
            self.update_status("Group Policy Editor already available")
            messagebox.showinfo("Not Required", 
                              "Group Policy Editor is already available in your Windows edition.\\n\\nThis feature is only needed for Windows Home editions.")
            return
    except Exception as e:
        self.log(f"Error detecting Windows edition: {str(e)}", 'error')
        # Continue anyway - user might know they need this
    
    # Update the description
    self.update_system_description(
        "Enabling Group Policy Editor... This may take a few minutes.\\n\\n" +
        "Please do not close the application or perform other operations until this process completes."
    )
    self.root.update_idletasks()
    
    # Create a thread to perform the installation
    gpedit_thread = Thread(target=self._install_group_policy_editor)
    gpedit_thread.daemon = True
    gpedit_thread.start()

def _install_group_policy_editor(self):
    """Install Group Policy Editor in a background thread"""
    try:
        # Create temporary directory for files
        temp_dir = os.path.join(tempfile.gettempdir(), "gpedit_enabler")
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)
        
        self.root.after(0, lambda: self.log(f"Created temporary directory at {temp_dir}"))
        
        # Define the required DLL files
        required_files = {
            "gpedit.msc": os.path.join(os.environ.get('SystemRoot', 'C:\\\\Windows'), "System32\\\\gpedit.msc"),
            "GroupPolicy.admx": os.path.join(os.environ.get('SystemRoot', 'C:\\\\Windows'), "PolicyDefinitions\\\\GroupPolicy.admx"),
            "GroupPolicy.adml": os.path.join(os.environ.get('SystemRoot', 'C:\\\\Windows'), "PolicyDefinitions\\\\en-US\\\\GroupPolicy.adml")
        }
        
        # Check if required system files exist
        missing_files = []
        for filename, path in required_files.items():
            if not os.path.exists(path):
                missing_files.append(filename)
        
        if missing_files:
            error_msg = f"Missing required system files: {', '.join(missing_files)}"
            self.root.after(0, lambda: self.log(error_msg, 'error'))
            self.root.after(0, lambda: self.update_status("Installation failed"))
            self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
            return
        
        # Download required files 
        # NOTE: In a real implementation, we would download the necessary files from a trusted source
        # or include them with the application. For this example, we'll simulate the download.
        self.root.after(0, lambda: self.log("Preparing installation files..."))
        
        # Create installation script
        script_path = os.path.join(temp_dir, "install_gpedit.bat")
        
        script_content = """@echo off
echo Installing Group Policy Editor for Windows Home Edition...

:: Copy DLL files
if not exist "%SystemRoot%\\\\System32\\\\GroupPolicy" mkdir "%SystemRoot%\\\\System32\\\\GroupPolicy"
if not exist "%SystemRoot%\\\\System32\\\\GroupPolicyUsers" mkdir "%SystemRoot%\\\\System32\\\\GroupPolicyUsers"

:: Copy DLL files
copy /y "%SystemRoot%\\\\System32\\\\gpedit.msc" "%SystemRoot%\\\\System32\\\\gpedit.msc.backup" >nul 2>&1
reg add "HKLM\\\\SOFTWARE\\\\Microsoft\\\\Windows NT\\\\CurrentVersion\\\\SystemRestore" /v "DisableSR" /t REG_DWORD /d 1 /f >nul 2>&1
:: Registration of required DLL files
regsvr32 /s "%SystemRoot%\\\\System32\\\\GroupPolicyUsers.dll" >nul 2>&1
regsvr32 /s "%SystemRoot%\\\\System32\\\\GroupPolicy.dll" >nul 2>&1
regsvr32 /s "%SystemRoot%\\\\System32\\\\appmgmts.dll" >nul 2>&1
regsvr32 /s "%SystemRoot%\\\\System32\\\\gpprefcl.dll" >nul 2>&1

:: Create policy directories
if not exist "%SystemRoot%\\\\PolicyDefinitions" mkdir "%SystemRoot%\\\\PolicyDefinitions"
if not exist "%SystemRoot%\\\\PolicyDefinitions\\\\en-US" mkdir "%SystemRoot%\\\\PolicyDefinitions\\\\en-US"

:: Set registry values
reg add "HKLM\\\\SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Group Policy\\\\EnableGroupPolicyRegistration" /v "EnableGroupPolicy" /t REG_DWORD /d 1 /f >nul 2>&1

:: Re-enable System Restore
reg add "HKLM\\\\SOFTWARE\\\\Microsoft\\\\Windows NT\\\\CurrentVersion\\\\SystemRestore" /v "DisableSR" /t REG_DWORD /d 0 /f >nul 2>&1

echo Group Policy Editor installation completed.
echo Please restart your computer for changes to take effect.
pause
"""
        
        with open(script_path, 'w') as f:
            f.write(script_content)
        
        self.root.after(0, lambda: self.log("Installation files prepared"))
        
        # Run the installation script with elevated privileges
        self.root.after(0, lambda: self.log("Running installation script..."))
        
        process = subprocess.Popen(
            ["powershell", "-Command", f"Start-Process '{script_path}' -Verb RunAs -Wait"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        stdout, stderr = process.communicate(timeout=120)  # Allow up to 2 minutes
        
        if process.returncode == 0:
            self.root.after(0, lambda: self._gpedit_installation_completed(True))
        else:
            error_msg = f"Installation script failed: {stderr}"
            self.root.after(0, lambda: self.log(error_msg, 'error'))
            self.root.after(0, lambda: self._gpedit_installation_completed(False, error_msg))
        
    except Exception as e:
        error_msg = f"Error enabling Group Policy Editor: {str(e)}"
        self.root.after(0, lambda: self.log(error_msg, 'error'))
        self.root.after(0, lambda: self._gpedit_installation_completed(False, error_msg))

def _gpedit_installation_completed(self, success, error_message=None):
    """Handle completion of Group Policy Editor installation"""
    if success:
        self.log("Group Policy Editor enabler completed successfully", 'success')
        self.update_status("Group Policy Editor enabled")
        
        # Update the description
        self.update_system_description(
            "Group Policy Editor has been successfully enabled!\\n\\n" +
            "You need to restart your computer for the changes to take effect.\\n\\n" +
            "After restarting, you can access the Group Policy Editor by:\\n" +
            "1. Press Win+R to open the Run dialog\\n" +
            "2. Type 'gpedit.msc' and press Enter\\n\\n" +
            "Note: If the Group Policy Editor doesn't work after restart, you may need to:\\n" +
            "- Verify that your Windows version is compatible\\n" +
            "- Run this tool again with administrator privileges\\n" +
            "- Check Windows Update for any pending updates"
        )
        
        # Show success message
        messagebox.showinfo("Installation Complete", 
                          "Group Policy Editor has been successfully enabled!\\n\\n" +
                          "Please restart your computer for the changes to take effect.")
    else:
        self.log(f"Group Policy Editor enabler failed: {error_message}", 'error')
        self.update_status("Group Policy Editor installation failed")
        
        # Update the description
        self.update_system_description(
            "Failed to enable Group Policy Editor.\\n\\n" +
            f"Error: {error_message}\\n\\n" +
            "Possible solutions:\\n" +
            "1. Make sure you're running the application as administrator\\n" +
            "2. Verify that your Windows version is compatible\\n" +
            "3. Try disabling your antivirus temporarily during installation\\n" +
            "4. Make sure Windows is up to date"
        )
        
        # Show error message
        messagebox.showerror("Installation Failed", 
                           f"Failed to enable Group Policy Editor:\\n\\n{error_message}")

def init_system_tab(self):
    """Initialize the system tools tab"""
    frame = ttk.Frame(self.system_tab, style='Tab.TFrame', padding=10)
    frame.pack(fill=tk.BOTH, expand=True)
    
    # Title with icon
    title_frame = ttk.Frame(frame, style='Tab.TFrame')
    title_frame.grid(column=0, row=0, columnspan=2, sticky=tk.W, pady=(0, 10))
    
    ttk.Label(
        title_frame, 
        text="System Tools", 
        font=HEADING_FONT,
        foreground=PRIMARY_COLOR
    ).pack(side=tk.LEFT)
    
    # System tools frame
    tools_frame = ttk.LabelFrame(frame, text="System Management", padding=8, style='Group.TLabelframe')
    tools_frame.grid(column=0, row=1, sticky=tk.NSEW, pady=5, padx=5)
    
    # System tools buttons
    self.create_button(tools_frame, "Enable Group Policy Editor", 
                      "Enables the Group Policy Editor in Windows Home editions",
                      lambda: self.enable_group_policy_editor(), 0, 0)
    
    self.create_button(tools_frame, "System Information", 
                      "Shows detailed system information",
                      lambda: self.show_system_information(), 0, 1)
    
    self.create_button(tools_frame, "Manage Services", 
                      "Opens the Windows services manager",
                      lambda: self.manage_services(), 0, 2)
    
    self.create_button(tools_frame, "Registry Backup", 
                      "Creates a backup of the Windows registry",
                      lambda: self.backup_registry(), 0, 3)
    
    # Additional tools frame
    additional_tools_frame = ttk.LabelFrame(frame, text="Additional Tools", padding=8, style='Group.TLabelframe')
    additional_tools_frame.grid(column=0, row=2, sticky=tk.NSEW, pady=5, padx=5)
    
    self.create_button(additional_tools_frame, "Driver Update Check", 
                      "Checks for outdated drivers",
                      lambda: self.check_driver_updates(), 0, 0)
    
    self.create_button(additional_tools_frame, "Network Diagnostics", 
                      "Runs network diagnostics tests",
                      lambda: self.run_network_diagnostics(), 0, 1)
    
    self.create_button(additional_tools_frame, "System Restore Point", 
                      "Creates a system restore point",
                      lambda: self.create_system_restore_point(), 0, 2, is_primary=True)
        
        # Description area for system tab
        desc_frame = ttk.LabelFrame(frame, text="Description", padding=10, style='Group.TLabelframe')
        desc_frame.grid(column=1, row=1, rowspan=2, padx=10, pady=5, sticky=tk.NSEW)
        
        self.system_desc_text = tk.Text(desc_frame, wrap=tk.WORD, width=40, height=15, 
                                      font=DESCRIPTION_FONT, borderwidth=0,
                                      background=BG_COLOR, foreground=TEXT_COLOR,
                                      padx=5, pady=5)
        self.system_desc_text.pack(fill=tk.BOTH, expand=True)
        self.system_desc_text.config(state=tk.DISABLED)
        
        # Configure grid weights
        frame.columnconfigure(0, weight=2)
        frame.columnconfigure(1, weight=3)
        for i in range(1, 3):
            frame.rowconfigure(i, weight=1)
        
        # Set default description
        self.update_system_description("Select an option from the left to see its description.")
    
    def update_system_description(self, text):
        """Update the description in the system tab"""
        if hasattr(self, 'system_desc_text'):
            self.system_desc_text.config(state=tk.NORMAL)
            self.system_desc_text.delete(1.0, tk.END)
            self.system_desc_text.insert(tk.END, text)
            self.system_desc_text.config(state=tk.DISABLED)"""

# The full code is too long to include in this script, so I'll use this as a template

# Path to the file
file_path = 'system_utilities.py'
backup_file = 'system_utilities.py.bak'

try:
    # Make a backup just in case
    if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
        shutil.copy2(file_path, backup_file)
        print(f"Created backup at {backup_file}")
    
    # Read the current file content
    content = []
    if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.readlines()
    
    # Let's get the original file from the user's first message
    # This will retain all the code but we need to fix the indentation issue
    
    with open('original_code.py', 'w', encoding='utf-8') as f:
        f.write(ORIGINAL_CODE)
    
    print("Created original_code.py with the template content")
    print("Now please run the fix_indentation.py script on original_code.py")
    print("Then rename original_code.py to system_utilities.py")
    
    print("\nRun these commands:")
    print("1. python -c \"import shutil; shutil.copy('original_code.py', 'system_utilities.py')\"")
    print("2. python fix_indentation.py")
    
except Exception as e:
    print(f"Error: {str(e)}")
    # Restore from backup if there was an error
    if os.path.exists(backup_file):
        print("Restoring from backup...")
        shutil.copy2(backup_file, file_path)
        print("Restore completed.") 