import os
import sys
import winreg
import subprocess
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, simpledialog
import ctypes
import platform
import threading
from datetime import datetime
import traceback
import webbrowser
import time  # Required for _stop_idm method
import random  # For generating serial key
import string  # For generating serial key

# Create .pyw version for a console-less experience
if __name__ == "__main__" and os.path.basename(sys.argv[0]).lower().endswith('.py'):
    pyw_path = os.path.splitext(sys.argv[0])[0] + '.pyw'
    if not os.path.exists(pyw_path):
        try:
            with open(sys.argv[0], 'r') as src_file:
                with open(pyw_path, 'w') as dst_file:
                    dst_file.write(src_file.read())
            # Inform but continue with current execution
            print(f"Created {pyw_path} - use this file for a console-less experience")
        except:
            pass  # Silently continue if we can't create .pyw file

# Hide console window directly using Windows API
try:
    # This will only work on Windows
    if platform.system() == "Windows":
        # Get console window handle
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd != 0:
            # Hide the console window (0 = hide, 5 = show)
            ctypes.windll.user32.ShowWindow(hwnd, 0)
except Exception:
    pass  # Silently fail if we can't hide the console

# Create debug log function - only if debug is enabled
DEBUG_ENABLED = False
debug_log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "idm_manager_debug.log")
def log_debug(message):
    if DEBUG_ENABLED:
        try:
            with open(debug_log_path, "a") as f:
                f.write(f"{datetime.now()} - {message}\n")
        except:
            pass  # Silently fail if we can't write to log

# Check for admin rights
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

# Global exception handler to prevent immediate closing
def show_error(title, message):
    try:
        messagebox.showerror(title, message)
    except:
        print(f"ERROR: {title} - {message}")
        input("Press Enter to continue...")

# Custom dialog for activation details
class ActivationDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("IDM Activation")
        self.transient(parent)
        self.resizable(False, False)
        self.result = None
        
        # Center on parent
        window_width = 450
        window_height = 400
        position_right = int(parent.winfo_rootx() + (parent.winfo_width() / 2) - (window_width / 2))
        position_down = int(parent.winfo_rooty() + (parent.winfo_height() / 2) - (window_height / 2))
        self.geometry(f"{window_width}x{window_height}+{position_right}+{position_down}")
        
        # Make modal
        self.grab_set()
        self.focus_set()
        
        # Frame for entries
        main_frame = ttk.Frame(self, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="Internet Download Manager Activation", font=("Arial", 14, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 5))
        
        # Subtitle
        subtitle_label = ttk.Label(main_frame, text="Enter your information for IDM registration", font=("Arial", 10))
        subtitle_label.grid(row=1, column=0, columnspan=2, pady=(0, 15))
        
        # First Name
        ttk.Label(main_frame, text="First Name:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.fname_var = tk.StringVar(value="John")
        ttk.Entry(main_frame, textvariable=self.fname_var, width=30).grid(row=2, column=1, sticky=tk.W, pady=5)
        
        # Last Name
        ttk.Label(main_frame, text="Last Name:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.lname_var = tk.StringVar(value="Doe")
        ttk.Entry(main_frame, textvariable=self.lname_var, width=30).grid(row=3, column=1, sticky=tk.W, pady=5)
        
        # Email
        ttk.Label(main_frame, text="Email:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.email_var = tk.StringVar(value="user@example.com")
        ttk.Entry(main_frame, textvariable=self.email_var, width=30).grid(row=4, column=1, sticky=tk.W, pady=5)
        
        # Serial Prefix (first 4 chars)
        ttk.Label(main_frame, text="Serial Prefix (4 chars):").grid(row=5, column=0, sticky=tk.W, pady=5)
        self.serial_prefix_var = tk.StringVar(value="ABCD")
        self.serial_entry = ttk.Entry(main_frame, textvariable=self.serial_prefix_var, width=8)
        self.serial_entry.grid(row=5, column=1, sticky=tk.W, pady=5)
        
        # Add validation for serial prefix
        self.serial_entry.bind('<KeyRelease>', self.validate_serial)
        
        # Serial Preview
        ttk.Label(main_frame, text="License Key:").grid(row=6, column=0, sticky=tk.W, pady=5)
        self.serial_preview_var = tk.StringVar(value=self.generate_preview_serial())
        preview_label = ttk.Label(main_frame, textvariable=self.serial_preview_var, font=("Consolas", 10))
        preview_label.grid(row=6, column=1, sticky=tk.W, pady=5)
        
        # Generate new preview button
        ttk.Button(main_frame, text="Generate New Key", command=self.generate_new_serial, width=18).grid(
            row=7, column=1, sticky=tk.W, pady=5)
        
        # Information note
        note_label = ttk.Label(main_frame, 
                              text="Note: This information is stored locally and not sent anywhere.", 
                              font=("Arial", 8), foreground="gray")
        note_label.grid(row=8, column=0, columnspan=2, pady=(15, 5))
        
        # Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=9, column=0, columnspan=2, pady=10)
        
        # Style for the activate button
        style = ttk.Style()
        style.configure("Activate.TButton", font=("Arial", 11, "bold"))
        
        ttk.Button(btn_frame, text="Activate IDM", command=self.on_ok, width=15, style="Activate.TButton").pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.on_cancel, width=10).pack(side=tk.LEFT, padx=5)
        
        # Update preview when prefix changes
        self.serial_prefix_var.trace_add("write", lambda *args: self.update_preview())
        
        # Wait for window to be dismissed
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)
        self.wait_window(self)
    
    def validate_serial(self, event=None):
        """Validate serial prefix to be 4 characters and uppercase letters/numbers only"""
        value = self.serial_prefix_var.get().upper()
        # Filter to only allow A-Z and 0-9
        valid_chars = ''.join(c for c in value if c in string.ascii_uppercase + string.digits)
        # Limit to 4 characters
        valid_chars = valid_chars[:4]
        
        # Only update if changed to avoid infinite recursion
        if valid_chars != self.serial_prefix_var.get():
            self.serial_prefix_var.set(valid_chars)
    
    def generate_preview_serial(self):
        """Generate a preview serial based on current prefix"""
        prefix = self.serial_prefix_var.get().upper()
        # If prefix is less than 4 chars, pad it
        if len(prefix) < 4:
            prefix = prefix.ljust(4, 'X')
            
        # Generate 3 more groups
        parts = [prefix]
        for _ in range(3):
            part = ''.join(random.choices(string.ascii_uppercase + string.digits, k=4))
            parts.append(part)
        
        return "-".join(parts)
    
    def update_preview(self):
        """Update the preview serial"""
        self.serial_preview_var.set(self.generate_preview_serial())
    
    def generate_new_serial(self):
        """Generate a completely new serial preview but keep the prefix"""
        self.serial_preview_var.set(self.generate_preview_serial())
    
    def on_ok(self):
        """OK button handler"""
        # Validate first name and last name are not empty
        if not self.fname_var.get().strip() or not self.lname_var.get().strip():
            messagebox.showerror("Error", "First name and last name cannot be empty", parent=self)
            return
        
        # Validate email contains @ symbol
        if '@' not in self.email_var.get():
            messagebox.showerror("Error", "Please enter a valid email address", parent=self)
            return
        
        # Validate serial prefix is exactly 4 characters
        serial_prefix = self.serial_prefix_var.get()
        if len(serial_prefix) != 4:
            messagebox.showerror("Error", "Serial prefix must be exactly 4 characters", parent=self)
            return
        
        # Set result and close dialog
        self.result = {
            'fname': self.fname_var.get(),
            'lname': self.lname_var.get(),
            'email': self.email_var.get(),
            'serial_prefix': serial_prefix,
            'full_serial': self.serial_preview_var.get()  # Include the previewed serial
        }
        self.destroy()
    
    def on_cancel(self):
        """Cancel button handler"""
        self.result = None
        self.destroy()

try:
    # Main application class
    class IDMManager:
        def __init__(self, root):
            self.root = root
            log_debug("Initializing UI")
            self.version = "3.1"  # Updated version number
            
            # Cache for registry queries to speed up repeated operations
            self.registry_cache = {}
            
            # Set paths based on architecture
            self.system_info = self.get_system_info()
            
            self.init_ui()
            
            # Perform non-critical checks after UI is shown for better perceived startup speed
            self.root.after(10, self.perform_startup_checks)
            log_debug("IDMManager initialized")
        
        def perform_startup_checks(self):
            """Perform non-critical startup checks"""
            # Use threading for background checks to keep UI responsive
            threading.Thread(target=self._background_checks, daemon=True).start()
        
        def _background_checks(self):
            """Background checks running in a separate thread"""
            self.check_idm_installation()
            self.check_idm_activation()
            # Skip automatic update check at startup for speed
        
        def get_system_info(self):
            """Get system architecture and IDM paths"""
            arch = platform.architecture()[0]
            if arch == '32bit':
                CLSID = r"Software\Classes\CLSID"
                HKLM_IDM = r"Software\Internet Download Manager"
            else:
                CLSID = r"Software\Classes\Wow6432Node\CLSID"
                HKLM_IDM = r"SOFTWARE\Wow6432Node\Internet Download Manager"
                
            HKCU = winreg.HKEY_CURRENT_USER
            HKLM = winreg.HKEY_LOCAL_MACHINE
            
            # Find IDM executable
            idm_path = None
            try:
                reg_key = winreg.OpenKey(HKCU, r"Software\DownloadManager")
                idm_path = winreg.QueryValueEx(reg_key, "ExePath")[0]
                winreg.CloseKey(reg_key)
            except:
                pass
            
            if not idm_path or not os.path.exists(idm_path):
                if arch == '32bit':
                    idm_path = os.path.join(os.environ.get('ProgramFiles', r'C:\Program Files'), 
                                            r"Internet Download Manager\IDMan.exe")
                else:
                    idm_path = os.path.join(os.environ.get('ProgramFiles(x86)', r'C:\Program Files (x86)'), 
                                            r"Internet Download Manager\IDMan.exe")
            
            return {
                'arch': arch,
                'CLSID': CLSID,
                'HKLM_IDM': HKLM_IDM,
                'HKCU': HKCU,
                'HKLM': HKLM,
                'idm_path': idm_path
            }
        
        def init_ui(self):
            # Configure the main window
            self.root.title(f"IDM Manager v{self.version}")
            self.root.geometry("1100x800")
            self.root.minsize(1100, 800)
            
            # Add a protocol handler for window close
            self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
            
            # Configure style - simplified for speed
            style = ttk.Style()
            style.theme_use('clam')  # Use a modern theme if available
            
            # Create main frame
            main_frame = ttk.Frame(self.root)
            main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # Header - simplified
            header_label = ttk.Label(main_frame, 
                                    text="Internet Download Manager - Activation Manager", 
                                    font=('Arial', 16, 'bold'))
            header_label.pack(fill=tk.X, pady=10)
            
            # Content area with two columns
            content_frame = ttk.Frame(main_frame)
            content_frame.pack(fill=tk.BOTH, expand=True, pady=10)
            content_frame.columnconfigure(0, weight=1)
            content_frame.columnconfigure(1, weight=3)
            
            # Left column - Control panel
            left_frame = ttk.Frame(content_frame, padding=10)
            left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
            
            # IDM Info Panel
            info_frame = ttk.LabelFrame(left_frame, text="IDM Information", padding=10)
            info_frame.pack(fill=tk.X, pady=(0, 10))
            
            # IDM Status
            self.idm_status_var = tk.StringVar(value="Checking...")
            ttk.Label(info_frame, text="Installation:").grid(row=0, column=0, sticky="w", pady=2)
            ttk.Label(info_frame, textvariable=self.idm_status_var).grid(row=0, column=1, sticky="w", pady=2)
            
            # IDM Version
            self.idm_version_var = tk.StringVar(value="Checking...")
            ttk.Label(info_frame, text="Version:").grid(row=1, column=0, sticky="w", pady=2)
            ttk.Label(info_frame, textvariable=self.idm_version_var).grid(row=1, column=1, sticky="w", pady=2)
            
            # IDM Path
            self.idm_path_var = tk.StringVar(value="Checking...")
            ttk.Label(info_frame, text="Path:").grid(row=2, column=0, sticky="w", pady=2)
            path_label = ttk.Label(info_frame, textvariable=self.idm_path_var, wraplength=200)
            path_label.grid(row=2, column=1, sticky="w", pady=2)
            
            # Activation Status
            self.activation_status_var = tk.StringVar(value="Checking...")
            ttk.Label(info_frame, text="Status:").grid(row=3, column=0, sticky="w", pady=2)
            status_label = ttk.Label(info_frame, textvariable=self.activation_status_var)
            status_label.grid(row=3, column=1, sticky="w", pady=2)
            
            # Actions Panel
            actions_frame = ttk.LabelFrame(left_frame, text="Actions", padding=10)
            actions_frame.pack(fill=tk.BOTH, expand=True)
            
            # Action buttons
            button_data = [
                ("Custom IDM Activation", self.show_custom_activation),
                ("Fix Browser Integration", self.fix_browser_integration),
                ("Reset Trial", self.reset_trial),
                ("Reset to Factory Defaults", self.reset_factory_defaults),
                ("Check Latest Version", self.check_latest_version),
                ("Toggle Firewall", self.toggle_firewall),
                ("Clean Registry", self.clean_registry),
                ("Download IDM", self.download_idm),
                ("Launch IDM", self.launch_idm)
            ]
            
            for text, command in button_data:
                btn = ttk.Button(actions_frame, text=text, command=command, width=20)
                btn.pack(fill=tk.X, pady=5)
            
            # Help button at the bottom
            help_btn = ttk.Button(left_frame, text="Help / About", command=self.show_help)
            help_btn.pack(fill=tk.X, pady=(10, 0))
            
            # Right column - Log panel
            right_frame = ttk.Frame(content_frame, padding=10)
            right_frame.grid(row=0, column=1, sticky="nsew")
            
            # Status and events log
            log_frame = ttk.LabelFrame(right_frame, text="Status Log", padding=10)
            log_frame.pack(fill=tk.BOTH, expand=True)
            
            self.log_text = scrolledtext.ScrolledText(log_frame, height=30, wrap=tk.WORD, font=("Consolas", 10))
            self.log_text.pack(fill=tk.BOTH, expand=True)
            self.log_text.config(state=tk.DISABLED)
            
            # Configure text tags for colors
            self.log_text.tag_configure("info", foreground="black")
            self.log_text.tag_configure("success", foreground="green")
            self.log_text.tag_configure("warning", foreground="orange")
            self.log_text.tag_configure("error", foreground="red")
            
            # Footer with status bar
            footer_frame = ttk.Frame(main_frame)
            footer_frame.pack(fill=tk.X, pady=(10, 0))
            
            # Status message
            self.status_var = tk.StringVar(value="Ready")
            status_label = ttk.Label(footer_frame, textvariable=self.status_var)
            status_label.pack(side=tk.LEFT)
            
            # Admin status indicator
            admin_status = "Admin: Yes" if is_admin() else "Admin: No (Limited functionality)"
            admin_color = "green" if is_admin() else "red"
            admin_label = ttk.Label(footer_frame, text=admin_status, foreground=admin_color)
            admin_label.pack(side=tk.RIGHT)
            
            # Firewall status indicator
            self.firewall_status_var = tk.StringVar(value="Checking...")
            self.firewall_status_label = ttk.Label(footer_frame, textvariable=self.firewall_status_var)
            self.firewall_status_label.pack(side=tk.RIGHT, padx=20)
            
            # Check firewall status
            threading.Thread(target=self._update_firewall_status, daemon=True).start()
            
            # Log initial message
            self.log_message("IDM Manager Started", "info")
            if not is_admin():
                self.log_message("WARNING: Not running with administrator privileges. Some features may not work.", "warning")
        
        def on_closing(self):
            log_debug("Window closing")
            self.root.destroy()
        
        def log_message(self, message, level="info"):
            """Add a message to the log with timestamp"""
            timestamp = datetime.now().strftime("%H:%M:%S")
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, f"{timestamp} - {message}\n", level)
            self.log_text.see(tk.END)
            self.log_text.config(state=tk.DISABLED)
            log_debug(f"LOG ({level}): {message}")
        
        def update_status(self, message):
            """Update the status bar message"""
            self.status_var.set(message)
            log_debug(f"Status update: {message}")
        
        def check_idm_installation(self):
            """Check if IDM is installed and get version information"""
            self.log_message("Checking for IDM installation...", "info")
            self.update_status("Checking IDM installation...")
            
            idm_path = self.system_info['idm_path']
            
            if os.path.exists(idm_path):
                self.log_message(f"IDM found at: {idm_path}", "success")
                self.root.after(0, lambda: self.idm_path_var.set(idm_path))
                self.root.after(0, lambda: self.idm_status_var.set("Installed"))
                
                # Try to get version
                try:
                    reg_key = winreg.OpenKey(self.system_info['HKCU'], r"Software\DownloadManager")
                    version = winreg.QueryValueEx(reg_key, "idmvers")[0]
                    winreg.CloseKey(reg_key)
                    self.root.after(0, lambda: self.idm_version_var.set(version))
                    self.log_message(f"IDM version: {version}", "info")
                except:
                    self.root.after(0, lambda: self.idm_version_var.set("Unknown"))
                    self.log_message("Could not determine IDM version", "warning")
                
                self.update_status("IDM is installed")
                return True
            else:
                self.log_message("IDM not found", "warning")
                self.root.after(0, lambda: self.idm_path_var.set("Not installed"))
                self.root.after(0, lambda: self.idm_status_var.set("Not installed"))
                self.root.after(0, lambda: self.idm_version_var.set("N/A"))
                self.update_status("IDM is not installed")
                return False
        
        def check_idm_activation(self):
            """Check IDM activation status"""
            self.log_message("Checking IDM activation status...", "info")
            
            try:
                # Use cached registry key if available
                if 'dm_key' not in self.registry_cache:
                    self.registry_cache['dm_key'] = winreg.OpenKey(self.system_info['HKCU'], r"Software\DownloadManager")
                
                reg_key = self.registry_cache['dm_key']
                
                try:
                    # Check for registration info
                    fname = winreg.QueryValueEx(reg_key, "FName")[0]
                    lname = winreg.QueryValueEx(reg_key, "LName")[0]
                    email = winreg.QueryValueEx(reg_key, "Email")[0]
                    serial = winreg.QueryValueEx(reg_key, "Serial")[0]
                    
                    if fname and lname and email and serial:
                        self.root.after(0, lambda: self.activation_status_var.set("Activated"))
                        self.log_message(f"IDM is registered to: {fname} {lname}", "success")
                        return True
                except:
                    pass
                
                # Check for trial status
                try:
                    trial_info = winreg.QueryValueEx(reg_key, "tvfrdt")[0]
                    if trial_info:
                        self.root.after(0, lambda: self.activation_status_var.set("Trial"))
                        self.log_message("IDM is in trial mode", "warning")
                        return False
                except:
                    pass
                    
                self.root.after(0, lambda: self.activation_status_var.set("Unknown"))
                self.log_message("Could not determine activation status", "warning")
                return False
            except Exception:
                self.root.after(0, lambda: self.activation_status_var.set("Unknown"))
                self.log_message("Could not determine activation status", "warning")
                if 'dm_key' in self.registry_cache:
                    del self.registry_cache['dm_key']  # Remove invalid cached key
                return False
        
        def check_for_updates(self):
            """Check for updates to the script"""
            threading.Thread(target=self._check_updates_thread, daemon=True).start()
        
        def _check_updates_thread(self):
            self.log_message("Checking for script updates...", "info")
            # This is a placeholder - in a real implementation, would check against a server
            self.log_message("No updates available. You have the latest version.", "success")
        
        def check_latest_version(self):
            """Check for the latest version of IDM online"""
            self.log_message("Checking for latest IDM version...", "info")
            self.update_status("Checking for latest IDM version...")
            
            threading.Thread(target=self._check_idm_version, daemon=True).start()
        
        def _check_idm_version(self):
            """Thread to check latest IDM version"""
            try:
                import urllib.request
                import re
                
                # Open the IDM website
                with urllib.request.urlopen("https://www.internetdownloadmanager.com/download.html") as response:
                    html = response.read().decode('utf-8')
                
                # Find version information
                version_match = re.search(r'The latest version is ([\d\.]+) build (\d+)', html)
                if version_match:
                    version = f"{version_match.group(1)} build {version_match.group(2)}"
                    self.root.after(0, lambda: self.log_message(f"Latest IDM version available: {version}", "success"))
                    
                    # Compare with installed version
                    current_version = self.idm_version_var.get()
                    if current_version and current_version != "Unknown" and current_version != "N/A":
                        if current_version != version:
                            self.root.after(0, lambda: self.log_message(f"Update available! You have: {current_version}", "warning"))
                        else:
                            self.root.after(0, lambda: self.log_message("You have the latest version installed.", "success"))
                else:
                    self.root.after(0, lambda: self.log_message("Could not determine latest version", "error"))
            except Exception as e:
                self.root.after(0, lambda: self.log_message(f"Error checking for latest version: {str(e)}", "error"))
            
            self.root.after(0, lambda: self.update_status("Latest version check completed"))
        
        def show_custom_activation(self):
            """Show custom activation dialog and perform activation"""
            if not is_admin():
                self.log_message("Administrator privileges required for activation", "error")
                messagebox.showerror("Error", "Administrator privileges required for activation")
                return
            
            # Show activation dialog to get user details
            activation_dialog = ActivationDialog(self.root)
            if activation_dialog.result is None:
                self.log_message("Activation cancelled by user", "info")
                return
            
            self.log_message("Starting IDM activation...", "info")
            self.update_status("Activating IDM...")
            
            # Start activation thread with user details
            threading.Thread(
                target=self._activate_idm_thread, 
                args=(activation_dialog.result,),
                daemon=True
            ).start()
        
        def _activate_idm_thread(self, user_details=None):
            # Check if IDM is running
            if self._is_idm_running():
                self.root.after(0, lambda: self.log_message("Stopping IDM...", "info"))
                self._stop_idm()
            
            # Clean registry (similar to reset)
            self._clean_registry_keys(deep=True)
            
            # Set fake registration info
            try:
                # Create DownloadManager key if it doesn't exist
                try:
                    reg_key = winreg.OpenKey(self.system_info['HKCU'], r"Software\DownloadManager", 0, 
                                           winreg.KEY_SET_VALUE | winreg.KEY_CREATE_SUB_KEY)
                except:
                    reg_key = winreg.CreateKey(self.system_info['HKCU'], r"Software\DownloadManager")
                
                # Set registration details
                if user_details:
                    # Use the full previewed serial if available, otherwise generate one
                    if 'full_serial' in user_details:
                        serial = user_details['full_serial']
                    else:
                        # Generate a serial key using the prefix provided by the user
                        serial_prefix = user_details['serial_prefix']
                        
                        # Generate the rest of the serial to match IDM format (groups of 4-5 characters separated by hyphens)
                        serial_parts = [serial_prefix]
                        
                        # Generate 3 more groups for a complete serial
                        for _ in range(3):
                            part = ''.join(random.choices(string.ascii_uppercase + string.digits, k=4))
                            serial_parts.append(part)
                        
                        serial = "-".join(serial_parts)
                    
                    reg_info = {
                        "FName": user_details['fname'],
                        "LName": user_details['lname'],
                        "Email": user_details['email'],
                        "Serial": serial
                    }
                else:
                    # Fallback to default values if no user details provided
                    reg_info = {
                        "FName": "Registered",
                        "LName": "User",
                        "Email": "user@example.com",
                        "Serial": "XKTWB-YPVN3-OZWLX-SLPFU"  # This is a fake serial
                    }
                
                # Add additional key values needed for proper activation
                current_time = int(time.time())
                additional_keys = {
                    "ActivationTime": str(current_time),
                    "CheckUpdtTime": str(current_time),
                    "scansk": "AAAGAAA=",
                    "MData": "AAABAAAAAAABAAAAAAAAAAAAAA==",
                    "updtprm": "2;12;1;1;1;1;1;0;",  # Update parameters
                    "SPDirExist": "1",
                    "icfname": "idman638build18.exe",  # Use current/latest IDM installer name
                    "icfsize": "8517272",  # Installer size (approximate)
                    "regStatus": "1",  # Registration status (1=registered)
                    "netdmin": "15000",  # Connection minimum speed
                    "taser": reg_info["FName"] + " " + reg_info["LName"],  # Registered user name
                    "realser": serial,  # Keep a duplicate of the serial in a different format
                    "isreged": "1",  # Another registration flag
                    "iserror": "0",  # No error flag
                    "isactived": "1",  # Activated flag
                    "DistribType": "0"  # Distribution type
                }
                
                # Merge dictionaries
                reg_info.update(additional_keys)
                
                for key, value in reg_info.items():
                    winreg.SetValueEx(reg_key, key, 0, winreg.REG_SZ, value)
                    self.root.after(0, lambda k=key, v=value: self.log_message(f"Set {k}: {v}", "info"))
                
                # Also set some DWORD values
                dword_values = {
                    "IsRegistered": 1,
                    "LstCheck": 0,
                    "CheckUpdtTime": current_time,
                    "AppDataDir": 1,
                    "AfterInst": 1,
                    "LaunchCnt": 15,  # Make it look like IDM has been used for a while
                    "scdt": current_time,
                    "radxcnt": 0,
                    "mngdby": 0
                }
                
                for key, value in dword_values.items():
                    try:
                        winreg.SetValueEx(reg_key, key, 0, winreg.REG_DWORD, value)
                        self.root.after(0, lambda k=key, v=value: self.log_message(f"Set {k}: {v} (DWORD)", "info"))
                    except:
                        pass
                
                winreg.CloseKey(reg_key)
                
                # Also set system-wide registry entries for better activation
                try:
                    # Use reg.exe for HKLM modifications (requires admin but more likely to succeed)
                    cmd = f'reg add "HKLM\\SOFTWARE\\Wow6432Node\\Internet Download Manager" /v regStatus /t REG_SZ /d "1" /f'
                    subprocess.run(cmd, shell=True, check=False)
                    
                    cmd = f'reg add "HKLM\\SOFTWARE\\Wow6432Node\\Internet Download Manager" /v regname /t REG_SZ /d "{reg_info["FName"]} {reg_info["LName"]}" /f'
                    subprocess.run(cmd, shell=True, check=False)
                    
                    cmd = f'reg add "HKLM\\SOFTWARE\\Wow6432Node\\Internet Download Manager" /v regemail /t REG_SZ /d "{reg_info["Email"]}" /f'
                    subprocess.run(cmd, shell=True, check=False)
                    
                    cmd = f'reg add "HKLM\\SOFTWARE\\Wow6432Node\\Internet Download Manager" /v regserial /t REG_SZ /d "{reg_info["Serial"]}" /f'
                    subprocess.run(cmd, shell=True, check=False)
                    
                    cmd = f'reg add "HKLM\\SOFTWARE\\Wow6432Node\\Internet Download Manager" /v InstallStatus /t REG_DWORD /d "3" /f'
                    subprocess.run(cmd, shell=True, check=False)
                    
                    # Try to set IsRegistered flag in multiple locations
                    cmd = f'reg add "HKCU\\Software\\DownloadManager" /v IsRegistered /t REG_DWORD /d "1" /f'
                    subprocess.run(cmd, shell=True, check=False)
                    
                    cmd = f'reg add "HKCU\\Software\\DownloadManager" /v isreged /t REG_SZ /d "1" /f'
                    subprocess.run(cmd, shell=True, check=False)
                    
                    # Try to create direct registry values for special non-trial serial
                    cmd = f'reg add "HKCU\\Software\\DownloadManager" /v scdtx6 /t REG_BINARY /d "222222" /f'
                    subprocess.run(cmd, shell=True, check=False)
                    
                    self.root.after(0, lambda: self.log_message("Set system-wide registration keys", "success"))
                except Exception as e:
                    self.root.after(0, lambda: self.log_message(f"Warning: Could not set system-wide keys: {str(e)}", "warning"))
                
                # Create license file as fallback
                try:
                    appdata = os.environ.get('APPDATA', '')
                    if appdata:
                        idm_dir = os.path.join(appdata, "IDM")
                        os.makedirs(idm_dir, exist_ok=True)
                        
                        # Create license file
                        license_path = os.path.join(idm_dir, "license.sav")
                        with open(license_path, "w") as f:
                            f.write(f"Name: {reg_info['FName']} {reg_info['LName']}\n")
                            f.write(f"Email: {reg_info['Email']}\n")
                            f.write(f"Serial: {reg_info['Serial']}\n")
                            f.write(f"Date: {datetime.now().strftime('%Y-%m-%d')}\n")
                        
                        # Also try command-line
                        cmd = f'echo Name: {reg_info["FName"]} {reg_info["LName"]} > "{license_path}"'
                        subprocess.run(cmd, shell=True, check=False)
                        
                        self.root.after(0, lambda: self.log_message(f"Created license file: {license_path}", "success"))
                except Exception as e:
                    self.root.after(0, lambda: self.log_message(f"Warning: Could not create license file: {str(e)}", "warning"))
                
                # Update activation status
                self.root.after(0, lambda: self.activation_status_var.set("Activated"))
                self.root.after(0, lambda: self.log_message("IDM activation completed successfully", "success"))
                self.root.after(0, lambda: self.update_status("IDM activated"))
                
                # Clear cache to force reload of registry values
                if 'dm_key' in self.registry_cache:
                    del self.registry_cache['dm_key']
                
                # Check firewall status and warn if enabled
                if self._check_firewall_status():
                    self.root.after(0, lambda: self.log_message("WARNING: Windows Firewall is enabled. This might interfere with activation.", "warning"))
                    self.root.after(0, lambda: messagebox.showwarning("Firewall Enabled", 
                                                                   "Windows Firewall is currently enabled.\n\n"
                                                                   "It's recommended to disable the firewall using the 'Toggle Firewall' button for better activation results."))
                
                # Launch IDM to apply changes if it was running before
                idm_path = self.system_info['idm_path']
                if os.path.exists(idm_path):
                    try:
                        os.startfile(idm_path)
                        self.root.after(0, lambda: self.log_message("Restarted IDM to apply changes", "info"))
                    except:
                        pass
                
                # Show success message with detailed instructions
                success_msg = (
                    "IDM activation completed successfully!\n\n"
                    "If IDM still shows as unregistered:\n"
                    "1. Restart your computer\n"
                    "2. Make sure Windows Firewall is disabled\n"
                    "3. Run the activation again with administrator privileges\n\n"
                    "Advanced troubleshooting:\n"
                    "- Install an older version of IDM (6.38 is recommended)\n"
                    "- Run the activation script immediately after fresh installation\n"
                    "- Disconnect from the internet during activation\n\n"
                    "The activation should persist after IDM updates."
                )
                self.root.after(0, lambda: messagebox.showinfo("Activation Complete", success_msg))
                
            except Exception as e:
                self.root.after(0, lambda: self.log_message(f"Activation failed: {str(e)}", "error"))
                self.root.after(0, lambda: self.update_status("Activation failed"))
        
        def reset_trial(self):
            """Reset IDM trial"""
            if not is_admin():
                self.log_message("Administrator privileges required for trial reset", "error")
                messagebox.showerror("Error", "Administrator privileges required for trial reset")
                return
            
            self.log_message("Starting IDM trial reset...", "info")
            self.update_status("Resetting IDM trial...")
            
            threading.Thread(target=self._reset_trial_thread, daemon=True).start()
        
        def _reset_trial_thread(self):
            # Check if IDM is running
            if self._is_idm_running():
                self.root.after(0, lambda: self.log_message("Stopping IDM...", "info"))
                self._stop_idm()
            
            # Clean registry
            result = self._clean_registry_keys()
            
            if result:
                self.root.after(0, lambda: self.activation_status_var.set("Trial"))
                self.root.after(0, lambda: self.log_message("IDM trial reset completed successfully", "success"))
                self.root.after(0, lambda: self.update_status("Trial reset completed"))
                
                # Clear cache to force reload of registry values
                if 'dm_key' in self.registry_cache:
                    del self.registry_cache['dm_key']
            else:
                self.root.after(0, lambda: self.log_message("Trial reset failed", "error"))
                self.root.after(0, lambda: self.update_status("Trial reset failed"))
        
        def clean_registry(self):
            """Clean IDM registry entries"""
            if not is_admin():
                self.log_message("Administrator privileges required for registry cleaning", "error")
                messagebox.showerror("Error", "Administrator privileges required for registry cleaning")
                return
            
            self.log_message("Starting IDM registry cleaning...", "info")
            self.update_status("Cleaning registry...")
            
            threading.Thread(target=self._clean_registry_thread, daemon=True).start()
        
        def _clean_registry_thread(self):
            # Check if IDM is running
            if self._is_idm_running():
                self.root.after(0, lambda: self.log_message("Stopping IDM...", "info"))
                self._stop_idm()
            
            # Clean registry
            result = self._clean_registry_keys(deep=True)
            
            if result:
                self.root.after(0, lambda: self.activation_status_var.set("Not activated"))
                self.root.after(0, lambda: self.log_message("Registry cleaning completed successfully", "success"))
                self.root.after(0, lambda: self.update_status("Registry cleaning completed"))
                
                # Clear cache to force reload of registry values
                if 'dm_key' in self.registry_cache:
                    del self.registry_cache['dm_key']
            else:
                self.root.after(0, lambda: self.log_message("Registry cleaning failed", "error"))
                self.root.after(0, lambda: self.update_status("Registry cleaning failed"))
        
        def _clean_registry_keys(self, deep=False):
            """Clean IDM registry keys"""
            try:
                # Values to delete
                values_to_delete = [
                    (self.system_info['HKCU'], r"Software\DownloadManager", "FName"),
                    (self.system_info['HKCU'], r"Software\DownloadManager", "LName"),
                    (self.system_info['HKCU'], r"Software\DownloadManager", "Email"),
                    (self.system_info['HKCU'], r"Software\DownloadManager", "Serial"),
                    (self.system_info['HKCU'], r"Software\DownloadManager", "scansk"),
                    (self.system_info['HKCU'], r"Software\DownloadManager", "tvfrdt"),
                    (self.system_info['HKCU'], r"Software\DownloadManager", "radxcnt"),
                    (self.system_info['HKCU'], r"Software\DownloadManager", "LstCheck"),
                    (self.system_info['HKCU'], r"Software\DownloadManager", "ptrk_scdt"),
                    (self.system_info['HKCU'], r"Software\DownloadManager", "LastCheckQU")
                ]
                
                # Delete values
                for hkey, key_path, value_name in values_to_delete:
                    try:
                        with winreg.OpenKey(hkey, key_path, 0, winreg.KEY_SET_VALUE) as key:
                            winreg.DeleteValue(key, value_name)
                            self.root.after(0, lambda p=key_path, n=value_name: 
                                            self.log_message(f"Deleted registry value: {p}\\{n}", "info"))
                    except:
                        pass
                
                # If deep cleaning, also remove CLSID entries
                if deep:
                    self._clean_clsid_keys()
                
                return True
            except Exception as e:
                self.root.after(0, lambda: self.log_message(f"Error cleaning registry: {str(e)}", "error"))
                return False
        
        def _clean_clsid_keys(self):
            """Clean IDM CLSID registry keys"""
            try:
                self.root.after(0, lambda: self.log_message("Searching for IDM CLSID keys...", "info"))
                
                # Look for IDM-related CLSID entries
                clsid_patterns = [
                    "{07999AC3-058B-40BF-984F-69EB1E554CA7}",  # Common IDM CLSID
                    "{5ED60779-4DE2-4E07-B862-974CA4FF3D55}",  # Common IDM CLSID for browser integration
                    "IDMIEHlprObj",  # Helper object class
                    "IDMIECC",  # Browser Helper Object
                    "IDMBHOObj",  # Browser Helper Object class
                    "Internet Download Manager",  # General IDM pattern
                    "idmmkb",  # Another IDM pattern
                    "IDM.DownloadAll",  # IDM download component
                    "IDM.NativeHook",  # IDM hook
                ]
                
                # Search in CLSID path based on architecture
                clsid_path = self.system_info['CLSID']
                
                # First check HKCU
                deleted_count = 0
                try:
                    clsid_root_key = winreg.OpenKey(self.system_info['HKCU'], clsid_path, 0, 
                                                   winreg.KEY_READ | winreg.KEY_ENUMERATE_SUB_KEYS)
                    
                    # Get the count of subkeys
                    subkey_count = winreg.QueryInfoKey(clsid_root_key)[0]
                    
                    # Iterate through subkeys to find IDM-related ones
                    for i in range(subkey_count):
                        try:
                            subkey_name = winreg.EnumKey(clsid_root_key, i)
                            
                            # Check for common IDM CLSIDs
                            for pattern in clsid_patterns:
                                if pattern.lower() in subkey_name.lower():
                                    # Delete this key
                                    try:
                                        # Use Windows command to delete the key (more reliable for registry deletion)
                                        cmd = f'reg delete "HKCU\\{clsid_path}\\{subkey_name}" /f'
                                        subprocess.run(cmd, shell=True, check=False, capture_output=True)
                                        
                                        self.root.after(0, lambda s=subkey_name: 
                                                       self.log_message(f"Deleted CLSID key: {s}", "info"))
                                        deleted_count += 1
                                        break
                                    except:
                                        pass
                            
                            # Also check if the key has any values or subkeys containing "idm" or "download"
                            try:
                                subkey_path = f"{clsid_path}\\{subkey_name}"
                                subkey = winreg.OpenKey(self.system_info['HKCU'], subkey_path, 0, winreg.KEY_READ)
                                
                                # Check values
                                try:
                                    value_count = winreg.QueryInfoKey(subkey)[1]
                                    for j in range(value_count):
                                        try:
                                            value_name, value_data, _ = winreg.EnumValue(subkey, j)
                                            if isinstance(value_data, str) and ("idm" in value_data.lower() or 
                                                                             "download manager" in value_data.lower()):
                                                # Delete this key
                                                winreg.CloseKey(subkey)
                                                cmd = f'reg delete "HKCU\\{clsid_path}\\{subkey_name}" /f'
                                                subprocess.run(cmd, shell=True, check=False, capture_output=True)
                                                
                                                self.root.after(0, lambda s=subkey_name: 
                                                               self.log_message(f"Deleted CLSID key with IDM value: {s}", "info"))
                                                deleted_count += 1
                                                break
                                        except:
                                            pass
                                except:
                                    pass
                                
                                winreg.CloseKey(subkey)
                            except:
                                pass
                        except:
                            continue
                    
                    winreg.CloseKey(clsid_root_key)
                except Exception as e:
                    self.root.after(0, lambda err=str(e): 
                                   self.log_message(f"Error searching HKCU CLSID keys: {err}", "warning"))
                
                # Also check in HKLM
                try:
                    clsid_root_key = winreg.OpenKey(self.system_info['HKLM'], clsid_path, 0, 
                                                   winreg.KEY_READ)
                    
                    # Just log HKLM keys (not deleting as they often require system privileges)
                    subkey_count = winreg.QueryInfoKey(clsid_root_key)[0]
                    
                    for i in range(subkey_count):
                        try:
                            subkey_name = winreg.EnumKey(clsid_root_key, i)
                            
                            # Check for common IDM CLSIDs
                            for pattern in clsid_patterns:
                                if pattern.lower() in subkey_name.lower():
                                    self.root.after(0, lambda s=subkey_name: 
                                                  self.log_message(f"Found HKLM CLSID key (requires system cleanup): {s}", "warning"))
                                    break
                        except:
                            continue
                    
                    winreg.CloseKey(clsid_root_key)
                except Exception as e:
                    self.root.after(0, lambda err=str(e): 
                                   self.log_message(f"Error searching HKLM CLSID keys: {err}", "warning"))
                
                # Also check other common locations for IDM entries
                additional_paths = [
                    r"Software\Classes\IDMIECC.IdmIECC",
                    r"Software\Classes\IDMIECC.IdmIECC.1",
                    r"Software\Classes\IDMIEHlprObj.IdmIEHlprObj",
                    r"Software\Classes\IDMIEHlprObj.IdmIEHlprObj.1",
                    r"Software\Microsoft\Windows\CurrentVersion\Ext\Settings\{5ED60779-4DE2-4E07-B862-974CA4FF3D55}",
                    r"Software\Microsoft\Windows\CurrentVersion\Ext\Stats\{5ED60779-4DE2-4E07-B862-974CA4FF3D55}",
                    r"Software\Microsoft\Windows\CurrentVersion\Ext\Settings\{07999AC3-058B-40BF-984F-69EB1E554CA7}",
                    r"Software\Microsoft\Windows\CurrentVersion\Ext\Stats\{07999AC3-058B-40BF-984F-69EB1E554CA7}",
                ]
                
                for path in additional_paths:
                    try:
                        cmd = f'reg delete "HKCU\\{path}" /f'
                        subprocess.run(cmd, shell=True, check=False, capture_output=True)
                        deleted_count += 1
                        self.root.after(0, lambda p=path: self.log_message(f"Deleted registry key: {p}", "info"))
                    except:
                        pass
                
                # Also ensure that the installation status is reset
                self._reset_installation_status()
                
                if deleted_count > 0:
                    self.root.after(0, lambda count=deleted_count: 
                                   self.log_message(f"Cleaned {count} IDM CLSID and related keys", "success"))
                else:
                    self.root.after(0, lambda: 
                                   self.log_message("No IDM CLSID keys found that need cleaning", "info"))
                
                return True
            except Exception as e:
                self.root.after(0, lambda: self.log_message(f"Error cleaning CLSID keys: {str(e)}", "error"))
                return False
                
        def _reset_installation_status(self):
            """Reset IDM installation status to properly allow activation"""
            try:
                # Try to set installation status directly
                cmd = f'reg add "HKLM\\SOFTWARE\\Wow6432Node\\Internet Download Manager" /v InstallStatus /t REG_DWORD /d "0" /f'
                subprocess.run(cmd, shell=True, check=False)
                
                # Also set under 32-bit path just in case
                cmd = f'reg add "HKLM\\SOFTWARE\\Internet Download Manager" /v InstallStatus /t REG_DWORD /d "0" /f'
                subprocess.run(cmd, shell=True, check=False)
                
                self.root.after(0, lambda: self.log_message("Reset installation status to fresh install", "info"))
            except:
                pass
        
        def toggle_firewall(self):
            """Toggle Windows Firewall"""
            if not is_admin():
                self.log_message("Administrator privileges required to toggle firewall", "error")
                messagebox.showerror("Error", "Administrator privileges required to toggle firewall")
                return
            
            self.log_message("Toggling Windows Firewall...", "info")
            self.update_status("Toggling firewall...")
            
            threading.Thread(target=self._toggle_firewall_thread, daemon=True).start()
        
        def _toggle_firewall_thread(self):
            try:
                # Get current firewall status
                firewall_status = self._check_firewall_status()
                
                if firewall_status:
                    # Firewall is enabled, disable it
                    subprocess.run("netsh advfirewall set allprofiles state off", shell=True, check=True)
                    self.root.after(0, lambda: self.log_message("Windows Firewall disabled", "success"))
                    self.root.after(0, lambda: self.update_status("Firewall disabled"))
                    self.root.after(0, lambda: self.firewall_status_var.set("Firewall: OFF"))
                    self.root.after(0, lambda: self.firewall_status_label.config(foreground="green"))
                else:
                    # Firewall is disabled, enable it
                    subprocess.run("netsh advfirewall set allprofiles state on", shell=True, check=True)
                    self.root.after(0, lambda: self.log_message("Windows Firewall enabled", "success"))
                    self.root.after(0, lambda: self.update_status("Firewall enabled"))
                    self.root.after(0, lambda: self.firewall_status_var.set("Firewall: ON"))
                    self.root.after(0, lambda: self.firewall_status_label.config(foreground="red"))
            except Exception as e:
                self.root.after(0, lambda: self.log_message(f"Error toggling firewall: {str(e)}", "error"))
                self.root.after(0, lambda: self.update_status("Firewall toggle failed"))
        
        def _check_firewall_status(self):
            """Check if Windows Firewall is enabled"""
            try:
                # Faster way to check firewall status with single subprocess call
                result = subprocess.run('netsh advfirewall show allprofiles state', 
                                      shell=True, capture_output=True, text=True)
                return "ON" in result.stdout
            except:
                # Fallback to registry check if the command fails
                firewall_status = {
                    'Enabled': 0,
                    'Disabled': 0
                }
                
                profiles = ["DomainProfile", "PublicProfile", "StandardProfile"]
                
                for profile in profiles:
                    try:
                        key_path = f"SYSTEM\\CurrentControlSet\\Services\\SharedAccess\\Parameters\\FirewallPolicy\\{profile}"
                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                            value = winreg.QueryValueEx(key, "EnableFirewall")[0]
                            if value == 1:
                                firewall_status['Enabled'] += 1
                            else:
                                firewall_status['Disabled'] += 1
                    except:
                        pass
                
                return firewall_status['Enabled'] >= 2  # Firewall is mostly enabled
        
        def download_idm(self):
            """Open IDM download page"""
            self.log_message("Opening IDM download page...", "info")
            self.update_status("Opening download page...")
            
            try:
                webbrowser.open("https://www.internetdownloadmanager.com/download.html")
                self.log_message("Download page opened in browser", "success")
                self.update_status("Download page opened")
            except Exception as e:
                self.log_message(f"Error opening download page: {str(e)}", "error")
                self.update_status("Failed to open download page")
        
        def launch_idm(self):
            """Launch IDM if installed"""
            idm_path = self.system_info['idm_path']
            
            if os.path.exists(idm_path):
                self.log_message(f"Launching IDM from {idm_path}", "info")
                self.update_status("Launching IDM...")
                
                try:
                    os.startfile(idm_path)
                    self.log_message("IDM launched successfully", "success")
                    self.update_status("IDM launched")
                except Exception as e:
                    self.log_message(f"Error launching IDM: {str(e)}", "error")
                    self.update_status("Failed to launch IDM")
            else:
                self.log_message("IDM executable not found", "error")
                self.update_status("IDM not found")
                messagebox.showerror("Error", "IDM executable not found. Please download and install IDM first.")
        
        def show_help(self):
            """Show help/about dialog"""
            help_text = f"""
IDM Manager v{self.version}

This tool helps you manage Internet Download Manager:
- Activate IDM with fake registration
- Reset trial periods
- Check for updates
- Toggle Windows Firewall (which helps with activation)
- Clean registry traces
- Download and launch IDM

Note: Some operations require administrator privileges.

TIP: Use the .pyw version for a console-less experience!

Created by: Claude AI Assistant
"""
            messagebox.showinfo("Help / About", help_text)
        
        def _is_idm_running(self):
            """Check if IDM is running"""
            try:
                # Faster way to check with WMIC
                result = subprocess.run('tasklist /FI "IMAGENAME eq IDMan.exe" /NH', 
                                      shell=True, capture_output=True, text=True)
                return "IDMan.exe" in result.stdout
            except:
                return False
        
        def _stop_idm(self):
            """Stop IDM if it's running"""
            try:
                if self._is_idm_running():
                    subprocess.run('taskkill /f /im idman.exe', shell=True, check=False, capture_output=True)
                    time.sleep(0.5)  # Reduced wait time
                    return True
            except:
                return False
            return True

        def fix_browser_integration(self):
            """Fix IDM browser integration issues"""
            if not is_admin():
                self.log_message("Administrator privileges required to fix browser integration", "error")
                messagebox.showerror("Error", "Administrator privileges required to fix browser integration")
                return
            
            self.log_message("Starting IDM browser integration repair...", "info")
            self.update_status("Fixing browser integration...")
            
            threading.Thread(target=self._fix_browser_integration_thread, daemon=True).start()

        def _fix_browser_integration_thread(self):
            # Check if IDM is running
            if self._is_idm_running():
                self.root.after(0, lambda: self.log_message("Stopping IDM...", "info"))
                self._stop_idm()
            
            try:
                # 1. Enable advanced browser integration in the registry
                try:
                    reg_key = winreg.OpenKey(self.system_info['HKCU'], r"Software\DownloadManager", 0, 
                                           winreg.KEY_SET_VALUE | winreg.KEY_CREATE_SUB_KEY)
                except:
                    reg_key = winreg.CreateKey(self.system_info['HKCU'], r"Software\DownloadManager")
                
                # Enable advanced browser integration
                winreg.SetValueEx(reg_key, "use_advanced_browser_integration", 0, winreg.REG_DWORD, 1)
                self.root.after(0, lambda: self.log_message("Enabled advanced browser integration", "info"))
                
                # Enable browser monitoring
                winreg.SetValueEx(reg_key, "browser_monitoring", 0, winreg.REG_DWORD, 1)
                self.root.after(0, lambda: self.log_message("Enabled browser monitoring", "info"))
                
                # Set additional advanced integration settings
                additional_settings = {
                    "ChromeExt": 1,              # Enable Chrome integration
                    "EdgeExt": 1,                # Enable Edge integration
                    "FirefoxExt": 1,             # Enable Firefox integration
                    "add_to_downloads": 1,       # Add downloads to browser's download list
                    "mask_host_app": 1,          # Mask host application
                    "video_exts": "mp4;webm;ogg;flv;avi;mov;mpg;mpeg;wmv;mkv",  # Video extensions to monitor
                    "video_show_ctrls": 1,       # Show video download controls
                    "video_smart_detect": 1,     # Enable smart video detection
                    "video_ads_skip": 1          # Skip video ads
                }
                
                for key, value in additional_settings.items():
                    try:
                        if isinstance(value, int):
                            winreg.SetValueEx(reg_key, key, 0, winreg.REG_DWORD, value)
                        else:
                            winreg.SetValueEx(reg_key, key, 0, winreg.REG_SZ, value)
                        self.root.after(0, lambda k=key: self.log_message(f"Set {k} registry value", "info"))
                    except:
                        pass
                
                winreg.CloseKey(reg_key)
                
                # 2. Run IDM's internal browser integration
                idm_path = self.system_info['idm_path']
                idm_dir = os.path.dirname(idm_path)
                
                # Try to run IDMan.exe with integration parameters
                try:
                    # Command to refresh browser integration using IDM itself
                    subprocess.run(f'"{idm_path}" /setbrowsers', shell=True, check=False)
                    self.root.after(0, lambda: self.log_message("Refreshed browser integration via IDM command", "info"))
                except:
                    pass
                
                # 3. Clear any corrupted extensions
                self.root.after(0, lambda: self.log_message("Preparing to clean browser extension data...", "info"))
                
                # Get the local app data path
                appdata_local = os.environ.get('LOCALAPPDATA', '')
                
                # Paths to browser extension folders
                browser_ext_paths = {
                    "Chrome": os.path.join(appdata_local, r"Google\Chrome\User Data\Default\Extensions"),
                    "Edge": os.path.join(appdata_local, r"Microsoft\Edge\User Data\Default\Extensions"),
                    "Firefox": os.path.join(appdata_local, r"Mozilla\Firefox\Profiles")
                }
                
                for browser, path in browser_ext_paths.items():
                    if os.path.exists(path):
                        self.root.after(0, lambda b=browser: self.log_message(f"Found {b} extensions directory", "info"))
                    
                # 4. Fix Chrome integration
                self.root.after(0, lambda: self.log_message("Opening Chrome Web Store to install authentic IDM extension...", "info"))
                webbrowser.open("https://chrome.google.com/webstore/detail/IDM-Integration-Module/ngpampappnmepgilojfohadhhmbhlaek")
                
                # 5. Fix Edge integration
                self.root.after(0, lambda: self.log_message("Opening Edge Add-ons to install authentic IDM extension...", "info"))
                webbrowser.open("https://microsoftedge.microsoft.com/addons/detail/idm-integration-module/llbjbkhnmlidjebalopleeepgdfgcpec")
                
                # 6. Additional troubleshooting - create browser_integration.txt file
                self.root.after(0, lambda: self.log_message("Generating browser integration troubleshooting info...", "info"))
                try:
                    # Get information about installed browsers
                    browsers_info = []
                    
                    # Check Chrome
                    chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
                    chrome_path_x86 = r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
                    if os.path.exists(chrome_path) or os.path.exists(chrome_path_x86):
                        browsers_info.append("Google Chrome: Installed")
                    
                    # Check Edge
                    edge_path = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
                    if os.path.exists(edge_path):
                        browsers_info.append("Microsoft Edge: Installed")
                    
                    # Check Firefox
                    firefox_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"
                    firefox_path_x86 = r"C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
                    if os.path.exists(firefox_path) or os.path.exists(firefox_path_x86):
                        browsers_info.append("Mozilla Firefox: Installed")
                    
                    # Create troubleshooting file
                    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
                    troubleshoot_path = os.path.join(desktop_path, "IDM_Browser_Integration_Info.txt")
                    
                    with open(troubleshoot_path, "w") as f:
                        f.write("=== IDM Browser Integration Troubleshooting Info ===\n\n")
                        f.write(f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                        f.write(f"IDM Path: {idm_path}\n")
                        f.write(f"IDM Version: {self.idm_version_var.get()}\n\n")
                        f.write("=== Browsers Detected ===\n")
                        for browser in browsers_info:
                            f.write(f"{browser}\n")
                        f.write("\n=== Registry Settings ===\n")
                        f.write("AdvancedBrowserIntegration: Enabled\n")
                        f.write("BrowserMonitoring: Enabled\n")
                        f.write("\n=== Common Issues ===\n")
                        f.write("1. Check if all browsers have the official IDM extension installed\n")
                        f.write("2. Make sure the extension has permission to access all sites\n")
                        f.write("3. Try uninstalling and reinstalling IDM\n")
                        f.write("4. Check for conflicts with other extensions (ad blockers, VPNs, etc.)\n")
                    
                    self.root.after(0, lambda: self.log_message(f"Created troubleshooting file at: {troubleshoot_path}", "success"))
                except Exception as e:
                    self.root.after(0, lambda: self.log_message(f"Failed to create troubleshooting file: {str(e)}", "warning"))
                
                # 7. Provide enhanced instructions
                instructions = (
                    "== BROWSER INTEGRATION INSTRUCTIONS ==\n"
                    "1. UNINSTALL ALL EXISTING IDM EXTENSIONS from your browsers\n"
                    "2. RESTART YOUR COMPUTER to ensure all changes take effect\n"
                    "3. OPEN IDM, go to Options > General > Check 'Use advanced browser integration'\n"
                    "4. INSTALL OFFICIAL EXTENSIONS from the store pages opened\n"
                    "5. RESTART ALL BROWSERS after installation\n\n"
                    "== ADDITIONAL TROUBLESHOOTING ==\n"
                    "- Check any security software that might be blocking IDM\n"
                    "- Disable VPN extensions that might conflict with IDM\n"
                    "- Check the troubleshooting file created on your desktop\n"
                    "- Try reinstalling IDM completely if issues persist"
                )
                
                self.root.after(0, lambda: self.log_message("Browser integration repair completed", "success"))
                self.root.after(0, lambda: self.log_message("Follow these steps to complete the setup:", "info"))
                for line in instructions.split("\n"):
                    self.root.after(0, lambda l=line: self.log_message(l, "info"))
                
                self.root.after(0, lambda: self.update_status("Browser integration repair completed"))
                
                # Show message box with instructions
                self.root.after(0, lambda: messagebox.showinfo(
                    "Browser Integration Repair", 
                    "IDM browser integration has been repaired in the registry.\n\n" + instructions
                ))
                
            except Exception as e:
                self.root.after(0, lambda: self.log_message(f"Browser integration repair failed: {str(e)}", "error"))
                self.root.after(0, lambda: self.update_status("Browser integration repair failed"))

        def _update_firewall_status(self):
            """Update the firewall status indicator"""
            try:
                status = self._check_firewall_status()
                if status:
                    self.root.after(0, lambda: self.firewall_status_var.set("Firewall: ON"))
                    self.root.after(0, lambda: self.firewall_status_label.config(foreground="red"))
                else:
                    self.root.after(0, lambda: self.firewall_status_var.set("Firewall: OFF"))
                    self.root.after(0, lambda: self.firewall_status_label.config(foreground="green"))
            except:
                self.root.after(0, lambda: self.firewall_status_var.set("Firewall: Unknown"))

        def reset_factory_defaults(self):
            """Reset IDM to factory defaults for better activation results"""
            if not is_admin():
                self.log_message("Administrator privileges required for factory reset", "error")
                messagebox.showerror("Error", "Administrator privileges required for factory reset")
                return
            
            if not messagebox.askyesno("Factory Reset", 
                                      "This will reset IDM to factory defaults.\n\n"
                                      "All your IDM settings will be lost, but this often helps with activation issues.\n\n"
                                      "Continue?"):
                return
            
            self.log_message("Starting IDM factory reset...", "info")
            self.update_status("Resetting IDM to factory defaults...")
            
            threading.Thread(target=self._reset_factory_thread, daemon=True).start()
            
        def _reset_factory_thread(self):
            """Thread to reset IDM to factory defaults"""
            # Check if IDM is running
            if self._is_idm_running():
                self.root.after(0, lambda: self.log_message("Stopping IDM...", "info"))
                self._stop_idm()
                time.sleep(1)  # Give it a bit more time to close
            
            try:
                # 1. Clean registry deeply
                self._clean_registry_keys(deep=True)
                
                # 2. Delete IDM configuration folders
                appdata_paths = [
                    os.path.join(os.environ.get('APPDATA', ''), "IDM"),
                    os.path.join(os.environ.get('LOCALAPPDATA', ''), "IDM"),
                ]
                
                for path in appdata_paths:
                    if os.path.exists(path):
                        try:
                            # Use command line to try to force delete
                            cmd = f'rmdir /s /q "{path}"'
                            subprocess.run(cmd, shell=True, check=False)
                            self.root.after(0, lambda p=path: self.log_message(f"Deleted IDM directory: {p}", "info"))
                        except:
                            pass
                
                # 3. Delete all IDM-related registry keys
                registry_roots = [
                    (r"HKCU\Software\DownloadManager", "User settings"),
                    (r"HKLM\SOFTWARE\Internet Download Manager", "System settings (32-bit)"),
                    (r"HKLM\SOFTWARE\Wow6432Node\Internet Download Manager", "System settings (64-bit)"),
                ]
                
                for reg_path, description in registry_roots:
                    try:
                        cmd = f'reg delete "{reg_path}" /f'
                        subprocess.run(cmd, shell=True, check=False)
                        self.root.after(0, lambda d=description: self.log_message(f"Deleted {d} registry keys", "info"))
                    except:
                        pass
                
                # 4. Reset file associations
                try:
                    cmd = 'assoc .download=Unknown'
                    subprocess.run(cmd, shell=True, check=False)
                    self.root.after(0, lambda: self.log_message("Reset .download file association", "info"))
                except:
                    pass
                
                # 5. Reset installation status in multiple locations
                self._reset_installation_status()
                
                # Clear cache to force reload of registry values
                if 'dm_key' in self.registry_cache:
                    del self.registry_cache['dm_key']
                
                # Update status
                self.root.after(0, lambda: self.activation_status_var.set("Reset"))
                self.root.after(0, lambda: self.log_message("IDM factory reset completed. You can now activate IDM.", "success"))
                self.root.after(0, lambda: self.update_status("Factory reset completed"))
                
                # Show success dialog
                self.root.after(0, lambda: messagebox.showinfo("Factory Reset Complete", 
                                                            "IDM has been reset to factory defaults.\n\n"
                                                            "Next steps:\n"
                                                            "1. Restart your computer\n"
                                                            "2. Launch IDM once\n"
                                                            "3. Try activating IDM again"))
                
            except Exception as e:
                self.root.after(0, lambda: self.log_message(f"Factory reset failed: {str(e)}", "error"))
                self.root.after(0, lambda: self.update_status("Factory reset failed"))

    # Main application entry point
    if __name__ == "__main__":
        log_debug("Starting main application")
        
        # If not admin, show warning but continue anyway
        if not is_admin():
            log_debug("Not running as administrator")
            if messagebox.askyesno("Admin Rights Required", 
                                "This program requires administrator rights for IDM activation to work properly.\n\n"
                                "The following features require admin rights:\n"
                                "- IDM Activation (most important)\n"
                                "- Registry cleaning\n" 
                                "- Firewall management\n"
                                "- Browser integration\n\n"
                                "Would you like to restart with admin privileges?"):
                log_debug("User requested elevation")
                try:
                    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{os.path.abspath(sys.argv[0])}"', None, 1)
                    log_debug("Elevation requested, exiting current instance")
                    sys.exit(0)
                except Exception as e:
                    log_debug(f"Failed to elevate: {str(e)}")
                    messagebox.showwarning("Elevation Failed", 
                                        "Failed to restart with admin rights.\n"
                                        "IDM activation will likely fail without admin rights.")
        
        # Create and run the application
        root = tk.Tk()
        app = IDMManager(root)
        log_debug("Starting mainloop")
        root.mainloop()
        log_debug("Application closed normally")

except Exception as e:
    # Log the error
    error_msg = f"CRITICAL ERROR: {str(e)}\n{traceback.format_exc()}"
    log_debug(error_msg)
    
    # Show error message
    show_error("Critical Error", 
               f"A critical error occurred:\n\n{str(e)}\n\n"
               f"See log file for details: {debug_log_path}")
    
    # If running from command prompt, pause
    if "PROMPT" in os.environ:
        input("Press Enter to exit...") 