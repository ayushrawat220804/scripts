import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog, simpledialog
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
import platform
import glob
import winreg
from datetime import datetime
import time
import csv
import socket
import functools
import traceback
import concurrent.futures
import fnmatch

try:
    import winshell
except ImportError:
    pass  # Handle later in the code

# Constants for UI
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
                            f"An unexpected error occurred:\n\n{str(exc_value)}\n\nSee console for details.")
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
                                        r"SOFTWARE\Microsoft\Windows NT\CurrentVersion")
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
            messagebox.showerror("Error", f"An unexpected error occurred in {func.__name__}:\n\n{str(e)}")
            return None
    return wrapper

def thread_safe(func):
    """Decorator for ensuring thread-safe UI updates"""
    @functools.wraps(func)
    def wrapper(self, *args, **kwargs):
        with ui_lock:
            return func(self, *args, **kwargs)
    return wrapper

class SystemUtilities:
    def __init__(self, root):
        self.root = root
        self.root.title("Windows System Utilities")
        
        # Make window open in fullscreen by default
        self.root.state('zoomed')  # For Windows, this maximizes the window
        
        # Get screen dimensions for responsive design
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # Set minimum size for the window
        self.root.minsize(900, 600)
        
        # Set the application icon
        try:
            self.root.iconbitmap("system_icon.ico")
        except:
            # Icon file not found, continue without it
            pass
        
        # Configure the style for better appearance
        self.style = ttk.Style()
        # Apply clam theme if available
        if 'clam' in self.style.theme_names():
            self.style.theme_use('clam')
            
        self.style.configure('TFrame', background=BG_COLOR)
        self.style.configure('Tab.TFrame', background=BG_COLOR)
        self.style.configure('TLabel', background=BG_COLOR, foreground=TEXT_COLOR, font=NORMAL_FONT)
        self.style.configure('Header.TLabel', font=HEADING_FONT, foreground=PRIMARY_COLOR)
        self.style.configure('TButton', font=BUTTON_FONT)
        self.style.configure('Primary.TButton', background=PRIMARY_COLOR)
        self.style.configure('Secondary.TButton', background=SECONDARY_COLOR)
        self.style.configure('Accent.TButton', background=ACCENT_COLOR)
        self.style.configure('TLabelframe', background=BG_COLOR)
        self.style.configure('Group.TLabelframe', background=BG_COLOR)
        self.style.configure('Group.TLabelframe.Label', background=BG_COLOR, foreground=TEXT_COLOR, font=NORMAL_FONT)
        
        # Set up the main frame with split-pane layout
        main_frame = ttk.Frame(root, style='TFrame')
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create PanedWindow for split layout
        self.paned = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Left side - Tool tabs
        self.left_frame = ttk.Frame(self.paned, style='TFrame')
        
        # Right side - Logs and status
        self.right_frame = ttk.Frame(self.paned, style='TFrame')
        
        # Add the frames to the paned window
        self.paned.add(self.left_frame, weight=3)
        self.paned.add(self.right_frame, weight=1)
        
        # Create tabs on the left side
        self.tabs = ttk.Notebook(self.left_frame)
        self.tabs.pack(fill=tk.BOTH, expand=True)
        
        # Create the tabs
        self.cleanup_tab = ttk.Frame(self.tabs, style='Tab.TFrame')
        self.system_tab = ttk.Frame(self.tabs, style='Tab.TFrame')
        self.update_tab = ttk.Frame(self.tabs, style='Tab.TFrame')
        self.storage_tab = ttk.Frame(self.tabs, style='Tab.TFrame')
        self.network_tab = ttk.Frame(self.tabs, style='Tab.TFrame')
        self.hyperv_tab = ttk.Frame(self.tabs, style='Tab.TFrame')
        self.optimize_tab = ttk.Frame(self.tabs, style='Tab.TFrame')
        
        # Add tabs to the notebook
        self.tabs.add(self.cleanup_tab, text="Cleanup")
        self.tabs.add(self.system_tab, text="System Tools")
        self.tabs.add(self.update_tab, text="Windows Update")
        self.tabs.add(self.storage_tab, text="Storage")
        self.tabs.add(self.network_tab, text="Network")
        self.tabs.add(self.hyperv_tab, text="Hyper-V")
        self.tabs.add(self.optimize_tab, text="Optimize")
        
        # Add tab change event 
        self.tabs.bind("<<NotebookTabChanged>>", self.on_tab_changed)
        
        # Set up the log area on the right side
        self.create_log_area()
        
        # Initialize tab contents
        self.init_cleanup_tab()
        self.init_system_tab()
        self.init_update_tab()
        self.init_storage_tab()
        self.init_network_tab()
        self.init_hyperv_tab()
        self.init_optimize_tab()
        
        # Add status bar at the bottom
        self.status_bar = ttk.Label(root, text="Ready", relief=tk.SUNKEN, anchor=tk.W, font=NORMAL_FONT)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Initialize the thread pool for background tasks
        self.executor = concurrent.futures.ThreadPoolExecutor(max_workers=5)
        
        # Set background monitoring parameters
        self.background_monitoring = True
        self.show_verbose_logs = True
        
        # Show a welcome message
        self.log("System Utilities initialized successfully", "info")
        self.update_status("Ready")
        
        # Start background monitoring
        self.start_background_monitoring()
        
        # Periodically collect garbage to prevent memory leaks
        def garbage_collect():
            import gc
            gc.collect()
            self.root.after(300000, garbage_collect)  # Every 5 minutes
        
        # Start the garbage collection timer
        self.root.after(300000, garbage_collect)
    
    def on_tab_changed(self, event):
        """Handle tab change event to collect and show relevant system information"""
        tab_id = self.tabs.select()
        tab_name = self.tabs.tab(tab_id, "text")
        
        # Log the tab change
        self.log(f"Switched to {tab_name} tab", "info")
        
        # Show relevant system information based on the tab
        if tab_name == "Network":
            self.log_network_info()
        elif tab_name == "Hyper-V":
            self.log_hyperv_info()
        elif tab_name == "Optimize":
            self.log_performance_info()
        elif tab_name == "Storage":
            self.log_storage_info()
    
    def start_background_monitoring(self):
        """Start background monitoring threads"""
        self.log("Starting background system monitoring...", "background")
        
        # Start network monitoring thread
        network_thread = Thread(target=self._monitor_network_activity)
        network_thread.daemon = True
        network_thread.start()
        
        # Start performance monitoring thread
        performance_thread = Thread(target=self._monitor_system_performance)
        performance_thread.daemon = True
        performance_thread.start()
        
        # Start Hyper-V monitoring thread (if applicable)
        hyperv_thread = Thread(target=self._monitor_hyperv_status)
        hyperv_thread.daemon = True
        hyperv_thread.start()
    
    def update_background_logs(self):
        """Update background activity logs periodically"""
        self.log("Background monitoring active", "background")
        
        # Log system info
        cpu_percent = psutil.cpu_percent(interval=0.5)
        memory = psutil.virtual_memory()
        
        if cpu_percent > 80:
            self.log(f"High CPU usage: {cpu_percent}%", "system")
        
        if memory.percent > 85:
            self.log(f"High memory usage: {memory.percent}%", "system")
        
        # Get network activity
        try:
            connections = len(psutil.net_connections())
            if connections > 100:
                self.log(f"High network activity: {connections} connections", "network")
        except:
            pass
        
        # Schedule the next update (every 30 seconds)
        self.root.after(30000, self.update_background_logs)
    
    def _monitor_network_activity(self):
        """Background thread to monitor network activity"""
        try:
            # Initialize network stats
            prev_net_io = psutil.net_io_counters()
            
            while self.background_monitoring:
                time.sleep(10)  # Check every 10 seconds
                
                try:
                    # Get current network stats
                    net_io = psutil.net_io_counters()
                    
                    # Calculate data transferred since last check
                    sent_mb = (net_io.bytes_sent - prev_net_io.bytes_sent) / (1024 * 1024)
                    recv_mb = (net_io.bytes_recv - prev_net_io.bytes_recv) / (1024 * 1024)
                    
                    # Log significant network activity
                    if sent_mb > 5 or recv_mb > 5:  # Only log if more than 5MB in 10 seconds
                        self.root.after(0, lambda s=sent_mb, r=recv_mb: 
                                    self.log(f"Data transfer: {s:.2f}MB sent, {r:.2f}MB received in last 10 seconds", "network"))
                    
                    # Update previous stats
                    prev_net_io = net_io
                    
                    # Check for connection changes every 30 seconds
                    if time.time() % 30 < 10:
                        try:
                            # Get connection stats (requires admin on Windows)
                            # This might fail on non-admin, which is fine
                            connections = len(psutil.net_connections())
                            self.root.after(0, lambda c=connections: 
                                        self.log(f"Active network connections: {c}", "network"))
                        except:
                            pass
                except Exception as e:
                    # Silent fail for background monitoring
                    print(f"Network monitoring error: {str(e)}")
        
        except Exception as e:
            print(f"Error in network monitoring thread: {str(e)}")
    
    def _monitor_system_performance(self):
        """Background thread to monitor system performance"""
        try:
            while self.background_monitoring:
                time.sleep(15)  # Check every 15 seconds
                
                try:
                    # Get system performance metrics
                    cpu_percent = psutil.cpu_percent(interval=1)
                    memory = psutil.virtual_memory()
                    
                    # Log when system is under high load
                    if cpu_percent > 75:
                        self.root.after(0, lambda p=cpu_percent: 
                                    self.log(f"High CPU usage: {p}%", "system"))
                    
                    if memory.percent > 80:
                        self.root.after(0, lambda p=memory.percent: 
                                    self.log(f"High memory usage: {p}%", "system"))
                    
                    # Periodically log general system info (every 60 seconds)
                    if time.time() % 60 < 15:
                        # Get disk I/O
                        try:
                            disk_io = psutil.disk_io_counters()
                            if hasattr(self, 'prev_disk_io'):
                                read_mb = (disk_io.read_bytes - self.prev_disk_io.read_bytes) / (1024 * 1024)
                                write_mb = (disk_io.write_bytes - self.prev_disk_io.write_bytes) / (1024 * 1024)
                                
                                if read_mb > 50 or write_mb > 50:  # Only log high activity
                                    self.root.after(0, lambda r=read_mb, w=write_mb: 
                                                self.log(f"Disk activity: {r:.2f}MB read, {w:.2f}MB written in last minute", "system"))
                            
                            self.prev_disk_io = disk_io
                        except:
                            pass
                    
                except Exception as e:
                    # Silent fail for background monitoring
                    print(f"Performance monitoring error: {str(e)}")
        
        except Exception as e:
            print(f"Error in performance monitoring thread: {str(e)}")
    
    def _monitor_hyperv_status(self):
        """Background thread to monitor Hyper-V status"""
        try:
            # Initial delay to not overload startup
            time.sleep(20)
            
            while self.background_monitoring:
                # Check less frequently
                time.sleep(60)
                
                try:
                    # Only do this check on Windows
                    if platform.system() != 'Windows':
                        continue
                    
                    # Check Hyper-V status using PowerShell
                    ps_command = [
                        "powershell",
                        "-Command",
                        "(Get-Service -Name 'vmms' -ErrorAction SilentlyContinue).Status"
                    ]
                    
                    process = subprocess.Popen(
                        ps_command,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                        text=True,
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    
                    stdout, stderr = process.communicate(timeout=10)
                    
                    if process.returncode == 0 and stdout.strip():
                        status = stdout.strip()
                        if status == "Running":
                            # Log running VMs
                            try:
                                vm_command = [
                                    "powershell",
                                    "-Command",
                                    "Get-VM | Where-Object {$_.State -eq 'Running'} | Select-Object -ExpandProperty Name"
                                ]
                                
                                vm_process = subprocess.Popen(
                                    vm_command,
                                    stdout=subprocess.PIPE,
                                    stderr=subprocess.PIPE,
                                    text=True,
                                    creationflags=subprocess.CREATE_NO_WINDOW
                                )
                                
                                vm_stdout, vm_stderr = vm_process.communicate(timeout=10)
                                
                                if vm_process.returncode == 0 and vm_stdout.strip():
                                    running_vms = vm_stdout.strip().split('\n')
                                    running_vms = [vm.strip() for vm in running_vms if vm.strip()]
                                    
                                    if running_vms:
                                        self.root.after(0, lambda vms=running_vms: 
                                                    self.log(f"Running VMs: {', '.join(vms)}", "hyperv"))
                            except:
                                pass
                
                except Exception as e:
                    # Silent fail for background monitoring
                    print(f"Hyper-V monitoring error: {str(e)}")
        
        except Exception as e:
            print(f"Error in Hyper-V monitoring thread: {str(e)}")
    
    def create_log_area(self):
        """Create the log area on the right side"""
        self.log_frame = ttk.Frame(self.right_frame, style='TFrame')
        self.log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(self.log_frame, wrap=tk.WORD, font=LOG_FONT, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Create a tag for each log level
        self.log_text.tag_configure("info", foreground="black")
        self.log_text.tag_configure("warning", foreground=WARNING_FG, background=WARNING_BG)
        self.log_text.tag_configure("error", foreground=ERROR_FG, background=ERROR_BG)
        self.log_text.tag_configure("success", foreground=SUCCESS_FG, background=SUCCESS_BG)
        self.log_text.tag_configure("network", foreground="#007AFF")  # Blue for network
        self.log_text.tag_configure("hyperv", foreground="#FF9500")  # Orange for Hyper-V
        self.log_text.tag_configure("optimize", foreground="#00B359")  # Green for optimization
        self.log_text.tag_configure("storage", foreground="#8E44AD")  # Purple for storage
        self.log_text.tag_configure("system", foreground="#2C3E50")  # Dark gray for system
        self.log_text.tag_configure("background", foreground="#7F8C8D")  # Light gray for background
        
        # Create a tag for the timestamp
        self.log_text.tag_configure("timestamp", foreground="#555555")  # Gray for timestamp
        
        # Create a tag for the log level indicator
        self.log_text.tag_configure("level", font=("Segoe UI", 9, "bold"))
        
        # Create a tag for the log message
        self.log_text.tag_configure("message", font=LOG_FONT)
        
        # Create a tag for the separator line
        self.log_text.tag_configure("separator", foreground="#CCCCCC")
        
        # Create button frame below the log area
        log_button_frame = ttk.Frame(self.log_frame)
        log_button_frame.pack(fill=tk.X, pady=5)
        
        # Clear log button
        clear_log_btn = ttk.Button(log_button_frame, text="Clear Log", command=self.clear_log)
        clear_log_btn.pack(side=tk.LEFT, padx=5)
        
        # Save log button
        save_log_btn = ttk.Button(log_button_frame, text="Save Log", command=self.save_log)
        save_log_btn.pack(side=tk.LEFT, padx=5)
    
    @thread_safe
    def log(self, message, level="info"):
        """Log a message to the log area"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        formatted_message = f"[{timestamp}] [{level.upper()}] {message}\n"
        
        # Insert the message with the appropriate tag
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, formatted_message, (level, "timestamp", "level", "message"))
        self.log_text.insert(tk.END, "-" * 50 + "\n", "separator")
        self.log_text.configure(state=tk.DISABLED)
        
        # Autoscroll to the end
        self.log_text.see(tk.END)
    
    def clear_log(self):
        """Clear the log area"""
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state=tk.DISABLED)
        self.log("Log cleared", "info")
    
    def save_log(self):
        """Save the log to a file"""
        try:
            # Get the log content
            self.log_text.configure(state=tk.NORMAL)
            log_content = self.log_text.get(1.0, tk.END)
            self.log_text.configure(state=tk.DISABLED)
            
            # Ask for a file name
            file_path = filedialog.asksaveasfilename(
                initialdir=os.path.expanduser("~"),
                title="Save Log",
                defaultextension=".log",
                filetypes=[("Log Files", "*.log"), ("Text Files", "*.txt"), ("All Files", "*.*")]
            )
            
            if file_path:
                # Write the log to the file
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write(log_content)
                
                self.log(f"Log saved to {file_path}", "success")
                self.update_status(f"Log saved to {file_path}")
        except Exception as e:
            self.log(f"Error saving log: {str(e)}", "error")
            self.update_status("Error saving log")
    
    @thread_safe
    def update_status(self, message):
        """Update the status bar with a message"""
        self.status_bar.config(text=message)
    
    def log_network_info(self):
        """Log detailed network information when network tab is selected"""
        self.log("Collecting network information...", "network")
        
        # Start thread to collect detailed network info
        network_info_thread = Thread(target=self._collect_network_info)
        network_info_thread.daemon = True
        network_info_thread.start()
    
    def _collect_network_info(self):
        """Collect detailed network information in a background thread"""
        try:
            # Get network interfaces
            interfaces = psutil.net_if_addrs()
            active_interfaces = []
            
            for interface, addrs in interfaces.items():
                for addr in addrs:
                    if addr.family == socket.AF_INET and not addr.address.startswith('127.'):
                        active_interfaces.append(f"{interface}: {addr.address}")
            
            if active_interfaces:
                self.root.after(0, lambda ifs=active_interfaces: 
                            self.log(f"Active network interfaces: {', '.join(ifs)}", "network"))
            
            # Get default gateway
            try:
                gateways = psutil.net_if_stats()
                for name, stats in gateways.items():
                    if stats.isup:
                        self.root.after(0, lambda n=name, s=stats: 
                                    self.log(f"Interface {n} is up, speed: {s.speed} Mbps", "network"))
            except:
                pass
            
            # Check internet connectivity
            try:
                # Try to get external IP
                urls = ["https://api.ipify.org", "https://ifconfig.me/ip", "https://icanhazip.com"]
                
                import urllib.request
                
                for url in urls:
                    try:
                        response = urllib.request.urlopen(url, timeout=3)
                        data = response.read().decode('utf-8')
                        external_ip = data.strip()
                        self.root.after(0, lambda ip=external_ip: 
                                    self.log(f"Internet connection active, external IP: {ip}", "network"))
                        break
                    except:
                        continue
            except:
                pass
            
        except Exception as e:
            print(f"Error collecting network info: {str(e)}")
    
    def log_hyperv_info(self):
        """Log detailed Hyper-V information when Hyper-V tab is selected"""
        self.log("Collecting Hyper-V information...", "hyperv")
        
        # Start thread to collect detailed Hyper-V info
        hyperv_info_thread = Thread(target=self._collect_hyperv_info)
        hyperv_info_thread.daemon = True
        hyperv_info_thread.start()
    
    def _collect_hyperv_info(self):
        """Collect detailed Hyper-V information in a background thread"""
        try:
            # Check Hyper-V status
            ps_command = [
                "powershell",
                "-Command",
                "(Get-WindowsOptionalFeature -FeatureName Microsoft-Hyper-V-All -Online -ErrorAction SilentlyContinue).State"
            ]
            
            process = subprocess.Popen(
                ps_command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            stdout, stderr = process.communicate(timeout=10)
            
            if process.returncode == 0 and stdout.strip():
                state = stdout.strip()
                is_enabled = state == "Enabled"
                
                if is_enabled:
                    self.root.after(0, lambda: self.log("Hyper-V is enabled on this system", "hyperv"))
                    
                    # Get VM count
                    vm_command = [
                        "powershell",
                        "-Command",
                        "Get-VM -ErrorAction SilentlyContinue | Measure-Object | Select-Object -ExpandProperty Count"
                    ]
                    
                    vm_process = subprocess.Popen(
                        vm_command,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                        text=True,
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    
                    vm_stdout, vm_stderr = vm_process.communicate(timeout=10)
                    
                    if vm_process.returncode == 0 and vm_stdout.strip():
                        try:
                            vm_count = int(vm_stdout.strip())
                            self.root.after(0, lambda c=vm_count: 
                                        self.log(f"Found {c} virtual machines", "hyperv"))
                            
                            # Get running VMs
                            if vm_count > 0:
                                running_command = [
                                    "powershell",
                                    "-Command",
                                    "Get-VM | Where-Object {$_.State -eq 'Running'} | Select-Object -ExpandProperty Name"
                                ]
                                
                                running_process = subprocess.Popen(
                                    running_command,
                                    stdout=subprocess.PIPE,
                                    stderr=subprocess.PIPE,
                                    text=True,
                                    creationflags=subprocess.CREATE_NO_WINDOW
                                )
                                
                                running_stdout, running_stderr = running_process.communicate(timeout=10)
                                
                                if running_process.returncode == 0 and running_stdout.strip():
                                    running_vms = running_stdout.strip().split('\n')
                                    running_vms = [vm.strip() for vm in running_vms if vm.strip()]
                                    
                                    if running_vms:
                                        self.root.after(0, lambda vms=running_vms: 
                                                    self.log(f"Running VMs: {', '.join(vms)}", "hyperv"))
                        except Exception as e:
                            self.log(f"Error getting Hyper-V info: {str(e)}", "error")
        
        except Exception as e:
            print(f"Error collecting Hyper-V info: {str(e)}")
    
    def log_performance_info(self):
        """Log detailed performance information when Optimize tab is selected"""
        self.log("Collecting performance information...", "optimize")
        
        # Start thread to collect detailed performance info
        performance_info_thread = Thread(target=self._collect_performance_info)
        performance_info_thread.daemon = True
        performance_info_thread.start()
    
    def _collect_performance_info(self):
        """Collect detailed performance information in a background thread"""
        try:
            # Get CPU info
            cpu_info = psutil.cpu_freq()
            self.root.after(0, lambda i=cpu_info: 
                        self.log(f"CPU frequency: {i.current} MHz (min: {i.min} MHz, max: {i.max} MHz)", "optimize"))
            
            # Get memory info
            memory_info = psutil.virtual_memory()
            self.root.after(0, lambda m=memory_info: 
                        self.log(f"Memory: {m.total / (1024 ** 3):.2f} GB (used: {m.used / (1024 ** 3):.2f} GB, free: {m.free / (1024 ** 3):.2f} GB)", "optimize"))
            
            # Get disk info
            disk_info = psutil.disk_usage('/')
            self.root.after(0, lambda d=disk_info: 
                        self.log(f"Disk: {d.total / (1024 ** 3):.2f} GB (used: {d.used / (1024 ** 3):.2f} GB, free: {d.free / (1024 ** 3):.2f} GB)", "optimize"))
            
            # Get network info
            network_info = psutil.net_io_counters()
            self.root.after(0, lambda n=network_info: 
                        self.log(f"Network: sent {n.bytes_sent / (1024 ** 2):.2f} MB, received {n.bytes_recv / (1024 ** 2):.2f} MB", "optimize"))
            
            # Get boot time
            boot_time = psutil.boot_time()
            boot_time_str = datetime.fromtimestamp(boot_time).strftime("%Y-%m-%d %H:%M:%S")
            self.root.after(0, lambda bt=boot_time_str: 
                        self.log(f"System boot time: {bt}", "optimize"))
            
            # Get uptime
            uptime = time.time() - boot_time
            uptime_str = str(datetime.timedelta(seconds=int(uptime)))
            self.root.after(0, lambda ut=uptime_str: 
                        self.log(f"System uptime: {ut}", "optimize"))
            
        except Exception as e:
            print(f"Error collecting performance info: {str(e)}")
    
    def log_storage_info(self):
        """Log detailed storage information when Storage tab is selected"""
        self.log("Collecting storage information...", "storage")
        
        # Start thread to collect detailed storage info
        storage_info_thread = Thread(target=self._collect_storage_info)
        storage_info_thread.daemon = True
        storage_info_thread.start()
    
    def _collect_storage_info(self):
        """Collect detailed storage information in a background thread"""
        try:
            # Get disk usage
            disk_usage = psutil.disk_usage('/')
            self.root.after(0, lambda du=disk_usage: 
                        self.log(f"Disk usage: {du.percent}% ({du.used / (1024 ** 3):.2f} GB used, {du.free / (1024 ** 3):.2f} GB free)", "storage"))
            
            # Get disk partitions
            partitions = psutil.disk_partitions(all=False)
            for p in partitions:
                try:
                    partition_info = psutil.disk_usage(p.mountpoint)
                    self.root.after(0, lambda pi=partition_info, p=p: 
                                self.log(f"Partition {p.device} ({p.fstype}): {pi.percent}% ({pi.used / (1024 ** 3):.2f} GB used, {pi.free / (1024 ** 3):.2f} GB free)", "storage"))
                except:
                    # Skip partitions that can't be accessed
                    pass
            
        except Exception as e:
            print(f"Error collecting storage info: {str(e)}")
    
    def init_cleanup_tab(self):
        """Initialize the cleanup tab"""
        frame = ttk.Frame(self.cleanup_tab, style='Tab.TFrame', padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Title with icon
        title_frame = ttk.Frame(frame, style='Tab.TFrame')
        title_frame.grid(column=0, row=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        ttk.Label(
            title_frame, 
            text="System Cleanup", 
            font=HEADING_FONT,
            foreground=PRIMARY_COLOR
        ).pack(side=tk.LEFT)
        
        # Cleanup tools frame
        cleanup_frame = ttk.LabelFrame(frame, text="Cleanup Tools", padding=8, style='Group.TLabelframe')
        cleanup_frame.grid(column=0, row=1, sticky=tk.NSEW, pady=5, padx=5)
        
        # Cleanup tools buttons
        self.create_button(cleanup_frame, "Disk Cleanup", 
                          "Runs Windows Disk Cleanup utility",
                          lambda: self.run_disk_cleanup(), 0, 0)
        
        self.create_button(cleanup_frame, "Empty Recycle Bin", 
                          "Empties the Recycle Bin",
                          lambda: self.empty_recycle_bin(), 0, 1)
        
        self.create_button(cleanup_frame, "Clear Temp Files", 
                          "Removes temporary files",
                          lambda: self.clear_temp_files(), 0, 2)
        
        self.create_button(cleanup_frame, "Clear Windows Cache", 
                          "Clears Windows cache files",
                          lambda: self.clear_windows_cache(), 0, 3)
        
        # Description area for cleanup tab
        desc_frame = ttk.LabelFrame(frame, text="Description", padding=10, style='Group.TLabelframe')
        desc_frame.grid(column=1, row=1, rowspan=2, padx=10, pady=5, sticky=tk.NSEW)
        
        self.cleanup_desc_text = tk.Text(desc_frame, wrap=tk.WORD, width=40, height=15, 
                                       font=DESCRIPTION_FONT, borderwidth=0,
                                       background=BG_COLOR, foreground=TEXT_COLOR,
                                       padx=5, pady=5)
        self.cleanup_desc_text.pack(fill=tk.BOTH, expand=True)
        self.cleanup_desc_text.config(state=tk.DISABLED)
        
        # Configure grid weights
        frame.columnconfigure(0, weight=2)
        frame.columnconfigure(1, weight=3)
        for i in range(1, 3):
            frame.rowconfigure(i, weight=1)
        
        # Set default description
        self.update_cleanup_description("Select a cleanup option from the left to see its description.")
    
    def update_cleanup_description(self, text):
        """Update the description in the cleanup tab"""
        if hasattr(self, 'cleanup_desc_text'):
            self.cleanup_desc_text.config(state=tk.NORMAL)
            self.cleanup_desc_text.delete(1.0, tk.END)
            self.cleanup_desc_text.insert(tk.END, text)
            self.cleanup_desc_text.config(state=tk.DISABLED)
    
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
            self.system_desc_text.config(state=tk.DISABLED)
    
    def init_update_tab(self):
        """Initialize the Windows Update tab"""
        # Simplified implementation
        frame = ttk.Frame(self.update_tab, style='Tab.TFrame', padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(
            frame, 
            text="Windows Update", 
            font=HEADING_FONT,
            foreground=PRIMARY_COLOR
        ).pack(pady=(0, 10), anchor=tk.W)
        
        ttk.Label(
            frame,
            text="Windows Update features will be implemented in a future version.",
            font=NORMAL_FONT
        ).pack(pady=20)
    
    def init_storage_tab(self):
        """Initialize the Storage tab"""
        # Simplified implementation
        frame = ttk.Frame(self.storage_tab, style='Tab.TFrame', padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(
            frame, 
            text="Storage Management", 
            font=HEADING_FONT,
            foreground=PRIMARY_COLOR
        ).pack(pady=(0, 10), anchor=tk.W)
        
        ttk.Label(
            frame,
            text="Storage management features will be implemented in a future version.",
            font=NORMAL_FONT
        ).pack(pady=20)
    
    def init_network_tab(self):
        """Initialize the Network tab"""
        # Simplified implementation
        frame = ttk.Frame(self.network_tab, style='Tab.TFrame', padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(
            frame, 
            text="Network Tools", 
            font=HEADING_FONT,
            foreground=PRIMARY_COLOR
        ).pack(pady=(0, 10), anchor=tk.W)
        
        ttk.Label(
            frame,
            text="Network monitoring is active. Details will be shown in the log panel.",
            font=NORMAL_FONT
        ).pack(pady=20)
    
    def init_hyperv_tab(self):
        """Initialize the Hyper-V tab"""
        # Simplified implementation
        frame = ttk.Frame(self.hyperv_tab, style='Tab.TFrame', padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(
            frame, 
            text="Hyper-V Management", 
            font=HEADING_FONT,
            foreground=PRIMARY_COLOR
        ).pack(pady=(0, 10), anchor=tk.W)
        
        ttk.Label(
            frame,
            text="Hyper-V monitoring is active. VM information will be shown in the log panel.",
            font=NORMAL_FONT
        ).pack(pady=20)
    
    def init_optimize_tab(self):
        """Initialize the Optimize tab"""
        # Simplified implementation
        frame = ttk.Frame(self.optimize_tab, style='Tab.TFrame', padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(
            frame, 
            text="System Optimization", 
            font=HEADING_FONT,
            foreground=PRIMARY_COLOR
        ).pack(pady=(0, 10), anchor=tk.W)
        
        ttk.Label(
            frame,
            text="System performance monitoring is active. Details will be shown in the log panel.",
            font=NORMAL_FONT
        ).pack(pady=20)
    
    def create_button(self, parent, text, tooltip=None, command=None, row=0, column=0, 
                     columnspan=1, rowspan=1, padx=5, pady=5, sticky=tk.NSEW, 
                     is_primary=False, is_secondary=False):
        """Helper to create a styled button with tooltip"""
        if is_primary:
            button = ttk.Button(parent, text=text, command=command, style='Primary.TButton')
        elif is_secondary:
            button = ttk.Button(parent, text=text, command=command, style='Secondary.TButton')
        else:
            button = ttk.Button(parent, text=text, command=command)
            
        button.grid(row=row, column=column, padx=padx, pady=pady, 
                   columnspan=columnspan, rowspan=rowspan, sticky=sticky)
        
        if tooltip:
            # Create a simple tooltip
            def show_tooltip(event):
                tooltip_window = tk.Toplevel(parent)
                tooltip_window.wm_overrideredirect(True)
                tooltip_window.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
                tip_label = ttk.Label(tooltip_window, text=tooltip, background="#ffffe0", 
                                    foreground="black", relief=tk.SOLID, borderwidth=1)
                tip_label.pack()
                
                def hide_tooltip():
                    tooltip_window.destroy()
                
                button.after(2000, hide_tooltip)  # Hide after 2 seconds
                
            button.bind("<Enter>", show_tooltip)
        
        return button
    
    def enable_group_policy_editor(self):
        """Enable Group Policy Editor in Windows Home editions"""
        self.log("Starting Group Policy Editor enabler...", "info")
        self.update_status("Enabling Group Policy Editor...")
        self.update_system_description(
            "Group Policy Editor is being enabled. This process may take several minutes.\n\n" +
            "Please wait while the necessary components are installed."
        )
        messagebox.showinfo("Information", "Group Policy Editor enabler would run here.\nThis is a placeholder in the combined file.")
    
    def show_system_information(self):
        """Show detailed system information"""
        self.log("Collecting system information...", "info")
        self.update_status("Gathering system information...")
        
        # Display basic system info
        info = (
            f"Operating System: {platform.system()} {platform.version()}\n"
            f"CPU: {platform.processor()}\n"
            f"Architecture: {platform.machine()}\n"
            f"Hostname: {platform.node()}\n"
            f"Python: {sys.version.split()[0]}"
        )
        
        self.update_system_description(f"System Information:\n\n{info}")
        self.log("System information collected", "success")
    
    def run_disk_cleanup(self):
        """Run the Windows Disk Cleanup utility"""
        self.log("Running Disk Cleanup...", "info")
        self.update_status("Running Disk Cleanup...")
        self.update_cleanup_description(
            "Windows Disk Cleanup is running.\n\n" +
            "This tool helps free up space by removing temporary files, emptying the Recycle Bin, and removing system files that are no longer needed."
        )
        messagebox.showinfo("Information", "Disk Cleanup would run here.\nThis is a placeholder in the combined file.")
    
    def empty_recycle_bin(self):
        """Empty the Recycle Bin"""
        self.log("Emptying Recycle Bin...", "info")
        self.update_status("Emptying Recycle Bin...")
        self.update_cleanup_description(
            "Emptying Recycle Bin.\n\n" +
            "This will permanently delete all files in the Recycle Bin."
        )
        messagebox.showinfo("Information", "Recycle Bin emptying would run here.\nThis is a placeholder in the combined file.")
    
    def clear_temp_files(self):
        """Clear temporary files"""
        self.log("Clearing temporary files...", "info")
        self.update_status("Clearing temporary files...")
        self.update_cleanup_description(
            "Clearing temporary files.\n\n" +
            "This removes unnecessary temporary files from your system to free up disk space."
        )
        messagebox.showinfo("Information", "Temporary file cleanup would run here.\nThis is a placeholder in the combined file.")
    
    def clear_windows_cache(self):
        """Clear Windows cache files"""
        self.log("Clearing Windows cache...", "info")
        self.update_status("Clearing Windows cache...")
        self.update_cleanup_description(
            "Clearing Windows cache files.\n\n" +
            "This removes various Windows cache files to improve system performance."
        )
        messagebox.showinfo("Information", "Windows cache cleaning would run here.\nThis is a placeholder in the combined file.")

# Main entry point
def show_splash_screen():
    """Show a splash screen while the application loads"""
    splash_root = tk.Tk()
    splash_root.overrideredirect(True)  # Borderless window
    
    # Calculate position (center of screen)
    screen_width = splash_root.winfo_screenwidth()
    screen_height = splash_root.winfo_screenheight()
    
    splash_width = 400
    splash_height = 200
    
    position_x = (screen_width - splash_width) // 2
    position_y = (screen_height - splash_height) // 2
    
    splash_root.geometry(f"{splash_width}x{splash_height}+{position_x}+{position_y}")
    
    # Create a frame with border
    splash_frame = tk.Frame(splash_root, borderwidth=2, relief=tk.RIDGE, bg=BG_COLOR)
    splash_frame.pack(fill=tk.BOTH, expand=True)
    
    # Add a title
    title_label = tk.Label(splash_frame, text="Windows System Utilities", 
                         font=('Segoe UI', 18, 'bold'), bg=BG_COLOR, fg=PRIMARY_COLOR)
    title_label.pack(pady=(30,5))
    
    # Add a loading message
    loading_label = tk.Label(splash_frame, text="Loading application...", 
                           font=('Segoe UI', 10), bg=BG_COLOR, fg=TEXT_COLOR)
    loading_label.pack(pady=5)
    
    # Add version
    version_label = tk.Label(splash_frame, text="Version 1.1", 
                           font=('Segoe UI', 8), bg=BG_COLOR, fg=TEXT_COLOR)
    version_label.pack(pady=5)
    
    # Create a progress bar
    progress = ttk.Progressbar(splash_frame, orient=tk.HORIZONTAL, length=300, mode='indeterminate')
    progress.pack(pady=20)
    progress.start(10)  # Start the animation
    
    # Function to show the main window
    def show_main_window():
        splash_root.destroy()  # Close the splash screen
        
        # Create and show the main application window
        root = tk.Tk()
        app = SystemUtilities(root)
        # Start background monitoring and periodic log updates
        root.after(2000, app.update_background_logs)
        # Start the main application
        root.mainloop()
    
    # Schedule the main window to appear after 2 seconds
    splash_root.after(2000, show_main_window)
    
    # Start the splash screen
    splash_root.mainloop()


# Debug app functionality
def run_debug_app():
    """Run debug application that shows initialization steps"""
    try:
        print("Importing Tkinter...")
        import tkinter as tk
        from tkinter import ttk
        print("Tkinter imported successfully!")
        
        print("Trying to create a root window...")
        root = tk.Tk()
        print("Root window created!")
        
        print("Importing SystemUtilities...")
        from combined_system_utils import SystemUtilities
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

# Utility function to fix indentation issues
def fix_indentation(file_path):
    """Fix indentation issues in the specified file"""
    try:
        # Check if file exists and is not empty
        if not os.path.exists(file_path) or os.path.getsize(file_path) == 0:
            print(f"Error: File {file_path} doesn't exist or is empty.")
            return False
        
        # Make a backup before editing
        backup_file = f"{file_path}.bak"
        shutil.copy2(file_path, backup_file)
        print(f"Created backup at {backup_file}")
        
        # Read the file
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.readlines()
        
        # Verify we have content
        if not content:
            print("Error: No content read from file.")
            return False
        
        # Process the file
        # Find the lines with "# Additional tools frame" and fix indentation of subsequent lines
        fixed_content = []
        found_section = False
        for line in content:
            if '# Additional tools frame' in line:
                found_section = True
                fixed_content.append(line)
            elif found_section and line.startswith('      '):  # Lines with 6 spaces
                # Fix indentation from 6 to 4 spaces
                fixed_content.append('    ' + line[6:])
                if 'is_primary=True' in line:  # End of the section
                    found_section = False
            else:
                fixed_content.append(line)
        
        # Verify we still have content after processing
        if not fixed_content:
            print("Error: No content after processing. Restoring from backup.")
            shutil.copy2(backup_file, file_path)
            return False
        
        # Write back to file
        with open(file_path, 'w', encoding='utf-8') as f:
            f.writelines(fixed_content)
        
        print("Indentation fixed successfully!")
        return True
        
    except Exception as e:
        print(f"Error: {str(e)}")
        # Restore from backup if there was an error
        if os.path.exists(backup_file):
            print("Restoring from backup...")
            shutil.copy2(backup_file, file_path)
            print("Restore completed.")
        return False

# Function to apply system utilities fixes
def fix_system_utilities(file_path):
    """Apply fixes to the system utilities file"""
    try:
        with open(file_path, "r") as f:
            content = f.read()
        
        # Update version number
        content = content.replace("text=\"Version 1.0\"", "text=\"Version 1.1\"")
        
        # Add background monitoring code
        content = content.replace(
            "app = SystemUtilities(root)", 
            "app = SystemUtilities(root)\n        # Start background monitoring and periodic log updates\n        root.after(2000, app.update_background_logs)"
        )
        
        # Write back to file
        with open(file_path, "w") as f:
            f.write(content)
        
        print("File updated successfully!")
        return True
    except Exception as e:
        print(f"Error fixing system utilities: {e}")
        return False

# This is the main entry point
if __name__ == "__main__":
    # If running directly, show the splash screen
    show_splash_screen() 