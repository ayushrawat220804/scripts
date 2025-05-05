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
from datetime import datetime, timedelta
import time
import csv
import socket
import functools
import traceback
import concurrent.futures
import fnmatch
import gc
import weakref
import queue
import matplotlib
matplotlib.use('TkAgg')  # Use TkAgg backend
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

try:
    import winshell
except ImportError:
    pass  # Handle later in the code

# Constants for UI
VERSION = "1.2"
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

# Message queue for thread-safe updates
ui_queue = queue.Queue()

# Improved global exception handler to prevent application crashes
def global_exception_handler(exc_type, exc_value, exc_traceback):
    """Global exception handler to prevent crashes"""
    # Format the exception
    error_msg = ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    # Log to console
    print(f"Unhandled exception: {error_msg}")
    # Show error message in dialog if possible, but use queue to avoid threading issues
    try:
        # Put the error in the queue for the main thread to handle
        ui_queue.put(("error_dialog", str(exc_value)))
    except:
        # If queue fails, just print to console
        print("Could not queue error dialog")
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
            if hasattr(self, 'log'):
                self.log(f"Error in {func.__name__}: {str(e)}", 'error')
            # Log the traceback for debugging
            if hasattr(self, 'show_verbose_logs') and self.show_verbose_logs:
                import traceback
                self.log(f"Traceback: {traceback.format_exc()}", 'error')
            # Show error message to user through queue to avoid threading issues
            if hasattr(self, 'update_status'):
                self.update_status(f"Error in {func.__name__}")
            # Push error dialog to queue
            ui_queue.put(("error_dialog", f"An unexpected error occurred in {func.__name__}:\n\n{str(e)}"))
            print(f"Error in {func.__name__}: {str(e)}")
            return None
    return wrapper

def thread_safe(func):
    """Decorator for ensuring thread-safe UI updates"""
    @functools.wraps(func)
    def wrapper(self, *args, **kwargs):
        if threading.current_thread() is not threading.main_thread():
            # We're not in the main thread, so we need to use the queue
            # Create a custom callback to run the function in the main thread
            result_container = []
            completion_event = threading.Event()
            
            def main_thread_callback():
                try:
                    with ui_lock:
                        result = func(self, *args, **kwargs)
                        result_container.append(result)
                except Exception as e:
                    print(f"Error in thread_safe callback: {str(e)}")
                finally:
                    completion_event.set()
            
            # Put the callback in the queue
            ui_queue.put(("callback", main_thread_callback))
            
            # Wait for completion with timeout to prevent deadlocks
            completed = completion_event.wait(timeout=5.0)
            if not completed:
                print(f"Warning: Timed out waiting for {func.__name__} to complete in main thread")
                return None
                
            # Return the result if available
            return result_container[0] if result_container else None
        else:
            # Already in main thread, just use the lock
            with ui_lock:
                return func(self, *args, **kwargs)
    return wrapper

class SystemUtilities:
    """Main application class for Windows System Utilities"""
    
    def __init__(self, root):
        """Initialize the application"""
        self.root = root
        self.root.title("Windows System Utilities")
        
        # Make window open in fullscreen by default
        self.root.state('zoomed')  # For Windows, this maximizes the window
        
        # Get screen dimensions for responsive design
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # Set minimum size for the window
        self.root.minsize(1024, 768)
        
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
        
        # Right side - Dashboard and real-time info
        self.right_frame = ttk.Frame(self.paned, style='TFrame')
        
        # Add the frames to the paned window
        self.paned.add(self.left_frame, weight=3)
        self.paned.add(self.right_frame, weight=2)
        
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
        
        # Set up the right side with dashboard widgets
        self.create_dashboard()
        
        # Set up log area at the bottom
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
        self.executor = concurrent.futures.ThreadPoolExecutor(max_workers=3)
        
        # Set background monitoring parameters
        self.background_monitoring = True
        self.show_verbose_logs = True
        
        # Initialize data for real-time monitoring
        self.monitor_data = {
            'times': [],
            'cpu': [],
            'memory': [],
            'disk': [],
            'network': []
        }
        self.max_data_points = 60  # Max data points to show in charts
        
        # Keep track of the figures to properly close them later
        self.figures = []
        
        # Track active threads for proper cleanup
        self.active_threads = weakref.WeakSet()
        
        # Process UI queue periodically
        self.process_ui_queue()
        
        # Show a welcome message
        self.log("System Utilities initialized successfully", "info")
        self.update_status("Ready")
        
        # Start background monitoring
        self.start_background_monitoring()
        
        # Periodically collect garbage to prevent memory leaks
        def garbage_collect():
            gc.collect()
            if hasattr(self, 'root') and self.root.winfo_exists():
                self.root.after(300000, garbage_collect)  # Every 5 minutes
        
        # Start the garbage collection timer
        self.root.after(300000, garbage_collect)
        
        # Setup proper cleanup on window close
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # Set reduced update frequency for better performance
        self.update_frequency = 2000  # milliseconds
    
    def process_ui_queue(self):
        """Process any pending UI updates from the queue"""
        try:
            # Process all available messages
            for _ in range(100):  # Process up to 100 messages at once to prevent blocking
                try:
                    message = ui_queue.get_nowait()
                    if message:
                        msg_type, msg_data = message
                        
                        if msg_type == "error_dialog":
                            # Show error dialog
                            messagebox.showerror("Error", msg_data)
                        elif msg_type == "info_dialog":
                            # Show info dialog
                            messagebox.showinfo("Information", msg_data)
                        elif msg_type == "warning_dialog":
                            # Show warning dialog
                            messagebox.showwarning("Warning", msg_data)
                        elif msg_type == "status_update":
                            # Update status bar
                            self.status_bar.config(text=msg_data)
                        elif msg_type == "log":
                            # Add to log
                            level, text = msg_data
                            self._direct_log(text, level)
                        elif msg_type == "callback":
                            # Execute callback function
                            callback_func = msg_data
                            callback_func()
                            
                    ui_queue.task_done()
                except queue.Empty:
                    break
                except Exception as e:
                    print(f"Error processing UI queue message: {str(e)}")
        finally:
            # Schedule the next check
            if hasattr(self, 'root') and self.root.winfo_exists():
                self.root.after(100, self.process_ui_queue)
    
    def on_close(self):
        """Handle window close event"""
        # Stop background monitoring
        self.background_monitoring = False
        
        # Close matplotlib figures to prevent memory leaks
        import matplotlib.pyplot as plt
        plt.close('all')
        
        # Destroy the window
        self.root.destroy()

    def create_dashboard(self):
        """Create the dashboard on the right side with real-time system information"""
        # Create notebook for dashboard tabs
        self.dashboard_tabs = ttk.Notebook(self.right_frame)
        self.dashboard_tabs.pack(fill=tk.BOTH, expand=True)
        
        # Create dashboard tabs
        self.overview_tab = ttk.Frame(self.dashboard_tabs, style='Tab.TFrame')
        self.performance_tab = ttk.Frame(self.dashboard_tabs, style='Tab.TFrame')
        self.log_tab = ttk.Frame(self.dashboard_tabs, style='Tab.TFrame')
        
        # Add tabs to notebook
        self.dashboard_tabs.add(self.overview_tab, text="Overview")
        self.dashboard_tabs.add(self.performance_tab, text="Performance")
        self.dashboard_tabs.add(self.log_tab, text="Logs")
        
        # Create the overview dashboard
        self.create_overview_dashboard()
        
        # Create the performance dashboard
        self.create_performance_dashboard()
        
        # Add log text area to log tab
        self.log_frame = ttk.Frame(self.log_tab, style='TFrame')
        self.log_frame.pack(fill=tk.BOTH, expand=True)
        
        # Current selected dashboard tab (to avoid needless updates)
        self.current_dashboard_tab = "Overview"
        self.dashboard_tabs.bind("<<NotebookTabChanged>>", self.on_dashboard_tab_changed)
    
    def on_dashboard_tab_changed(self, event):
        """Handle dashboard tab change event"""
        tab_id = self.dashboard_tabs.select()
        tab_name = self.dashboard_tabs.tab(tab_id, "text")
        self.current_dashboard_tab = tab_name
        # Force refresh the selected tab
        self.update_dashboard(force=True)
    
    def create_overview_dashboard(self):
        """Create the overview dashboard with system information widgets"""
        frame = ttk.Frame(self.overview_tab, style='Tab.TFrame', padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        ttk.Label(
            frame, 
            text="System Overview", 
            font=HEADING_FONT,
            foreground=PRIMARY_COLOR
        ).pack(pady=(0, 10), anchor=tk.W)
        
        # Top frame with system info
        top_frame = ttk.Frame(frame, style='TFrame')
        top_frame.pack(fill=tk.X, expand=False, pady=5)
        
        # System info frame
        system_frame = ttk.LabelFrame(top_frame, text="System Information", padding=8, style='Group.TLabelframe')
        system_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        
        # System info grid
        self.os_label = ttk.Label(system_frame, text="OS: Loading...", font=NORMAL_FONT)
        self.os_label.grid(row=0, column=0, sticky=tk.W, pady=2)
        
        self.cpu_label = ttk.Label(system_frame, text="CPU: Loading...", font=NORMAL_FONT)
        self.cpu_label.grid(row=1, column=0, sticky=tk.W, pady=2)
        
        self.memory_label = ttk.Label(system_frame, text="RAM: Loading...", font=NORMAL_FONT)
        self.memory_label.grid(row=2, column=0, sticky=tk.W, pady=2)
        
        self.uptime_label = ttk.Label(system_frame, text="Uptime: Loading...", font=NORMAL_FONT)
        self.uptime_label.grid(row=3, column=0, sticky=tk.W, pady=2)
        
        # Resource usage frame
        usage_frame = ttk.LabelFrame(top_frame, text="Resource Usage", padding=8, style='Group.TLabelframe')
        usage_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5)
        
        # CPU usage
        ttk.Label(usage_frame, text="CPU Usage:", font=NORMAL_FONT).grid(row=0, column=0, sticky=tk.W)
        self.cpu_progressbar = ttk.Progressbar(usage_frame, orient=tk.HORIZONTAL, length=200, mode='determinate')
        self.cpu_progressbar.grid(row=0, column=1, padx=5, pady=2, sticky=tk.EW)
        self.cpu_usage_label = ttk.Label(usage_frame, text="0%", font=NORMAL_FONT, width=5)
        self.cpu_usage_label.grid(row=0, column=2, sticky=tk.W)
        
        # Memory usage
        ttk.Label(usage_frame, text="Memory Usage:", font=NORMAL_FONT).grid(row=1, column=0, sticky=tk.W)
        self.memory_progressbar = ttk.Progressbar(usage_frame, orient=tk.HORIZONTAL, length=200, mode='determinate')
        self.memory_progressbar.grid(row=1, column=1, padx=5, pady=2, sticky=tk.EW)
        self.memory_usage_label = ttk.Label(usage_frame, text="0%", font=NORMAL_FONT, width=5)
        self.memory_usage_label.grid(row=1, column=2, sticky=tk.W)
        
        # Disk usage
        ttk.Label(usage_frame, text="Disk Usage:", font=NORMAL_FONT).grid(row=2, column=0, sticky=tk.W)
        self.disk_progressbar = ttk.Progressbar(usage_frame, orient=tk.HORIZONTAL, length=200, mode='determinate')
        self.disk_progressbar.grid(row=2, column=1, padx=5, pady=2, sticky=tk.EW)
        self.disk_usage_label = ttk.Label(usage_frame, text="0%", font=NORMAL_FONT, width=5)
        self.disk_usage_label.grid(row=2, column=2, sticky=tk.W)
        
        # Refresh button for system info
        refresh_btn = ttk.Button(frame, text="Refresh Information", command=self.refresh_system_info)
        refresh_btn.pack(pady=10)
        
        # Bottom frame with network and process info
        bottom_frame = ttk.Frame(frame, style='TFrame')
        bottom_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Network info frame
        network_frame = ttk.LabelFrame(bottom_frame, text="Network Activity", padding=8, style='Group.TLabelframe')
        network_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        
        # Network activity display (text widget with auto-update)
        self.network_text = tk.Text(network_frame, wrap=tk.WORD, height=8, width=30, 
                                  font=NORMAL_FONT, bg=LOG_BG)
        self.network_text.pack(fill=tk.BOTH, expand=True)
        self.network_text.config(state=tk.DISABLED)
        
        # Process info frame
        process_frame = ttk.LabelFrame(bottom_frame, text="Active Processes", padding=8, style='Group.TLabelframe')
        process_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5)
        
        # Create a frame for the treeview and scrollbar
        tree_frame = ttk.Frame(process_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # Process list (treeview)
        self.process_tree = ttk.Treeview(tree_frame, columns=("pid", "cpu", "memory"), 
                                        show="headings", height=6)
        self.process_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Add scrollbar to treeview
        process_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.process_tree.yview)
        process_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.process_tree.configure(yscrollcommand=process_scrollbar.set)
        
        # Configure treeview columns
        self.process_tree.heading("pid", text="PID")
        self.process_tree.heading("cpu", text="CPU %")
        self.process_tree.heading("memory", text="Memory")
        
        self.process_tree.column("pid", width=50)
        self.process_tree.column("cpu", width=50)
        self.process_tree.column("memory", width=80)
    
    def create_performance_dashboard(self):
        """Create the performance dashboard with charts"""
        frame = ttk.Frame(self.performance_tab, style='Tab.TFrame', padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        ttk.Label(
            frame, 
            text="Performance Monitor", 
            font=HEADING_FONT,
            foreground=PRIMARY_COLOR
        ).pack(pady=(0, 10), anchor=tk.W)
        
        # Create frame for charts
        charts_frame = ttk.Frame(frame, style='TFrame')
        charts_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create CPU usage chart
        cpu_frame = ttk.LabelFrame(charts_frame, text="CPU Usage", padding=8, style='Group.TLabelframe')
        cpu_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.cpu_figure = Figure(figsize=(5, 2), dpi=100)
        self.cpu_plot = self.cpu_figure.add_subplot(111)
        self.cpu_canvas = FigureCanvasTkAgg(self.cpu_figure, cpu_frame)
        self.cpu_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # Create memory usage chart
        memory_frame = ttk.LabelFrame(charts_frame, text="Memory Usage", padding=8, style='Group.TLabelframe')
        memory_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.memory_figure = Figure(figsize=(5, 2), dpi=100)
        self.memory_plot = self.memory_figure.add_subplot(111)
        self.memory_canvas = FigureCanvasTkAgg(self.memory_figure, memory_frame)
        self.memory_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # Create disk and network chart frame
        bottom_frame = ttk.Frame(charts_frame, style='TFrame')
        bottom_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Create disk activity chart
        disk_frame = ttk.LabelFrame(bottom_frame, text="Disk Activity", padding=8, style='Group.TLabelframe')
        disk_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0,5))
        
        self.disk_figure = Figure(figsize=(3, 2), dpi=100)
        self.disk_plot = self.disk_figure.add_subplot(111)
        self.disk_canvas = FigureCanvasTkAgg(self.disk_figure, disk_frame)
        self.disk_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # Create network activity chart
        network_frame = ttk.LabelFrame(bottom_frame, text="Network Activity", padding=8, style='Group.TLabelframe')
        network_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5,0))
        
        self.network_figure = Figure(figsize=(3, 2), dpi=100)
        self.network_plot = self.network_figure.add_subplot(111)
        self.network_canvas = FigureCanvasTkAgg(self.network_figure, network_frame)
        self.network_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # Auto-refresh checkbox
        self.auto_refresh_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            frame, 
            text="Auto-refresh (every 2s)",
            variable=self.auto_refresh_var
        ).pack(pady=5)
        
        # Manual refresh button
        refresh_btn = ttk.Button(frame, text="Refresh Now", command=lambda: self.update_dashboard(force=True))
        refresh_btn.pack(pady=5)
    
    def create_log_area(self):
        """Create the log area in the log tab"""
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
        """Log a message to the log area (thread-safe)"""
        # If not on main thread, use queue
        if threading.current_thread() is not threading.main_thread():
            ui_queue.put(("log", (level, message)))
            return
        
        # Direct logging if on main thread
        self._direct_log(message, level)
    
    def _direct_log(self, message, level="info"):
        """Direct logging implementation (should only be called from main thread)"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        formatted_message = f"[{timestamp}] [{level.upper()}] {message}\n"
        
        # Insert the message with the appropriate tag
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, formatted_message, (level, "timestamp", "level", "message"))
        self.log_text.insert(tk.END, "-" * 50 + "\n", "separator")
        self.log_text.configure(state=tk.DISABLED)
        
        # Autoscroll to the end
        self.log_text.see(tk.END)
        
        # If too many lines, trim the log
        if self.log_text.index('end-1c linestart') > '5000.0':
            self.log_text.configure(state=tk.NORMAL)
            self.log_text.delete('1.0', '1000.0')
            self.log_text.configure(state=tk.DISABLED)
    
    def clear_log(self):
        """Clear the log area"""
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
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
        """Update the status bar with a message (thread-safe)"""
        if threading.current_thread() is not threading.main_thread():
            ui_queue.put(("status_update", message))
            return
        self.status_bar.config(text=message)
    
    def start_background_monitoring(self):
        """Start background monitoring threads"""
        self.log("Starting background system monitoring...", "background")
        
        # Start system metrics collection in a separate thread
        def system_metrics_collector():
            try:
                while self.background_monitoring and hasattr(self, 'root') and self.root.winfo_exists():
                    try:
                        # Collect system metrics
                        self._collect_system_metrics()
                        # Sleep between collections
                        time.sleep(1)
                    except Exception as e:
                        print(f"Error in system metrics collector: {str(e)}")
                        time.sleep(2)  # Sleep a bit longer on error
            except Exception as e:
                print(f"System metrics collector thread died: {str(e)}")
        
        metrics_thread = Thread(target=system_metrics_collector)
        metrics_thread.daemon = True
        self.active_threads.add(metrics_thread)
        metrics_thread.start()
        
        # Start network monitoring thread (less frequent)
        network_thread = Thread(target=self._monitor_network_activity)
        network_thread.daemon = True
        self.active_threads.add(network_thread)
        network_thread.start()
        
        # Start Hyper-V monitoring thread (if applicable)
        hyperv_thread = Thread(target=self._monitor_hyperv_status)
        hyperv_thread.daemon = True
        self.active_threads.add(hyperv_thread)
        hyperv_thread.start()
        
        # Schedule first dashboard update
        self.root.after(2000, lambda: self.update_dashboard(first_time=True))
    
    def update_dashboard(self, first_time=False, force=False):
        """Update the dashboard with current system information"""
        if not hasattr(self, 'root') or not self.root.winfo_exists():
            return
            
        try:
            # Get CPU and memory usage
            cpu_percent = psutil.cpu_percent(interval=0)  # Non-blocking
            memory = psutil.virtual_memory()
            memory_percent = memory.percent
            
            # Update CPU usage display
            self.cpu_progressbar['value'] = cpu_percent
            self.cpu_usage_label.config(text=f"{cpu_percent}%")
            
            # Update memory usage display
            self.memory_progressbar['value'] = memory_percent
            self.memory_usage_label.config(text=f"{memory_percent}%")
            
            # Update disk usage less frequently (every 10 seconds)
            current_time = time.time()
            if not hasattr(self, '_last_disk_check') or current_time - self._last_disk_check > 10:
                try:
                    disk = psutil.disk_usage('/')
                    disk_percent = disk.percent
                    self.disk_progressbar['value'] = disk_percent
                    self.disk_usage_label.config(text=f"{disk_percent}%")
                    self._last_disk_check = current_time
                except Exception as e:
                    print(f"Error getting disk usage: {e}")
            
            # Update uptime less frequently (every minute)
            if not hasattr(self, '_last_uptime_check') or current_time - self._last_uptime_check > 60:
                try:
                    boot_time = psutil.boot_time()
                    uptime = current_time - boot_time
                    uptime_str = str(timedelta(seconds=int(uptime)))
                    self.uptime_label.config(text=f"Uptime: {uptime_str}")
                    self._last_uptime_check = current_time
                except Exception as e:
                    print(f"Error updating uptime: {e}")
            
            # Add data to monitor arrays
            current_time_str = datetime.now().strftime('%H:%M:%S')
            self.monitor_data['times'].append(current_time_str)
            self.monitor_data['cpu'].append(cpu_percent)
            self.monitor_data['memory'].append(memory_percent)
            
            # Keep arrays at a reasonable size
            if len(self.monitor_data['times']) > self.max_data_points:
                self.monitor_data['times'] = self.monitor_data['times'][-self.max_data_points:]
                self.monitor_data['cpu'] = self.monitor_data['cpu'][-self.max_data_points:]
                self.monitor_data['memory'] = self.monitor_data['memory'][-self.max_data_points:]
            
            # Only update these components conditionally to avoid freezing
            current_dashboard_tab = self.dashboard_tabs.tab(self.dashboard_tabs.select(), "text")
            
            # Update process list only when on Overview tab
            if current_dashboard_tab == "Overview" or force:
                try:
                    self.update_process_list()
                    self.update_network_info()
                except Exception as e:
                    print(f"Error updating overview components: {e}")
            
            # Update charts only when on Performance tab
            if current_dashboard_tab == "Performance" or force:
                try:
                    self.update_performance_charts()
                except Exception as e:
                    print(f"Error updating performance charts: {e}")
            
        except Exception as e:
            print(f"Dashboard update error: {e}")
            self.log(f"Error updating dashboard: {e}", "error")
        
        # Schedule next update with longer interval (2 seconds instead of 1)
        if hasattr(self, 'auto_refresh_var') and self.auto_refresh_var.get():
            self.root.after(2000, self.update_dashboard)
    
    def update_performance_charts(self):
        """Update all performance charts with new data"""
        try:
            # Only update if we have data
            if not self.monitor_data['times'] or len(self.monitor_data['times']) < 2:
                return
            
            # Update CPU chart
            self.cpu_plot.clear()
            self.cpu_plot.plot(range(len(self.monitor_data['cpu'])), 
                              self.monitor_data['cpu'], 
                              color=PRIMARY_COLOR)
            self.cpu_plot.set_ylim(0, 100)
            self.cpu_plot.set_ylabel('Percent')
            self.cpu_figure.tight_layout()
            try:
                # Draw idle to prevent blocking
                self.cpu_canvas.draw_idle()
            except:
                pass
            
            # Update Memory chart
            self.memory_plot.clear()
            self.memory_plot.plot(range(len(self.monitor_data['memory'])), 
                                 self.monitor_data['memory'], 
                                 color=SECONDARY_COLOR)
            self.memory_plot.set_ylim(0, 100)
            self.memory_plot.set_ylabel('Percent')
            self.memory_figure.tight_layout()
            try:
                self.memory_canvas.draw_idle()
            except:
                pass
            
            # Update other charts only occasionally to reduce CPU usage
            current_time = time.time()
            if not hasattr(self, '_last_chart_update') or current_time - self._last_chart_update > 5:
                # Update Disk and Network charts
                try:
                    # Disk chart
                    self.disk_plot.clear()
                    self.disk_plot.plot(range(len(self.monitor_data['disk'])), 
                                      self.monitor_data['disk'], 
                                      color=ACCENT_COLOR)
                    self.disk_plot.set_ylim(0, 100)
                    self.disk_figure.tight_layout()
                    self.disk_canvas.draw_idle()
                    
                    # Network chart
                    self.network_plot.clear()
                    self.network_plot.plot(range(len(self.monitor_data['network'])), 
                                         self.monitor_data['network'], 
                                         color="#9B59B6")
                    self.network_plot.set_ylim(0, 100)
                    self.network_figure.tight_layout()
                    self.network_canvas.draw_idle()
                    
                    self._last_chart_update = current_time
                except:
                    pass
            
        except Exception as e:
            print(f"Error updating charts: {e}")
    
    def update_process_list(self):
        """Update the process list in the system overview"""
        try:
            # Clear the current process list
            for item in self.process_tree.get_children():
                self.process_tree.delete(item)
            
            # Get top processes by CPU usage
            processes = []
            for proc in psutil.process_iter(['pid', 'name', 'cpu_percent', 'memory_info']):
                try:
                    info = proc.info
                    if info['cpu_percent'] > 0:  # Only track active processes
                        processes.append((
                            proc.name(),
                            info['pid'],
                            info['cpu_percent'],
                            info['memory_info'].rss / (1024 * 1024)  # MB
                        ))
                except:
                    continue
            
            # Sort and display only top processes
            processes.sort(key=lambda x: x[2], reverse=True)
            for name, pid, cpu, memory in processes[:10]:  # Show top 10
                self.process_tree.insert('', 'end', values=(
                    f"{pid}",
                    f"{cpu:.1f}",
                    f"{memory:.1f} MB"
                ), text=name)
        
        except Exception as e:
            print(f"Error updating process list: {e}")
    
    def _collect_system_metrics(self):
        """Collect system metrics in background thread"""
        try:
            # Get CPU and memory usage
            cpu_percent = psutil.cpu_percent()
            memory = psutil.virtual_memory()
            memory_percent = memory.percent
            
            # Add data to monitor data arrays with thread safety
            current_time = datetime.now().strftime('%H:%M:%S')
            
            with ui_lock:
                self.monitor_data['times'].append(current_time)
                self.monitor_data['cpu'].append(cpu_percent)
                self.monitor_data['memory'].append(memory_percent)
                
                # Keep only the last max_data_points in each array
                if len(self.monitor_data['times']) > self.max_data_points:
                    self.monitor_data['times'] = self.monitor_data['times'][-self.max_data_points:]
                    self.monitor_data['cpu'] = self.monitor_data['cpu'][-self.max_data_points:]
                    self.monitor_data['memory'] = self.monitor_data['memory'][-self.max_data_points:]
            
            # Get disk I/O for disk activity chart
            try:
                disk_io = psutil.disk_io_counters()
                if hasattr(self, 'prev_disk_io'):
                    read_mb = (disk_io.read_bytes - self.prev_disk_io.read_bytes) / (1024 * 1024)
                    write_mb = (disk_io.write_bytes - self.prev_disk_io.write_bytes) / (1024 * 1024)
                    
                    # Cap at 100 for chart display
                    disk_activity = min(100, (read_mb + write_mb) * 5)
                    
                    with ui_lock:
                        self.monitor_data['disk'].append(disk_activity)
                        
                        # Keep only max_data_points
                        if len(self.monitor_data['disk']) > self.max_data_points:
                            self.monitor_data['disk'] = self.monitor_data['disk'][-self.max_data_points:]
                else:
                    # First-time initialization
                    with ui_lock:
                        self.monitor_data['disk'] = [0] * min(len(self.monitor_data['times']), self.max_data_points)
                
                self.prev_disk_io = disk_io
            except:
                # If disk I/O monitoring fails, add a zero
                with ui_lock:
                    self.monitor_data['disk'].append(0)
                    if len(self.monitor_data['disk']) > self.max_data_points:
                        self.monitor_data['disk'] = self.monitor_data['disk'][-self.max_data_points:]
            
            # Get network I/O for network activity chart
            try:
                net_io = psutil.net_io_counters()
                if hasattr(self, 'prev_net_io'):
                    sent_mb = (net_io.bytes_sent - self.prev_net_io.bytes_sent) / (1024 * 1024)
                    recv_mb = (net_io.bytes_recv - self.prev_net_io.bytes_recv) / (1024 * 1024)
                    
                    # Cap at 100 for chart display
                    network_activity = min(100, (sent_mb + recv_mb) * 10)
                    
                    with ui_lock:
                        self.monitor_data['network'].append(network_activity)
                        
                        # Keep only max_data_points
                        if len(self.monitor_data['network']) > self.max_data_points:
                            self.monitor_data['network'] = self.monitor_data['network'][-self.max_data_points:]
                else:
                    # First-time initialization
                    with ui_lock:
                        self.monitor_data['network'] = [0] * min(len(self.monitor_data['times']), self.max_data_points)
                
                self.prev_net_io = net_io
            except:
                # If network I/O monitoring fails, add a zero
                with ui_lock:
                    self.monitor_data['network'].append(0)
                    if len(self.monitor_data['network']) > self.max_data_points:
                        self.monitor_data['network'] = self.monitor_data['network'][-self.max_data_points:]
        
        except Exception as e:
            print(f"Error collecting system metrics: {str(e)}")
    
    @thread_safe
    def refresh_system_indicators(self):
        """Update system indicators only (not full system info)"""
        try:
            # Update resource usage bars only
            cpu_percent = psutil.cpu_percent(interval=0)  # Non-blocking
            self.cpu_progressbar['value'] = cpu_percent
            self.cpu_usage_label.config(text=f"{cpu_percent}%")
            
            memory = psutil.virtual_memory()
            memory_percent = memory.percent
            self.memory_progressbar['value'] = memory_percent
            self.memory_usage_label.config(text=f"{memory_percent}%")
            
            # Only update disk usage occasionally (expensive operation)
            if not hasattr(self, '_last_disk_check') or time.time() - self._last_disk_check > 10:
                try:
                    disk = psutil.disk_usage('/')
                    disk_percent = disk.percent
                    self.disk_progressbar['value'] = disk_percent
                    self.disk_usage_label.config(text=f"{disk_percent}%")
                    self._last_disk_check = time.time()
                except:
                    pass
            
            # Update uptime occasionally
            if not hasattr(self, '_last_uptime_check') or time.time() - self._last_uptime_check > 60:
                try:
                    boot_time = psutil.boot_time()
                    uptime = time.time() - boot_time
                    uptime_str = str(timedelta(seconds=int(uptime)))
                    self.uptime_label.config(text=f"Uptime: {uptime_str}")
                    self._last_uptime_check = time.time()
                except:
                    pass
                    
        except Exception as e:
            print(f"Error refreshing system indicators: {str(e)}")
    
    @thread_safe
    def refresh_system_info(self):
        """Update all system information displays (full refresh)"""
        try:
            # Update system info labels (expensive, do less frequently)
            try:
                self.os_label.config(text=f"OS: {platform.system()} {platform.version()}")
                self.cpu_label.config(text=f"CPU: {platform.processor()}")
                
                # Update memory info
                memory = psutil.virtual_memory()
                total_gb = memory.total / (1024**3)
                self.memory_label.config(text=f"RAM: {total_gb:.2f} GB Total")
            except Exception as e:
                print(f"Error updating system labels: {str(e)}")
            
            # Update resource usage indicators
            self.refresh_system_indicators()
            
        except Exception as e:
            self.log(f"Error refreshing system info: {str(e)}", "error")
            print(f"System info refresh error: {str(e)}")
    
    @thread_safe
    def update_network_info(self):
        """Update network information in the network textbox"""
        try:
            # Get network interfaces
            interfaces = psutil.net_if_addrs()
            
            # Only update if widget exists
            if not hasattr(self, 'network_text') or not self.network_text.winfo_exists():
                return
                
            # Clear previous content
            self.network_text.config(state=tk.NORMAL)
            self.network_text.delete(1.0, tk.END)
            
            # Get network I/O
            net_io = psutil.net_io_counters()
            self.network_text.insert(tk.END, f"Total Sent: {net_io.bytes_sent / (1024**2):.2f} MB\n")
            self.network_text.insert(tk.END, f"Total Received: {net_io.bytes_recv / (1024**2):.2f} MB\n\n")
            
            # Active interfaces (only show IPv4)
            self.network_text.insert(tk.END, "Active Interfaces:\n")
            for interface_name, interface_addresses in interfaces.items():
                for addr in interface_addresses:
                    if str(addr.family) == 'AddressFamily.AF_INET' and not addr.address.startswith('127.'):
                        self.network_text.insert(tk.END, f"{interface_name}: {addr.address}\n")
            
            # Connection count (expensive, do less frequently)
            if not hasattr(self, '_last_connections_check') or time.time() - self._last_connections_check > 10:
                try:
                    # Only count established connections to reduce resource usage
                    connections = len([c for c in psutil.net_connections(kind='inet') 
                                     if c.status == 'ESTABLISHED'])
                    self.network_text.insert(tk.END, f"\nActive connections: {connections}\n")
                    self._last_connections_check = time.time()
                except:
                    pass
                
            self.network_text.config(state=tk.DISABLED)
            
        except Exception as e:
            print(f"Error updating network info: {str(e)}")
    
    def _monitor_network_activity(self):
        """Background thread to monitor network activity"""
        try:
            # Initialize network stats
            prev_net_io = psutil.net_io_counters()
            last_check_time = time.time()
            
            while self.background_monitoring:
                time.sleep(10)  # Check every 10 seconds
                
                # Exit if application is closing
                if not hasattr(self, 'root') or not self.root.winfo_exists():
                    return
                
                try:
                    # Get current network stats
                    net_io = psutil.net_io_counters()
                    current_time = time.time()
                    time_diff = current_time - last_check_time
                    
                    # Calculate data transferred since last check
                    sent_mb = (net_io.bytes_sent - prev_net_io.bytes_sent) / (1024 * 1024)
                    recv_mb = (net_io.bytes_recv - prev_net_io.bytes_recv) / (1024 * 1024)
                    
                    # Calculate rate per second
                    sent_rate = sent_mb / time_diff
                    recv_rate = recv_mb / time_diff
                    
                    # Log significant network activity to avoid flooding logs
                    if (sent_rate > 1 or recv_rate > 1) and self.show_verbose_logs:  # Only log if more than 1MB/s
                        # Use UI queue instead of direct UI access
                        ui_queue.put(("log", ("network", f"Network activity: {sent_rate:.2f} MB/s upload, {recv_rate:.2f} MB/s download")))
                    
                    # Update previous stats
                    prev_net_io = net_io
                    last_check_time = current_time
                    
                except Exception as e:
                    print(f"Network monitoring error: {str(e)}")
        
        except Exception as e:
            print(f"Error in network monitoring thread: {str(e)}")
    
    def _monitor_system_performance(self):
        """Background thread to monitor system performance"""
        try:
            while self.background_monitoring:
                time.sleep(30)  # Check less frequently - every 30 seconds
                
                # Exit if application is closing
                if not hasattr(self, 'root') or not self.root.winfo_exists():
                    return
                
                try:
                    # Get system performance metrics
                    cpu_percent = psutil.cpu_percent(interval=1)
                    memory = psutil.virtual_memory()
                    
                    # Log when system is under high load
                    if cpu_percent > 85:  # Only log higher thresholds
                        ui_queue.put(("log", ("system", f"High CPU usage: {cpu_percent}%")))
                    
                    if memory.percent > 90:  # Only log higher thresholds
                        ui_queue.put(("log", ("system", f"High memory usage: {memory.percent}%")))
                    
                except Exception as e:
                    print(f"Performance monitoring error: {str(e)}")
        
        except Exception as e:
            print(f"Error in performance monitoring thread: {str(e)}")
    
    def _monitor_hyperv_status(self):
        """Background thread to monitor Hyper-V status"""
        try:
            # Initial delay to not overload startup
            time.sleep(20)
            
            # Check if we're on Windows before continuing
            if platform.system() != 'Windows':
                return
                
            while self.background_monitoring:
                # Check much less frequently - every 5 minutes
                time.sleep(300)
                
                # Exit if application is closing
                if not hasattr(self, 'root') or not self.root.winfo_exists():
                    return
                
                try:
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
                    
                    if process.returncode == 0 and stdout.strip() == "Running":
                        # Only check running VMs if Hyper-V service is running
                        try:
                            vm_command = [
                                "powershell",
                                "-Command",
                                "Get-VM | Where-Object {$_.State -eq 'Running'} | Measure-Object | Select-Object -ExpandProperty Count"
                            ]
                            
                            vm_process = subprocess.Popen(
                                vm_command,
                                stdout=subprocess.PIPE,
                                stderr=subprocess.PIPE,
                                text=True,
                                creationflags=subprocess.CREATE_NO_WINDOW
                            )
                            
                            vm_stdout, vm_stderr = vm_process.communicate(timeout=10)
                            
                            # Only log if we have running VMs
                            if vm_process.returncode == 0 and vm_stdout.strip() and int(vm_stdout.strip()) > 0:
                                vm_count = int(vm_stdout.strip())
                                ui_queue.put(("log", ("hyperv", f"Running VMs: {vm_count}")))
                        except:
                            pass
                
                except Exception as e:
                    print(f"Hyper-V monitoring error: {str(e)}")
        
        except Exception as e:
            print(f"Error in Hyper-V monitoring thread: {str(e)}")
            
    def on_tab_changed(self, event):
        """Handle tab change event"""
        tab_id = self.tabs.select()
        tab_name = self.tabs.tab(tab_id, "text")
        
        # Log the tab change
        self.log(f"Switched to {tab_name} tab", "info")
        
        # Force refresh dashboard when changing tabs to ensure it's up to date
        self.update_dashboard(force=True)

    def init_cleanup_tab(self):
        """Initialize the cleanup tab with dedicated buttons"""
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
        
        # Additional cleanup tools
        self.create_button(cleanup_frame, "Remove Unused Packages", 
                          "Removes unused packages and updates",
                          lambda: self.remove_unused_packages(), 1, 0)
        
        self.create_button(cleanup_frame, "Clean Browser Data", 
                          "Cleans browser caches and temporary files",
                          lambda: self.clean_browser_data(), 1, 1)
        
        self.create_button(cleanup_frame, "Delete Old Restore Points", 
                          "Deletes old system restore points",
                          lambda: self.delete_restore_points(), 1, 2)
        
        self.create_button(cleanup_frame, "Clean Event Logs", 
                          "Clears Windows event logs",
                          lambda: self.clean_event_logs(), 1, 3)
        
        # Output area frame
        output_frame = ttk.LabelFrame(frame, text="Operation Output", padding=10, style='Group.TLabelframe')
        output_frame.grid(column=1, row=1, rowspan=2, padx=10, pady=5, sticky=tk.NSEW)
        
        # Output text area
        self.cleanup_output_text = scrolledtext.ScrolledText(output_frame, wrap=tk.WORD, width=50, height=15,
                                                     font=LOG_FONT, bg=LOG_BG)
        self.cleanup_output_text.pack(fill=tk.BOTH, expand=True)
        
        # Progress frame
        progress_frame = ttk.Frame(frame, style='TFrame')
        progress_frame.grid(column=0, row=2, pady=10, sticky=tk.NSEW)
        
        # Progress bar
        ttk.Label(progress_frame, text="Progress:").pack(side=tk.LEFT, padx=5)
        self.cleanup_progress = ttk.Progressbar(progress_frame, orient=tk.HORIZONTAL, length=200, mode='determinate')
        self.cleanup_progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # Configure grid weights
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        for i in range(1, 3):
            frame.rowconfigure(i, weight=1)
    
    def init_system_tab(self):
        """Initialize the system tools tab with dedicated buttons"""
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
        
        # System tools buttons - Column 1
        self.create_button(tools_frame, "Enable Group Policy Editor", 
                          "Enables the Group Policy Editor in Windows Home editions",
                          lambda: self.enable_group_policy_editor(), 0, 0)
        
        self.create_button(tools_frame, "System Information", 
                          "Shows detailed system information",
                          lambda: self.show_system_information(), 1, 0)
        
        self.create_button(tools_frame, "Task Manager", 
                          "Opens Windows Task Manager",
                          lambda: self.open_task_manager(), 2, 0)
        
        self.create_button(tools_frame, "Services", 
                          "Opens Windows Services",
                          lambda: self.open_services(), 3, 0)
        
        # System tools buttons - Column 2
        self.create_button(tools_frame, "Registry Editor", 
                          "Opens Registry Editor",
                          lambda: self.open_registry_editor(), 0, 1)
        
        self.create_button(tools_frame, "Device Manager", 
                          "Opens Device Manager",
                          lambda: self.open_device_manager(), 1, 1)
        
        self.create_button(tools_frame, "System Configuration", 
                          "Opens System Configuration (msconfig)",
                          lambda: self.open_system_config(), 2, 1)
        
        self.create_button(tools_frame, "Power Options", 
                          "Opens Power Options",
                          lambda: self.open_power_options(), 3, 1)
                          
        # Admin tools frame
        admin_frame = ttk.LabelFrame(frame, text="Administrative Tools", padding=8, style='Group.TLabelframe')
        admin_frame.grid(column=0, row=2, sticky=tk.NSEW, pady=5, padx=5)
        
        # Admin buttons - First row
        self.create_button(admin_frame, "Command Prompt", 
                          "Opens Command Prompt as Administrator",
                          lambda: self.open_command_prompt(), 0, 0)
        
        self.create_button(admin_frame, "PowerShell", 
                          "Opens PowerShell as Administrator",
                          lambda: self.open_powershell(), 0, 1)
        
        self.create_button(admin_frame, "Event Viewer", 
                          "Opens Event Viewer",
                          lambda: self.open_event_viewer(), 0, 2)
        
        # Admin buttons - Second row
        self.create_button(admin_frame, "Disk Management", 
                          "Opens Disk Management",
                          lambda: self.open_disk_management(), 1, 0)
        
        self.create_button(admin_frame, "Computer Management", 
                          "Opens Computer Management",
                          lambda: self.open_computer_management(), 1, 1)
        
        self.create_button(admin_frame, "Control Panel", 
                          "Opens Control Panel",
                          lambda: self.open_control_panel(), 1, 2)
        
        # Output area frame
        output_frame = ttk.LabelFrame(frame, text="System Information", padding=10, style='Group.TLabelframe')
        output_frame.grid(column=1, row=1, rowspan=2, padx=10, pady=5, sticky=tk.NSEW)
        
        # Output text area
        self.system_output_text = scrolledtext.ScrolledText(output_frame, wrap=tk.WORD, width=50, height=20,
                                                    font=LOG_FONT, bg=LOG_BG)
        self.system_output_text.pack(fill=tk.BOTH, expand=True)
        
        # Configure grid weights
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        for i in range(1, 3):
            frame.rowconfigure(i, weight=1)
    
    def init_update_tab(self):
        """Initialize the Windows Update tab with dedicated buttons"""
        frame = ttk.Frame(self.update_tab, style='Tab.TFrame', padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        ttk.Label(
            frame, 
            text="Windows Update", 
            font=HEADING_FONT,
            foreground=PRIMARY_COLOR
        ).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        # Update tools frame
        update_frame = ttk.LabelFrame(frame, text="Update Actions", padding=8, style='Group.TLabelframe')
        update_frame.grid(column=0, row=1, sticky=tk.NSEW, pady=5, padx=5)
        
        # Update action buttons
        self.create_button(update_frame, "Check for Updates", 
                          "Checks for available Windows updates",
                          lambda: self.check_for_updates(), 0, 0)
        
        self.create_button(update_frame, "Install Updates", 
                          "Installs available Windows updates",
                          lambda: self.install_updates(), 1, 0)
        
        self.create_button(update_frame, "Update History", 
                          "Shows Windows Update history",
                          lambda: self.show_update_history(), 2, 0)
        
        self.create_button(update_frame, "Reset Windows Update", 
                          "Resets Windows Update components",
                          lambda: self.reset_windows_update(), 3, 0)
        
        # Update settings frame
        settings_frame = ttk.LabelFrame(frame, text="Update Settings", padding=8, style='Group.TLabelframe')
        settings_frame.grid(column=0, row=2, sticky=tk.NSEW, pady=5, padx=5)
        
        # Update settings buttons
        self.create_button(settings_frame, "Configure Windows Update", 
                          "Opens Windows Update settings",
                          lambda: self.configure_windows_update(), 0, 0)
        
        self.create_button(settings_frame, "Pause Updates", 
                          "Temporarily pauses Windows updates",
                          lambda: self.pause_updates(), 1, 0)
        
        self.create_button(settings_frame, "Resume Updates", 
                          "Resumes Windows updates if paused",
                          lambda: self.resume_updates(), 2, 0)
        
        # Output area
        output_frame = ttk.LabelFrame(frame, text="Update Status", padding=10, style='Group.TLabelframe')
        output_frame.grid(column=1, row=1, rowspan=2, padx=10, pady=5, sticky=tk.NSEW)
        
        # Output text area
        self.update_output_text = scrolledtext.ScrolledText(output_frame, wrap=tk.WORD, width=50, height=20,
                                                    font=LOG_FONT, bg=LOG_BG)
        self.update_output_text.pack(fill=tk.BOTH, expand=True)
        
        # Configure grid weights
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        for i in range(1, 3):
            frame.rowconfigure(i, weight=1)
    
    def init_storage_tab(self):
        """Initialize the Storage tab with dedicated buttons"""
        frame = ttk.Frame(self.storage_tab, style='Tab.TFrame', padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Title 
        ttk.Label(
            frame, 
            text="Storage Management", 
            font=HEADING_FONT,
            foreground=PRIMARY_COLOR
        ).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        # Storage tools frame
        storage_frame = ttk.LabelFrame(frame, text="Storage Tools", padding=8, style='Group.TLabelframe')
        storage_frame.grid(column=0, row=1, sticky=tk.NSEW, pady=5, padx=5)
        
        # Storage tools buttons
        self.create_button(storage_frame, "Disk Analyzer", 
                          "Analyzes disk space usage",
                          lambda: self.analyze_disk_space(), 0, 0)
        
        self.create_button(storage_frame, "Find Large Files", 
                          "Locates large files that can be deleted",
                          lambda: self.find_large_files(), 1, 0)
        
        self.create_button(storage_frame, "Storage Sense", 
                          "Opens Windows Storage Sense settings",
                          lambda: self.open_storage_sense(), 2, 0)
        
        self.create_button(storage_frame, "Disk Defragmenter", 
                          "Opens Disk Defragmenter",
                          lambda: self.open_disk_defrag(), 3, 0)
        
        # Backup tools frame
        backup_frame = ttk.LabelFrame(frame, text="Backup Tools", padding=8, style='Group.TLabelframe')
        backup_frame.grid(column=0, row=2, sticky=tk.NSEW, pady=5, padx=5)
        
        # Backup tools buttons
        self.create_button(backup_frame, "File History", 
                          "Opens Windows File History",
                          lambda: self.open_file_history(), 0, 0)
        
        self.create_button(backup_frame, "System Backup", 
                          "Creates a full system backup",
                          lambda: self.create_system_backup(), 1, 0)
        
        self.create_button(backup_frame, "Restore Files", 
                          "Restores files from a backup",
                          lambda: self.restore_from_backup(), 2, 0)
        
        # Storage info frame
        info_frame = ttk.LabelFrame(frame, text="Drive Information", padding=10, style='Group.TLabelframe')
        info_frame.grid(column=1, row=1, rowspan=2, padx=10, pady=5, sticky=tk.NSEW)
        
        # Storage info with dynamic drive list
        self.storage_tree = ttk.Treeview(info_frame, columns=("size", "used", "free"), show="headings", height=10)
        self.storage_tree.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=5)
        
        # Configure treeview columns
        self.storage_tree.heading("size", text="Total Size")
        self.storage_tree.heading("used", text="Used Space")
        self.storage_tree.heading("free", text="Free Space")
        
        self.storage_tree.column("size", width=100)
        self.storage_tree.column("used", width=100)
        self.storage_tree.column("free", width=100)
        
        # Refresh button
        refresh_btn = ttk.Button(info_frame, text="Refresh Drive Information", 
                              command=self.refresh_drive_info)
        refresh_btn.pack(side=tk.BOTTOM, pady=5)
        
        # Configure grid weights
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        for i in range(1, 3):
            frame.rowconfigure(i, weight=1)
        
        # Initial drive info refresh
        self.refresh_drive_info()
    
    def init_network_tab(self):
        """Initialize the Network tab with dedicated buttons"""
        frame = ttk.Frame(self.network_tab, style='Tab.TFrame', padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        ttk.Label(
            frame, 
            text="Network Tools", 
            font=HEADING_FONT,
            foreground=PRIMARY_COLOR
        ).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        # Network tools frame
        network_frame = ttk.LabelFrame(frame, text="Network Tools", padding=8, style='Group.TLabelframe')
        network_frame.grid(column=0, row=1, sticky=tk.NSEW, pady=5, padx=5)
        
        # Network tools buttons - first column
        self.create_button(network_frame, "IP Configuration", 
                          "Shows IP configuration information",
                          lambda: self.show_ip_config(), 0, 0)
        
        self.create_button(network_frame, "Ping Test", 
                          "Tests network connectivity",
                          lambda: self.ping_test(), 1, 0)
        
        self.create_button(network_frame, "DNS Flush", 
                          "Flushes DNS cache",
                          lambda: self.flush_dns(), 2, 0)
        
        self.create_button(network_frame, "Open Ports Scan", 
                          "Scans for open ports",
                          lambda: self.scan_open_ports(), 3, 0)
        
        # Network tools buttons - second column
        self.create_button(network_frame, "Speed Test", 
                          "Tests internet connection speed",
                          lambda: self.test_connection_speed(), 0, 1)
        
        self.create_button(network_frame, "Network Reset", 
                          "Resets network adapter settings",
                          lambda: self.reset_network(), 1, 1)
        
        self.create_button(network_frame, "Network Connections", 
                          "Opens Network Connections panel",
                          lambda: self.open_network_connections(), 2, 1)
        
        self.create_button(network_frame, "Advanced Network Settings", 
                          "Opens advanced network settings",
                          lambda: self.open_network_settings(), 3, 1)
        
        # Wifi tools frame
        wifi_frame = ttk.LabelFrame(frame, text="WiFi Tools", padding=8, style='Group.TLabelframe')
        wifi_frame.grid(column=0, row=2, sticky=tk.NSEW, pady=5, padx=5)
        
        # WiFi tools buttons
        self.create_button(wifi_frame, "WiFi Networks", 
                          "Shows available WiFi networks",
                          lambda: self.show_wifi_networks(), 0, 0)
        
        self.create_button(wifi_frame, "Saved WiFi Passwords", 
                          "Shows saved WiFi passwords",
                          lambda: self.show_wifi_passwords(), 1, 0)
        
        self.create_button(wifi_frame, "Reset WiFi", 
                          "Resets WiFi adapter",
                          lambda: self.reset_wifi(), 2, 0)
        
        # Output frame
        output_frame = ttk.LabelFrame(frame, text="Network Information", padding=10, style='Group.TLabelframe')
        output_frame.grid(column=1, row=1, rowspan=2, padx=10, pady=5, sticky=tk.NSEW)
        
        # Network output text
        self.network_output_text = scrolledtext.ScrolledText(output_frame, wrap=tk.WORD, width=50, height=20,
                                                     font=LOG_FONT, bg=LOG_BG)
        self.network_output_text.pack(fill=tk.BOTH, expand=True)
        
        # Configure grid weights
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        for i in range(1, 3):
            frame.rowconfigure(i, weight=1)
    
    def init_hyperv_tab(self):
        """Initialize the Hyper-V tab with dedicated buttons"""
        frame = ttk.Frame(self.hyperv_tab, style='Tab.TFrame', padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        ttk.Label(
            frame, 
            text="Hyper-V Management", 
            font=HEADING_FONT,
            foreground=PRIMARY_COLOR
        ).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        # Hyper-V management frame
        hyperv_frame = ttk.LabelFrame(frame, text="Hyper-V Management", padding=8, style='Group.TLabelframe')
        hyperv_frame.grid(column=0, row=1, sticky=tk.NSEW, pady=5, padx=5)
        
        # Hyper-V management buttons
        self.create_button(hyperv_frame, "Install Hyper-V", 
                          "Installs Hyper-V feature on Windows",
                          lambda: self.install_hyperv(), 0, 0)
        
        self.create_button(hyperv_frame, "Hyper-V Manager", 
                          "Opens Hyper-V Manager",
                          lambda: self.open_hyperv_manager(), 1, 0)
        
        self.create_button(hyperv_frame, "VM Configuration", 
                          "Configure virtual machine settings",
                          lambda: self.configure_vm(), 2, 0)
        
        self.create_button(hyperv_frame, "Virtual Switch Manager", 
                          "Opens Virtual Switch Manager",
                          lambda: self.open_virtual_switch_manager(), 3, 0)
        
        # VM operations frame
        vm_frame = ttk.LabelFrame(frame, text="VM Operations", padding=8, style='Group.TLabelframe')
        vm_frame.grid(column=0, row=2, sticky=tk.NSEW, pady=5, padx=5)
        
        # VM operation buttons
        self.create_button(vm_frame, "Start VM", 
                          "Starts selected virtual machine",
                          lambda: self.start_vm(), 0, 0)
        
        self.create_button(vm_frame, "Stop VM", 
                          "Stops selected virtual machine",
                          lambda: self.stop_vm(), 0, 1)
        
        self.create_button(vm_frame, "Create VM", 
                          "Creates a new virtual machine",
                          lambda: self.create_vm(), 1, 0)
        
        self.create_button(vm_frame, "Delete VM", 
                          "Deletes selected virtual machine",
                          lambda: self.delete_vm(), 1, 1)
        
        # VM list frame
        vm_list_frame = ttk.LabelFrame(frame, text="Virtual Machines", padding=10, style='Group.TLabelframe')
        vm_list_frame.grid(column=1, row=1, rowspan=2, padx=10, pady=5, sticky=tk.NSEW)
        
        # VM list with treeview
        self.vm_tree = ttk.Treeview(vm_list_frame, columns=("state", "memory", "cpu"), show="headings", height=10)
        self.vm_tree.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=5)
        
        # Configure treeview columns
        self.vm_tree.heading("state", text="State")
        self.vm_tree.heading("memory", text="Memory")
        self.vm_tree.heading("cpu", text="CPU")
        
        self.vm_tree.column("state", width=80)
        self.vm_tree.column("memory", width=80)
        self.vm_tree.column("cpu", width=50)
        
        # Refresh button
        refresh_btn = ttk.Button(vm_list_frame, text="Refresh VM List", 
                              command=self.refresh_vm_list)
        refresh_btn.pack(side=tk.BOTTOM, pady=5)
        
        # Configure grid weights
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        for i in range(1, 3):
            frame.rowconfigure(i, weight=1)
        
        # Initial VM refresh
        self.refresh_vm_list()
    
    def init_optimize_tab(self):
        """Initialize the Optimize tab with dedicated buttons"""
        frame = ttk.Frame(self.optimize_tab, style='Tab.TFrame', padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        ttk.Label(
            frame, 
            text="System Optimization", 
            font=HEADING_FONT,
            foreground=PRIMARY_COLOR
        ).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        # Optimization tools frame
        optimize_frame = ttk.LabelFrame(frame, text="Optimization Tools", padding=8, style='Group.TLabelframe')
        optimize_frame.grid(column=0, row=1, sticky=tk.NSEW, pady=5, padx=5)
        
        # Optimization tools buttons
        self.create_button(optimize_frame, "One-Click Optimization", 
                          "Performs multiple optimization tasks at once",
                          lambda: self.one_click_optimize(), 0, 0, is_primary=True)
        
        self.create_button(optimize_frame, "Clear Startup Programs", 
                          "Manages startup programs",
                          lambda: self.manage_startup_programs(), 1, 0)
        
        self.create_button(optimize_frame, "Disable Unnecessary Services", 
                          "Disables unnecessary Windows services",
                          lambda: self.disable_unnecessary_services(), 2, 0)
        
        self.create_button(optimize_frame, "Memory Optimization", 
                          "Optimizes system memory usage",
                          lambda: self.optimize_memory(), 3, 0)
        
        # Performance tools frame
        perf_frame = ttk.LabelFrame(frame, text="Performance Tools", padding=8, style='Group.TLabelframe')
        perf_frame.grid(column=0, row=2, sticky=tk.NSEW, pady=5, padx=5)
        
        # Performance tools buttons
        self.create_button(perf_frame, "Visual Effects Settings", 
                          "Adjusts visual effects for performance",
                          lambda: self.adjust_visual_effects(), 0, 0)
        
        self.create_button(perf_frame, "Power Plan Settings", 
                          "Adjusts power plan for performance",
                          lambda: self.adjust_power_plan(), 1, 0)
        
        self.create_button(perf_frame, "Advanced System Settings", 
                          "Opens advanced system settings",
                          lambda: self.open_advanced_system_settings(), 2, 0)
        
        # Output frame
        output_frame = ttk.LabelFrame(frame, text="Optimization Results", padding=10, style='Group.TLabelframe')
        output_frame.grid(column=1, row=1, rowspan=2, padx=10, pady=5, sticky=tk.NSEW)
        
        # Optimization output text
        self.optimize_output_text = scrolledtext.ScrolledText(output_frame, wrap=tk.WORD, width=50, height=20,
                                                      font=LOG_FONT, bg=LOG_BG)
        self.optimize_output_text.pack(fill=tk.BOTH, expand=True)
        
        # Configure grid weights
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        for i in range(1, 3):
            frame.rowconfigure(i, weight=1)
    
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

    # Cleanup tab functions
    def run_disk_cleanup(self):
        """Run the Windows Disk Cleanup utility"""
        self.log("Running Disk Cleanup...", "info")
        self.update_status("Running Disk Cleanup...")
        
        # Clear output
        self.cleanup_output_text.delete(1.0, tk.END)
        self.cleanup_output_text.insert(tk.END, "Starting Windows Disk Cleanup...\n")
        
        # Set progress
        self.cleanup_progress['value'] = 10
        
        try:
            # Run cleanmgr.exe
            process = subprocess.Popen(
                ["cleanmgr.exe", "/sagerun:1"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            self.cleanup_output_text.insert(tk.END, "Disk Cleanup utility launched.\n")
            self.cleanup_output_text.insert(tk.END, "Please complete the operation in the opened window.\n")
            
            # Set progress
            self.cleanup_progress['value'] = 100
            
        except Exception as e:
            self.cleanup_output_text.insert(tk.END, f"Error: {str(e)}\n")
            self.log(f"Error running Disk Cleanup: {str(e)}", "error")
            
        self.update_status("Disk Cleanup launched")
    
    def empty_recycle_bin(self):
        """Empty the Recycle Bin"""
        self.log("Emptying Recycle Bin...", "info")
        self.update_status("Emptying Recycle Bin...")
        
        # Clear output
        self.cleanup_output_text.delete(1.0, tk.END)
        self.cleanup_output_text.insert(tk.END, "Emptying Recycle Bin...\n")
        
        # Set progress
        self.cleanup_progress['value'] = 20
        
        try:
            if 'winshell' not in sys.modules:
                self.cleanup_output_text.insert(tk.END, "Error: winshell module not available.\n")
                self.cleanup_output_text.insert(tk.END, "Please install it with: pip install winshell\n")
                self.log("winshell module not available", "error")
                return
                
            import winshell
            winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
            
            self.cleanup_output_text.insert(tk.END, "Recycle Bin emptied successfully.\n")
            self.log("Recycle Bin emptied", "success")
            
            # Set progress
            self.cleanup_progress['value'] = 100
            
        except Exception as e:
            self.cleanup_output_text.insert(tk.END, f"Error: {str(e)}\n")
            self.log(f"Error emptying Recycle Bin: {str(e)}", "error")
            
        self.update_status("Recycle Bin emptied")
    
    # System tab functions
    def open_task_manager(self):
        """Open Windows Task Manager"""
        self.log("Opening Task Manager...", "info")
        
        try:
            subprocess.Popen(["taskmgr.exe"])
            self.system_output_text.delete(1.0, tk.END)
            self.system_output_text.insert(tk.END, "Task Manager opened.\n")
        except Exception as e:
            self.log(f"Error opening Task Manager: {str(e)}", "error")
            self.system_output_text.delete(1.0, tk.END)
            self.system_output_text.insert(tk.END, f"Error opening Task Manager: {str(e)}\n")
    
    def show_system_information(self):
        """Show detailed system information"""
        self.log("Collecting system information...", "info")
        self.update_status("Gathering system information...")
        
        self.system_output_text.delete(1.0, tk.END)
        self.system_output_text.insert(tk.END, "Collecting system information...\n\n")
        
        try:
            # OS information
            os_name = platform.system()
            os_version = platform.version()
            os_build = platform.platform()
            self.system_output_text.insert(tk.END, f"OS: {os_name} {os_version}\n")
            self.system_output_text.insert(tk.END, f"Build: {os_build}\n\n")
            
            # CPU information
            cpu_info = platform.processor()
            cpu_count = psutil.cpu_count(logical=False)
            cpu_threads = psutil.cpu_count(logical=True)
            self.system_output_text.insert(tk.END, f"CPU: {cpu_info}\n")
            self.system_output_text.insert(tk.END, f"Physical cores: {cpu_count}\n")
            self.system_output_text.insert(tk.END, f"Logical processors: {cpu_threads}\n\n")
            
            # Memory information
            memory = psutil.virtual_memory()
            total_memory_gb = memory.total / (1024**3)
            available_memory_gb = memory.available / (1024**3)
            used_memory_gb = memory.used / (1024**3)
            self.system_output_text.insert(tk.END, f"Total memory: {total_memory_gb:.2f} GB\n")
            self.system_output_text.insert(tk.END, f"Available memory: {available_memory_gb:.2f} GB\n")
            self.system_output_text.insert(tk.END, f"Used memory: {used_memory_gb:.2f} GB\n\n")
            
            # Disk information
            self.system_output_text.insert(tk.END, "Disk Information:\n")
            for partition in psutil.disk_partitions():
                try:
                    partition_usage = psutil.disk_usage(partition.mountpoint)
                    self.system_output_text.insert(tk.END, 
                                                f"  {partition.device} ({partition.mountpoint})\n")
                    self.system_output_text.insert(tk.END, 
                                                f"    Total: {partition_usage.total / (1024**3):.2f} GB\n")
                    self.system_output_text.insert(tk.END, 
                                                f"    Used: {partition_usage.used / (1024**3):.2f} GB\n")
                    self.system_output_text.insert(tk.END, 
                                                f"    Free: {partition_usage.free / (1024**3):.2f} GB\n\n")
                except:
                    pass
            
            # Network information
            self.system_output_text.insert(tk.END, "Network Information:\n")
            for interface_name, interface_addresses in psutil.net_if_addrs().items():
                for address in interface_addresses:
                    if address.family == socket.AF_INET:
                        self.system_output_text.insert(tk.END, 
                                                    f"  {interface_name}: {address.address}\n")
            
            self.log("System information collected", "success")
        except Exception as e:
            self.system_output_text.insert(tk.END, f"Error collecting system information: {str(e)}\n")
            self.log(f"Error collecting system information: {str(e)}", "error")
        
        self.update_status("System information collected")
    
    # Network tab functions
    def show_ip_config(self):
        """Show IP configuration information"""
        self.log("Getting IP configuration...", "info")
        self.update_status("Getting IP configuration...")
        
        # Clear output
        self.network_output_text.delete(1.0, tk.END)
        self.network_output_text.insert(tk.END, "Collecting IP configuration information...\n\n")
        
        try:
            # Run ipconfig command
            process = subprocess.Popen(
                ["ipconfig", "/all"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            stdout, stderr = process.communicate(timeout=10)
            
            if stderr:
                self.network_output_text.insert(tk.END, f"Error: {stderr}\n")
            else:
                self.network_output_text.insert(tk.END, stdout)
            
            self.log("IP configuration retrieved", "success")
            
        except Exception as e:
            self.network_output_text.insert(tk.END, f"Error: {str(e)}\n")
            self.log(f"Error getting IP configuration: {str(e)}", "error")
            
        self.update_status("IP configuration retrieved")
    
    # Storage tab functions
    def refresh_drive_info(self):
        """Refresh drive information in storage tab"""
        self.log("Refreshing drive information...", "info")
        
        # Clear existing items
        for item in self.storage_tree.get_children():
            self.storage_tree.delete(item)
        
        try:
            # Get disk partitions
            partitions = psutil.disk_partitions()
            
            for partition in partitions:
                try:
                    usage = psutil.disk_usage(partition.mountpoint)
                    
                    # Calculate sizes in GB
                    total_gb = usage.total / (1024**3)
                    used_gb = usage.used / (1024**3)
                    free_gb = usage.free / (1024**3)
                    
                    # Add to treeview
                    self.storage_tree.insert('', 'end', text=partition.device, values=(
                        f"{total_gb:.2f} GB",
                        f"{used_gb:.2f} GB ({usage.percent}%)",
                        f"{free_gb:.2f} GB"
                    ))
                except:
                    # Skip drives that can't be accessed (like CD-ROM with no disc)
                    pass
            
            self.log("Drive information updated", "success")
            
        except Exception as e:
            self.log(f"Error refreshing drive information: {str(e)}", "error")
    
    # Hyper-V tab functions
    def refresh_vm_list(self):
        """Refresh the virtual machine list"""
        self.log("Refreshing VM list...", "info")
        
        # Clear existing items
        for item in self.vm_tree.get_children():
            self.vm_tree.delete(item)
        
        try:
            # Check if Hyper-V is installed
            ps_check_command = [
                "powershell",
                "-Command",
                "(Get-WindowsOptionalFeature -FeatureName Microsoft-Hyper-V-All -Online).State"
            ]
            
            process = subprocess.Popen(
                ps_check_command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            stdout, stderr = process.communicate(timeout=10)
            
            if process.returncode != 0 or stdout.strip() != "Enabled":
                self.vm_tree.insert('', 'end', text="Hyper-V not installed", values=(
                    "N/A", "N/A", "N/A"
                ))
                return
            
            # Get VM list with PowerShell
            ps_command = [
                "powershell",
                "-Command",
                "Get-VM | Select-Object Name, State, MemoryAssigned, ProcessorCount | ConvertTo-Csv -NoTypeInformation"
            ]
            
            process = subprocess.Popen(
                ps_command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            stdout, stderr = process.communicate(timeout=10)
            
            if process.returncode != 0 or not stdout.strip():
                self.vm_tree.insert('', 'end', text="No VMs found", values=(
                    "N/A", "N/A", "N/A"
                ))
                return
            
            # Parse CSV output
            import csv
            from io import StringIO
            
            csv_reader = csv.reader(StringIO(stdout))
            next(csv_reader)  # Skip header row
            
            for row in csv_reader:
                if len(row) >= 4:
                    name = row[0]
                    state = row[1]
                    memory = int(row[2]) / (1024**3) if row[2].isdigit() else 0
                    cpu = row[3]
                    
                    self.vm_tree.insert('', 'end', text=name, values=(
                        state,
                        f"{memory:.2f} GB",
                        cpu
                    ))
            
            self.log("VM list updated", "success")
            
        except Exception as e:
            self.log(f"Error refreshing VM list: {str(e)}", "error")
            self.vm_tree.insert('', 'end', text=f"Error: {str(e)}", values=(
                "N/A", "N/A", "N/A"
            ))

# Main entry point for the application
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
    version_label = tk.Label(splash_frame, text=f"Version {VERSION}", 
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
        root.after(2000, app.update_dashboard)
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

# This is the main entry point
if __name__ == "__main__":
    # If running directly, show the splash screen
    show_splash_screen() 