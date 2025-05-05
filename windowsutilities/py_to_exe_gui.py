"""
Python Script to EXE Converter GUI
A graphical interface for converting Python scripts to standalone executable files.
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import subprocess
import re
import shutil
from datetime import datetime
import traceback

class PyToExeGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Python to EXE Converter")
        self.root.geometry("750x600")
        self.root.resizable(True, True)
        self.root.minsize(750, 600)
        
        # Set window icon if available
        try:
            self.root.iconbitmap(default="icon.ico")
        except:
            pass
        
        # Variables for form fields
        self.input_file_var = tk.StringVar()
        self.output_name_var = tk.StringVar()
        self.dest_dir_var = tk.StringVar()
        self.onefile_var = tk.BooleanVar(value=True)
        self.console_var = tk.BooleanVar(value=False)
        self.icon_file_var = tk.StringVar()
        self.cleanup_var = tk.BooleanVar(value=True)
        self.same_name_folder_var = tk.BooleanVar(value=True)
        self.running = False
        self.process = None
        
        # Configure main grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=0)  # Header doesn't expand
        self.root.rowconfigure(1, weight=1)  # Content expands
        self.root.rowconfigure(2, weight=0)  # Footer doesn't expand
        
        # Create styles
        self.create_styles()
        
        # Create UI components
        self.create_header()
        self.create_content()
        self.create_footer()
        
        # Set up protocol for when the window is closed
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
    
    def create_styles(self):
        """Create custom styles for the application"""
        style = ttk.Style()
        
        # Configure TFrame style
        style.configure("TFrame", background="#f0f0f0")
        
        # Configure Header.TLabel style
        style.configure("Header.TLabel", 
                       font=("Arial", 16, "bold"),
                       background="#4a6984", 
                       foreground="white",
                       padding=10)
        
        # Configure subheader style
        style.configure("Subheader.TLabel", 
                       font=("Arial", 12, "bold"),
                       background="#f0f0f0",
                       padding=(0, 10, 0, 5))
        
        # Configure Primary button
        style.configure("Primary.TButton", 
                       font=("Arial", 11, "bold"))
        
        # Configure section frame
        style.configure("Section.TFrame", 
                       padding=10,
                       relief="groove",
                       borderwidth=1)
    
    def create_header(self):
        """Create the header section of the UI"""
        header_frame = ttk.Frame(self.root)
        header_frame.grid(row=0, column=0, sticky="ew")
        
        # Configure header_frame columns
        header_frame.columnconfigure(0, weight=1)
        
        # Add header label
        header_label = ttk.Label(
            header_frame, 
            text="Python to EXE Converter",
            style="Header.TLabel")
        header_label.grid(row=0, column=0, sticky="ew")
    
    def create_content(self):
        """Create the main content section of the UI"""
        content_frame = ttk.Frame(self.root, padding=10)
        content_frame.grid(row=1, column=0, sticky="nsew")
        
        # Configure content_frame columns and rows
        content_frame.columnconfigure(0, weight=1)
        content_frame.rowconfigure(1, weight=1)  # Log area expands
        
        # Create sections
        self.create_input_section(content_frame)
        self.create_log_section(content_frame)
    
    def create_input_section(self, parent):
        """Create the input form section"""
        input_frame = ttk.Frame(parent, style="Section.TFrame")
        input_frame.grid(row=0, column=0, sticky="ew", pady=5)
        
        # Configure columns
        input_frame.columnconfigure(0, weight=0)
        input_frame.columnconfigure(1, weight=1)
        input_frame.columnconfigure(2, weight=0)
        
        # Input file
        ttk.Label(input_frame, text="Python File:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(input_frame, textvariable=self.input_file_var).grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        ttk.Button(input_frame, text="Browse", command=self.browse_input_file).grid(row=0, column=2, padx=5, pady=5)
        
        # Output name
        ttk.Label(input_frame, text="Output Name:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(input_frame, textvariable=self.output_name_var).grid(row=1, column=1, sticky="ew", padx=5, pady=5, columnspan=2)
        
        # Destination directory
        ttk.Label(input_frame, text="Output Folder:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(input_frame, textvariable=self.dest_dir_var).grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        ttk.Button(input_frame, text="Browse", command=self.browse_dest_dir).grid(row=2, column=2, padx=5, pady=5)
        
        # Icon file
        ttk.Label(input_frame, text="Icon File:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(input_frame, textvariable=self.icon_file_var).grid(row=3, column=1, sticky="ew", padx=5, pady=5)
        ttk.Button(input_frame, text="Browse", command=self.browse_icon_file).grid(row=3, column=2, padx=5, pady=5)
        
        # Options
        options_frame = ttk.Frame(input_frame)
        options_frame.grid(row=4, column=0, columnspan=3, sticky="ew", padx=5, pady=5)
        
        ttk.Checkbutton(options_frame, text="Create a single file", variable=self.onefile_var).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Checkbutton(options_frame, text="Show console window", variable=self.console_var).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Checkbutton(options_frame, text="Clean up build files", variable=self.cleanup_var).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Checkbutton(options_frame, text="Save in EXE name folder", variable=self.same_name_folder_var).pack(side=tk.LEFT)
        
        # Buttons
        buttons_frame = ttk.Frame(input_frame)
        buttons_frame.grid(row=5, column=0, columnspan=3, pady=10)
        
        # Convert button
        convert_button = ttk.Button(
            buttons_frame, 
            text="Convert to EXE", 
            command=self.start_conversion,
            style="Primary.TButton",
            width=20)
        convert_button.pack(side=tk.LEFT, padx=(0, 10))
    
    def create_log_section(self, parent):
        """Create the log output section"""
        log_frame = ttk.Frame(parent, style="Section.TFrame")
        log_frame.grid(row=1, column=0, sticky="nsew", pady=5)
        
        # Configure layout
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(1, weight=1)
        
        # Add subheader
        ttk.Label(log_frame, text="Conversion Log", style="Subheader.TLabel").grid(row=0, column=0, sticky="w")
        
        # Add scrolled text widget for log output
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, font=("Consolas", 9))
        self.log_text.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        self.log_text.config(state=tk.DISABLED)
        
        # Configure text tags for coloring
        self.log_text.tag_configure("info", foreground="black")
        self.log_text.tag_configure("success", foreground="green")
        self.log_text.tag_configure("warning", foreground="orange")
        self.log_text.tag_configure("error", foreground="red")
    
    def create_footer(self):
        """Create the footer section of the UI"""
        footer_frame = ttk.Frame(self.root, style="TFrame", padding=5)
        footer_frame.grid(row=2, column=0, sticky="ew")
        
        # Status label on the left
        self.status_var = tk.StringVar(value="Ready")
        status_label = ttk.Label(footer_frame, textvariable=self.status_var)
        status_label.pack(side=tk.LEFT)
        
        # Version label on the right
        version_label = ttk.Label(footer_frame, text="v1.1")
        version_label.pack(side=tk.RIGHT)
    
    def browse_input_file(self):
        """Open file dialog to select Python file"""
        file_path = filedialog.askopenfilename(
            title="Select Python File",
            filetypes=[("Python Files", "*.py *.pyw"), ("All Files", "*.*")]
        )
        if file_path:
            self.input_file_var.set(file_path)
            # Set default output name based on input file
            basename = os.path.basename(file_path)
            output_name = os.path.splitext(basename)[0]
            if not self.output_name_var.get():
                self.output_name_var.set(output_name)
    
    def browse_dest_dir(self):
        """Open folder dialog to select destination directory"""
        dir_path = filedialog.askdirectory(title="Select Output Folder")
        if dir_path:
            self.dest_dir_var.set(dir_path)
    
    def browse_icon_file(self):
        """Open file dialog to select icon file"""
        file_path = filedialog.askopenfilename(
            title="Select Icon File",
            filetypes=[("Icon Files", "*.ico"), ("All Files", "*.*")]
        )
        if file_path:
            self.icon_file_var.set(file_path)
    
    def log_message(self, message, level="info"):
        """Add a message to the log with timestamp"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        full_message = f"{timestamp} - {message}\n"
        
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, full_message, level)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        
        # Force update to make log more responsive
        self.root.update_idletasks()
    
    def update_status(self, message):
        """Update the status bar message"""
        self.status_var.set(message)
        # Force update to make status more responsive
        self.root.update_idletasks()
    
    def start_conversion(self):
        """Start the conversion process in a separate thread"""
        # Validate input
        input_file = self.input_file_var.get().strip()
        if not input_file:
            messagebox.showerror("Error", "Please select a Python file to convert")
            return
        
        if not os.path.exists(input_file):
            messagebox.showerror("Error", "The selected Python file does not exist")
            return
        
        if not input_file.endswith(('.py', '.pyw')):
            if not messagebox.askyesno("Warning", "The selected file doesn't have a .py or .pyw extension. Continue anyway?"):
                return
        
        # Clear log
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        
        # Disable convert button during conversion
        for child in self.root.winfo_children():
            if isinstance(child, ttk.Frame):
                for widget in child.winfo_children():
                    if isinstance(widget, ttk.Button):
                        if widget["text"] == "Convert to EXE":
                            widget.config(state=tk.DISABLED)
        
        # Set running flag
        self.running = True
        
        # Start conversion in a separate thread
        threading.Thread(target=self.run_conversion, daemon=True).start()
    
    def prepare_output_directory(self):
        """Prepare the output directory structure"""
        # Get output name
        output_name = self.output_name_var.get() or os.path.splitext(os.path.basename(self.input_file_var.get()))[0]
        
        # Get destination directory or use default
        if self.dest_dir_var.get():
            base_dir = self.dest_dir_var.get()
        else:
            # Use directory where the input file is located
            base_dir = os.path.join(os.path.dirname(self.input_file_var.get()), "pyexe_output")
        
        # If same name folder option is enabled, create a subfolder with the output name
        if self.same_name_folder_var.get():
            base_dir = os.path.join(base_dir, output_name)
            self.log_message(f"Using same-name folder structure: {base_dir}", "info")
        
        # Create base directory if it doesn't exist
        try:
            os.makedirs(base_dir, exist_ok=True)
            self.log_message(f"Output directory prepared: {base_dir}", "info")
            
            # Create subdirectories for build artifacts
            build_dir = os.path.join(base_dir, "build")
            dist_dir = os.path.join(base_dir, "dist")
            logs_dir = os.path.join(base_dir, "logs")
            
            # Create these directories
            os.makedirs(build_dir, exist_ok=True)
            os.makedirs(dist_dir, exist_ok=True)
            os.makedirs(logs_dir, exist_ok=True)
            
            return {
                "base_dir": base_dir,
                "build_dir": build_dir,
                "dist_dir": dist_dir,
                "logs_dir": logs_dir,
                "output_name": output_name
            }
        except Exception as e:
            self.log_message(f"Error creating output directory: {str(e)}", "error")
            return None
    
    def run_conversion(self):
        """Run the conversion process"""
        try:
            self.log_message("Starting conversion process...", "info")
            self.update_status("Converting...")
            
            # Prepare output directory structure
            dirs = self.prepare_output_directory()
            if not dirs:
                messagebox.showerror("Error", "Failed to create output directories")
                self.enable_convert_button()
                self.running = False
                return
            
            # Check if py_to_exe_converter.py exists in the same directory
            converter_script = os.path.join(os.path.dirname(os.path.abspath(__file__)), "py_to_exe_converter.py")
            if not os.path.exists(converter_script):
                self.log_message(f"Converter script not found: {converter_script}", "error")
                self.log_message("Please ensure py_to_exe_converter.py is in the same directory as this GUI", "error")
                self.update_status("Error: Converter script not found")
                self.enable_convert_button()
                return
            
            # Get output name from directory structure
            output_name = dirs["output_name"]
            
            # Direct PyInstaller command instead of using converter script
            # This gives us more control over the process and output
            cmd = [sys.executable, "-m", "PyInstaller"]
            
            # Add PyInstaller options
            if self.onefile_var.get():
                cmd.append("--onefile")
            
            if not self.console_var.get():
                cmd.append("--windowed")
            
            # Add output name
            cmd.extend(["--name", output_name])
            
            # Specify work directory (build directory)
            cmd.extend(["--workpath", dirs["build_dir"]])
            
            # Specify dist directory
            cmd.extend(["--distpath", dirs["dist_dir"]])
            
            # Specify log file
            log_file = os.path.join(dirs["logs_dir"], f"pyinstaller_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
            cmd.extend(["--log-level", "INFO"])
            
            # Add icon if specified
            if self.icon_file_var.get() and os.path.exists(self.icon_file_var.get()):
                cmd.extend(["--icon", self.icon_file_var.get()])
            
            # Add input file
            cmd.append(self.input_file_var.get())
            
            # Log the command
            cmd_str = " ".join(cmd)
            self.log_message(f"Executing: {cmd_str}", "info")
            
            # Explain the directories that will be created
            self.log_message("Conversion will create these directories in the output folder:", "info")
            self.log_message(f"  • build: Temporary files in '{dirs['build_dir']}'", "info")
            self.log_message(f"  • dist: Finished executable in '{dirs['dist_dir']}'", "info")
            self.log_message(f"  • logs: Log files in '{dirs['logs_dir']}'", "info")
            
            # Install PyInstaller if needed
            try:
                import PyInstaller
                self.log_message("PyInstaller is already installed.", "info")
            except ImportError:
                self.log_message("PyInstaller not found. Installing...", "info")
                try:
                    subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
                    self.log_message("PyInstaller installed successfully.", "success")
                except Exception as e:
                    self.log_message(f"Failed to install PyInstaller: {str(e)}", "error")
                    self.enable_convert_button()
                    self.running = False
                    return
            
            # Run PyInstaller
            self.process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                universal_newlines=True
            )
            
            # Read output in real-time
            self.read_process_output()
            
            # Move files if cleanup is enabled
            if self.cleanup_var.get():
                # Check if spec file was created and move it to logs directory
                spec_file = f"{output_name}.spec"
                if os.path.exists(spec_file):
                    try:
                        shutil.move(spec_file, os.path.join(dirs["logs_dir"], spec_file))
                        self.log_message(f"Moved spec file to logs directory", "info")
                    except Exception as e:
                        self.log_message(f"Failed to move spec file: {str(e)}", "warning")
                
                # Remove build directory if it exists in current directory (not in output dir)
                if os.path.exists("build") and os.path.isdir("build"):
                    try:
                        shutil.rmtree("build")
                        self.log_message("Cleaned up build directory", "info")
                    except Exception as e:
                        self.log_message(f"Failed to clean up build directory: {str(e)}", "warning")
                
                # Remove __pycache__ directory if it exists
                if os.path.exists("__pycache__") and os.path.isdir("__pycache__"):
                    try:
                        shutil.rmtree("__pycache__")
                        self.log_message("Cleaned up __pycache__ directory", "info")
                    except Exception as e:
                        self.log_message(f"Failed to clean up __pycache__ directory: {str(e)}", "warning")
            
            # Final success message
            exe_path = os.path.join(dirs["dist_dir"], f"{output_name}.exe")
            if os.path.exists(exe_path):
                self.log_message("Conversion completed successfully!", "success")
                self.log_message(f"Executable created at: {exe_path}", "success")
                
                # Copy executable to base directory for convenience if using same-name folder
                if self.same_name_folder_var.get():
                    try:
                        exe_base_path = os.path.join(dirs["base_dir"], f"{output_name}.exe")
                        shutil.copy2(exe_path, exe_base_path)
                        self.log_message(f"Copied executable to: {exe_base_path}", "success")
                    except Exception as e:
                        self.log_message(f"Failed to copy executable to base directory: {str(e)}", "warning")
                
                # Offer to open the directory
                if messagebox.askyesno("Conversion Complete", 
                                      f"Executable created successfully!\n\nOpen containing folder?"):
                    # If using same-name folder, open that instead of dist directory
                    if self.same_name_folder_var.get():
                        os.startfile(dirs["base_dir"])
                    else:
                        os.startfile(dirs["dist_dir"])
                
                self.update_status("Conversion completed")
            else:
                self.log_message(f"Expected executable not found at: {exe_path}", "warning")
                self.update_status("Conversion failed")
        
        except Exception as e:
            self.log_message(f"Error during conversion: {str(e)}", "error")
            self.log_message(traceback.format_exc(), "error")
            self.update_status("Error during conversion")
        
        finally:
            # Re-enable the convert button
            self.enable_convert_button()
            self.running = False
    
    def read_process_output(self):
        """Read process output in real-time and update log"""
        for line in iter(self.process.stdout.readline, ''):
            if not line:
                break
            
            line = line.strip()
            if not line:
                continue
                
            # Determine log level based on content
            if re.search(r'\b(error|fail|failed)\b', line.lower()):
                self.log_message(line, "error")
            elif re.search(r'\b(warn|warning)\b', line.lower()):
                self.log_message(line, "warning")
            elif re.search(r'\b(success|succeeded|completed)\b', line.lower()):
                self.log_message(line, "success")
            else:
                self.log_message(line, "info")
        
        # Wait for process to complete
        return_code = self.process.wait()
        
        if return_code != 0:
            self.log_message(f"PyInstaller failed with return code {return_code}", "error")
    
    def enable_convert_button(self):
        """Re-enable the convert button"""
        self.root.after(0, lambda: self._enable_convert_button())
    
    def _enable_convert_button(self):
        """Helper method to enable the convert button (called from main thread)"""
        for child in self.root.winfo_children():
            if isinstance(child, ttk.Frame):
                for widget in child.winfo_children():
                    if isinstance(widget, ttk.Button) and widget["text"] == "Convert to EXE":
                        widget.config(state=tk.NORMAL)
    
    def on_close(self):
        """Handle window close event"""
        if self.running:
            if messagebox.askyesno("Confirm Exit", "Conversion is in progress. Are you sure you want to exit?"):
                # Kill process if running
                if self.process and self.process.poll() is None:
                    try:
                        self.process.terminate()
                    except:
                        pass
                self.root.destroy()
        else:
            self.root.destroy()

def main():
    root = tk.Tk()
    app = PyToExeGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main() 