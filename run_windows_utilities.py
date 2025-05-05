#!/usr/bin/env python3
"""
Windows System Utilities Runner
A comprehensive utility suite for Windows system management and monitoring.
Version 1.1
"""

import sys
import traceback
import os
import ctypes

def is_admin():
    """Check if the application is running with administrator privileges"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def main():
    """Main entry point for the application"""
    try:
        # Check for debug mode
        debug_mode = len(sys.argv) > 1 and sys.argv[1] == "--debug"
        
        # Determine if we need admin rights
        requires_admin = len(sys.argv) > 1 and sys.argv[1] == "--admin"
        
        # If requires admin and not running as admin, restart with admin privileges
        if requires_admin and not is_admin():
            script_path = os.path.abspath(sys.argv[0])
            if os.path.exists(script_path):
                print("Restarting with administrator privileges...")
                # Use ShellExecuteW to run the script as admin
                ctypes.windll.shell32.ShellExecuteW(
                    None, 
                    "runas", 
                    sys.executable, 
                    f'"{script_path}"', 
                    None, 
                    1
                )
                return
            else:
                print("Could not determine script path")
                return
        
        if debug_mode:
            print("Starting in debug mode...")
            from windows_system_utilities import run_debug_app
            run_debug_app()
        else:
            # Normal start with splash screen
            from windows_system_utilities import show_splash_screen
            show_splash_screen()
            
    except Exception as e:
        print(f"ERROR: {type(e).__name__}: {str(e)}")
        print("Traceback:")
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main() 