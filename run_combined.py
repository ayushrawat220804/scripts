#!/usr/bin/env python3
"""
Windows System Utilities Runner
A comprehensive utility suite for Windows system management and monitoring.
"""

import sys
import traceback
from combined_system_utils import show_splash_screen, run_debug_app

if __name__ == "__main__":
    try:
        # Check for debug mode
        if len(sys.argv) > 1 and sys.argv[1] == "--debug":
            print("Starting in debug mode...")
            run_debug_app()
        else:
            # Normal start with splash screen
            show_splash_screen()
    except Exception as e:
        print(f"ERROR: {type(e).__name__}: {str(e)}")
        print("Traceback:")
        traceback.print_exc()
        sys.exit(1) 