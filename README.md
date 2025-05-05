# Windows System Utilities

A comprehensive system utility suite for Windows that helps with system monitoring, cleanup, optimization, and troubleshooting through an easy-to-use GUI interface.

## Features

- **Dashboard**: Real-time system monitoring with visual graphs for CPU, memory, disk, and network usage
- **System Cleanup**: Remove temporary files, clear caches, and free up disk space
- **System Tools**: Access useful Windows administrative tools and system information
- **Performance Optimization**: Tools to optimize your system for better performance
- **Network Tools**: View network information and test connectivity
- **Storage Management**: Analyze disk space usage and manage storage devices
- **Hyper-V Management**: Monitor and manage Hyper-V virtual machines
- **Windows Update**: Manage and configure Windows Update settings

## Requirements

- Windows 10 or 11
- Python 3.8 or higher
- Required Python packages (see requirements.txt):
  - psutil
  - matplotlib
  - winshell
  - pywin32

## Installation

1. Clone or download this repository
2. Install required packages:
   ```
   pip install -r requirements.txt
   ```

## Usage

Run the application using:

```
python run_windows_utilities.py
```

For debug mode with additional console output:

```
python run_windows_utilities.py --debug
```

For administrative features (requires administrator privileges):

```
python run_windows_utilities.py --admin
```

## File Structure

- `windows_system_utilities.py` - Main application code
- `run_windows_utilities.py` - Runner script to start the application
- `requirements.txt` - Required Python packages

## Development Notes

- Version 1.2 includes significant performance improvements and bug fixes:
  - Resolved UI freezing issues with improved thread management and message queue system
  - Fixed memory leaks by properly cleaning up matplotlib resources
  - Reduced CPU usage with optimized background monitoring
  - Improved error handling to prevent application crashes
  - Dashboard updates are now more efficient and only refresh visible components
  - Monitoring intervals adjusted to reduce system resource usage

- Consolidated version of multiple utility scripts into a single application
- Diff algorithm timeout has been increased from 5000ms to 30000ms (30 seconds) to handle large files
- Dashboard-based UI replaces log-based monitoring for better user experience
- Each tab has dedicated action buttons instead of displaying everything in logs

## Troubleshooting

If you experience any issues:

1. Try running in debug mode to see more detailed information
2. Make sure all required packages are installed correctly
3. Ensure you have administrator privileges for certain system operations
4. Check if your version of Windows is compatible (Windows 10 or 11)

## License

MIT License 