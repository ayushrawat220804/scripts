import os
import shutil

# Path to the file
file_path = 'system_utilities.py'
backup_file = 'system_utilities.py.bak'

try:
    # Check if file exists and is not empty
    if not os.path.exists(file_path) or os.path.getsize(file_path) == 0:
        print(f"Error: File {file_path} doesn't exist or is empty.")
        # Since the file is already empty, we can't recover directly
        print("Please restore from your backup if available.")
        exit(1)
    
    # Make a backup before editing
    shutil.copy2(file_path, backup_file)
    print(f"Created backup at {backup_file}")
    
    # Read the file
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.readlines()
    
    # Verify we have content
    if not content:
        print("Error: No content read from file.")
        exit(1)
    
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
        exit(1)
    
    # Write back to file
    with open(file_path, 'w', encoding='utf-8') as f:
        f.writelines(fixed_content)
    
    print("Indentation fixed successfully!")
    
except Exception as e:
    print(f"Error: {str(e)}")
    # Restore from backup if there was an error
    if os.path.exists(backup_file):
        print("Restoring from backup...")
        shutil.copy2(backup_file, file_path)
        print("Restore completed.") 