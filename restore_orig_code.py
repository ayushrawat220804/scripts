import os
import re

# Path to the files
source_file = 'restore_full_code.py'
target_file = 'system_utilities.py'

try:
    # Read the source file to extract the ORIGINAL_CODE variable
    with open(source_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Extract the code between the triple quotes after ORIGINAL_CODE =
    pattern = r'ORIGINAL_CODE = """(.*?)"""'
    match = re.search(pattern, content, re.DOTALL)
    
    if match:
        original_code = match.group(1)
        
        # Write the extracted code to system_utilities.py
        with open(target_file, 'w', encoding='utf-8') as f:
            f.write(original_code)
        
        print(f"Successfully restored {len(original_code)} characters to {target_file}")
        print(f"Line count: {original_code.count(chr(10))+1}")
    else:
        print("Could not find the original code in the source file.")
    
except Exception as e:
    print(f"Error: {str(e)}") 