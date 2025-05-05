import re

# Read the file
with open('system_utilities.py', 'r', encoding='utf-8') as f:
    content = f.readlines()

# Fix indentation
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
        
# Write back to file
with open('system_utilities.py', 'w', encoding='utf-8') as f:
    f.writelines(fixed_content)

print("Indentation fixed successfully!") 