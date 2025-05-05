with open('restore_full_code.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Find the triple-quoted string that contains the original code
start_index = content.find('ORIGINAL_CODE = """')
if start_index == -1:
    print("Could not find the start of the original code.")
    exit(1)

# Move to the beginning of the actual code content
start_index = start_index + len('ORIGINAL_CODE = """')

# Find the end of the triple-quoted string
end_index = content.find('"""', start_index)
if end_index == -1:
    print("Could not find the end of the original code.")
    exit(1)

# Extract the original code
original_code = content[start_index:end_index]

# Write it to system_utilities.py
with open('system_utilities.py', 'w', encoding='utf-8') as f:
    f.write(original_code)

print(f"Successfully restored {len(original_code)} characters to system_utilities.py")
print(f"Line count: {original_code.count(chr(10))+1}") 