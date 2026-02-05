#!/usr/bin/env python3
"""
Test script to debug folder extraction
"""

import subprocess
import os

user_email = "martijn@christoffersonrobb.com"
script_path = "/home/oleg/AppsUtils/do_email/do_email_search.py"

# Run the command
result = subprocess.run(
    ["/home/oleg/AppsUtils/do_email/.venv/bin/python", script_path, "--folders"],
    capture_output=True,
    text=True,
    cwd=os.path.dirname(script_path)
)

print("=== STDOUT ===")
print(result.stdout)
print("\n=== STDERR ===")
print(result.stderr)
print("\n=== Return Code ===")
print(result.returncode)

# Parse output to extract folder names
folders = []
in_folder_section = False

for line in result.stdout.split('\n'):
    # Detect folder section
    if f"Folders for: {user_email}" in line:
        in_folder_section = True
        print(f"\n>>> Found folder section start: {line}")
        continue
    
    if in_folder_section:
        stripped = line.strip()
        print(f">>> Processing line: '{stripped}'")
        
        if stripped and len(stripped) > 0:
            # Check if line starts with a number followed by ". "
            if stripped[0].isdigit():
                parts = stripped.split('. ', 1)
                print(f"    Parts after split: {parts}")
                if len(parts) > 1:
                    folder_name = parts[1].split(' (ID:')[0].strip()
                    print(f"    Extracted folder: {folder_name}")
                    folders.append(folder_name)
            elif stripped.startswith('✓'):
                print(f">>> End of folder section: {stripped}")
                break

print(f"\n=== FINAL FOLDERS ===")
print(folders)
