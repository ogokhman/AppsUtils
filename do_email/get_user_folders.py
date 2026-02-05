#!/usr/bin/env python3
"""
Get folders for each user in marketer_team and save to do_user_folders.config
Ignores specific folders like Conversation History, Deleted Items, etc.
"""

import subprocess
import configparser
import os

# Folders to ignore
IGNORE_FOLDERS = {
    "Conversation History",
    "Deleted Items",
    "Money stuff",
    "Outbox",
    "Sync Issues",
    "Drafts"
}

# Folders that start with these prefixes should be ignored
IGNORE_PREFIXES = ["Junk", "RSS"]

def should_ignore_folder(folder_name):
    """Check if folder should be ignored"""
    if folder_name in IGNORE_FOLDERS:
        return True
    
    for prefix in IGNORE_PREFIXES:
        if folder_name.startswith(prefix):
            return True
    
    return False

def get_folders_for_user(user_email, script_path):
    """Run do_email_search.py --folders --user for a specific user"""
    # Run the command with --user parameter
    result = subprocess.run(
        ["/home/oleg/AppsUtils/do_email/.venv/bin/python", script_path, "--folders", "--user", user_email],
        capture_output=True,
        text=True,
        cwd=os.path.dirname(script_path)
    )
    
    # Parse output to extract folder names
    folders = []
    in_folder_section = False
    
    for line in result.stdout.split('\n'):
        # Detect folder section
        if f"Folders for: {user_email}" in line:
            in_folder_section = True
            continue
        
        if in_folder_section:
            # Look for lines like "1. FolderName (ID: ...)"
            stripped = line.strip()
            if stripped and len(stripped) > 0:
                # Check if line starts with a number followed by ". "
                if stripped[0].isdigit():
                    parts = stripped.split('. ', 1)
                    if len(parts) > 1:
                        folder_name = parts[1].split(' (ID:')[0].strip()
                        if not should_ignore_folder(folder_name):
                            folders.append(folder_name)
                elif stripped.startswith('✓'):
                    # End of folder section
                    break
    
    return folders

def main():
    # Read marketer_team members from do.config
    config = configparser.ConfigParser()
    config_path = os.path.join(os.path.dirname(__file__), "do.config")
    config.read(config_path)
    
    members = config.get("marketer_team", "members", fallback="").split(',')
    members = [m.strip() for m in members if m.strip()]
    
    print(f"Found {len(members)} members in marketer_team")
    print(f"Members: {', '.join(members)}\n")
    
    # Script path
    script_path = os.path.join(os.path.dirname(__file__), "do_email_search.py")
    
    # Collect results
    all_results = []
    
    for member in members:
        user_email = f"{member}@christoffersonrobb.com"
        print(f"Getting folders for {user_email}...")
        
        folders = get_folders_for_user(user_email, script_path)
        
        print(f"  Found {len(folders)} folders (after filtering)")
        
        all_results.append({
            'user': user_email,
            'folders': folders
        })
    
    # Write to do_user_folders.config
    output_path = os.path.join(os.path.dirname(__file__), "do_user_folders.config")
    
    with open(output_path, 'w') as f:
        for result in all_results:
            f.write(f"[user]\n")
            f.write(f"user = {result['user']}\n")
            f.write(f"[folders]\n")
            f.write(f"folders = {', '.join(result['folders'])}\n")
            f.write(f"\n")
    
    print(f"\n✓ Results saved to {output_path}")
    print(f"✓ Processed {len(all_results)} users")

if __name__ == "__main__":
    main()
