#!/usr/bin/env python3
"""
Script to remove UTF-8 BOM (Byte Order Mark) from JSON files.
UTF-8 BOM causes issues with JSON parsers that expect plain UTF-8 encoding.
"""

import os
import json
import glob
from pathlib import Path

def has_bom(file_path):
    """Check if a file starts with UTF-8 BOM."""
    try:
        with open(file_path, 'rb') as f:
            first_three_bytes = f.read(3)
            # UTF-8 BOM is 0xEF 0xBB 0xBF
            return first_three_bytes == b'\xef\xbb\xbf'
    except Exception as e:
        print(f"Error checking BOM in {file_path}: {e}")
        return False

def remove_bom(file_path):
    """Remove UTF-8 BOM from a file."""
    try:
        # Read with utf-8-sig encoding which automatically handles BOM
        with open(file_path, 'r', encoding='utf-8-sig') as f:
            content = f.read()
        
        # Write back without BOM (using regular utf-8)
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        
        return True
    except Exception as e:
        print(f"Error removing BOM from {file_path}: {e}")
        return False

def validate_json(file_path):
    """Validate that a file contains valid JSON."""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            json.load(f)
        return True
    except json.JSONDecodeError as e:
        print(f"JSON validation failed for {file_path}: {e}")
        return False
    except Exception as e:
        print(f"Error validating JSON in {file_path}: {e}")
        return False

def main():
    # Target directory
    json_template_dir = Path("order_generation/json_template")
    
    if not json_template_dir.exists():
        print(f"Directory {json_template_dir} does not exist!")
        return
    
    # Find all JSON files
    json_files = list(json_template_dir.glob("*.json"))
    
    if not json_files:
        print(f"No JSON files found in {json_template_dir}")
        return
    
    print(f"Found {len(json_files)} JSON files in {json_template_dir}")
    
    # Check which files have BOM
    files_with_bom = []
    for json_file in json_files:
        if has_bom(json_file):
            files_with_bom.append(json_file)
    
    if not files_with_bom:
        print("No files with UTF-8 BOM found!")
        return
    
    print(f"\nFound {len(files_with_bom)} files with UTF-8 BOM:")
    for file_path in files_with_bom:
        print(f"  - {file_path.name}")
    
    # Remove BOM from affected files
    print(f"\nRemoving BOM from {len(files_with_bom)} files...")
    
    success_count = 0
    error_count = 0
    
    for file_path in files_with_bom:
        print(f"Processing {file_path.name}...", end=" ")
        
        if remove_bom(file_path):
            # Validate the JSON is still valid after BOM removal
            if validate_json(file_path):
                print("✓ Success")
                success_count += 1
            else:
                print("✗ BOM removed but JSON validation failed")
                error_count += 1
        else:
            print("✗ Failed to remove BOM")
            error_count += 1
    
    print(f"\nResults:")
    print(f"  Successfully processed: {success_count}")
    print(f"  Errors: {error_count}")
    
    if success_count > 0:
        print(f"\n✓ Successfully removed UTF-8 BOM from {success_count} JSON files.")
        print("Files should now be readable with standard JSON parsers.")

if __name__ == "__main__":
    main()