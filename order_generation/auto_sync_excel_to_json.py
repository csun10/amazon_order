#!/usr/bin/env python3
"""
Automated Excel to JSON Sync Script

This script automatically syncs Excel files from PO_excel_export back to json_template
when they are modified within the last 3 days.

Usage:
    python auto_sync_excel_to_json.py [--check-all]
    
Options:
    --check-all    Check all files regardless of modification time
"""

import json
import sys
import os
from pathlib import Path
from datetime import datetime, timedelta
from typing import List, Tuple

try:
    from openpyxl import load_workbook
    from openpyxl.cell.rich_text import CellRichText
except ImportError:
    print("Error: openpyxl is required. Install it with: pip install openpyxl")
    sys.exit(1)

# Import the converter class
sys.path.insert(0, str(Path(__file__).parent))
from excel_to_json_template import ExcelToJsonConverter


class AutoSyncManager:
    def __init__(self, days_threshold: int = 3):
        """
        Initialize the auto-sync manager.
        
        Args:
            days_threshold: Only sync files modified within this many days
        """
        self.root_dir = Path(__file__).resolve().parent
        self.po_excel_dir = self.root_dir / "PO_excel"  # Changed from PO_excel_export
        self.json_template_dir = self.root_dir / "json_template"
        self.days_threshold = days_threshold
        self.converter = ExcelToJsonConverter()
        
        # Ensure directories exist
        self.po_excel_dir.mkdir(exist_ok=True)
        self.json_template_dir.mkdir(exist_ok=True)
    
    def get_recently_modified_excel_files(self, check_all: bool = False) -> List[Tuple[Path, datetime]]:
        """
        Get Excel files modified within the threshold period from PO_excel folder.
        
        Args:
            check_all: If True, return all Excel files regardless of modification time
            
        Returns:
            List of (file_path, modification_time) tuples
        """
        excel_files = []
        cutoff_time = datetime.now() - timedelta(days=self.days_threshold)
        
        for excel_file in self.po_excel_dir.glob("*.xlsx"):
            # Skip temporary Excel files
            if excel_file.name.startswith('~$'):
                continue
            
            # Skip test files
            if excel_file.name.startswith('test_'):
                continue
                
            mod_time = datetime.fromtimestamp(excel_file.stat().st_mtime)
            
            if check_all or mod_time > cutoff_time:
                excel_files.append((excel_file, mod_time))
        
        # Sort by modification time (newest first)
        excel_files.sort(key=lambda x: x[1], reverse=True)
        return excel_files
    
    def sync_excel_to_json(self, excel_file: Path) -> Tuple[bool, str, List[str]]:
        """
        Sync a single Excel file to its corresponding JSON template.
        
        Args:
            excel_file: Path to the Excel file
            
        Returns:
            Tuple of (success, message, updated_json_files)
        """
        try:
            # Convert Excel to JSON - returns list of Path objects
            json_files = self.converter.convert_excel_to_json(excel_file)
            
            if json_files:
                json_names = [f.name for f in json_files]
                message = f"[OK] Updated {len(json_files)} JSON file(s): {', '.join(json_names)}"
                # Convert Path objects to strings for the results
                json_file_strs = [str(f) for f in json_files]
                return True, message, json_file_strs
            else:
                return False, "[FAIL] No JSON files generated", []
                
        except Exception as e:
            return False, f"✗ Error: {str(e)}", []
    
    def run_sync(self, check_all: bool = False, silent: bool = False) -> dict:
        """
        Run the synchronization process.
        
        Args:
            check_all: If True, process all Excel files regardless of modification time
            silent: If True, suppress console output
            
        Returns:
            Dictionary with sync results
        """
        results = {
            'total_excel_files': 0,
            'successful_syncs': 0,
            'failed_syncs': 0,
            'updated_json_files': [],
            'errors': []
        }
        
        # Get files to sync
        excel_files = self.get_recently_modified_excel_files(check_all)
        results['total_excel_files'] = len(excel_files)
        
        if not excel_files:
            if not silent:
                print(f"No Excel files found modified within the last {self.days_threshold} days.")
            return results
        
        if not silent:
            threshold_msg = "all files" if check_all else f"files modified within {self.days_threshold} days"
            print(f"\n{'='*70}")
            print(f"Auto-Sync: Excel → JSON")
            print(f"Checking {threshold_msg}")
            print(f"{'='*70}\n")
        
        # Process each file
        for excel_file, mod_time in excel_files:
            time_ago = datetime.now() - mod_time
            days_ago = time_ago.days
            hours_ago = time_ago.seconds // 3600
            
            if days_ago > 0:
                time_str = f"{days_ago} day{'s' if days_ago > 1 else ''} ago"
            elif hours_ago > 0:
                time_str = f"{hours_ago} hour{'s' if hours_ago > 1 else ''} ago"
            else:
                minutes_ago = time_ago.seconds // 60
                time_str = f"{minutes_ago} minute{'s' if minutes_ago > 1 else ''} ago"
            
            if not silent:
                print(f"\n📄 {excel_file.name}")
                print(f"   Modified: {time_str}")
            
            success, message, json_files = self.sync_excel_to_json(excel_file)
            
            if success:
                results['successful_syncs'] += 1
                results['updated_json_files'].extend(json_files)
                if not silent:
                    print(f"   {message}")
            else:
                results['failed_syncs'] += 1
                results['errors'].append((excel_file.name, message))
                if not silent:
                    print(f"   {message}")
        
        # Print summary
        if not silent:
            print(f"\n{'='*70}")
            print(f"Sync Summary:")
            print(f"  Excel files processed: {results['total_excel_files']}")
            print(f"  Successful syncs: {results['successful_syncs']}")
            print(f"  Failed syncs: {results['failed_syncs']}")
            print(f"  JSON files updated: {len(results['updated_json_files'])}")
            print(f"{'='*70}\n")
            
            if results['errors']:
                print("Errors:")
                for filename, error in results['errors']:
                    print(f"  - {filename}: {error}")
                print()
        
        return results


def main():
    """Main entry point for the script."""
    check_all = '--check-all' in sys.argv
    
    manager = AutoSyncManager(days_threshold=3)
    results = manager.run_sync(check_all=check_all)
    
    # Exit with error code if there were failures
    if results['failed_syncs'] > 0:
        sys.exit(1)
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
