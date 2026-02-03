#!/usr/bin/env python3
"""Regenerate PO import files from existing JSON exports with underscore to dash conversion."""

import sys
from pathlib import Path
from fill_po_import import fill_po_import_for_order

def regenerate_all_po_imports():
    """Regenerate all PO import files for verify orders."""
    
    json_exports_dir = Path(__file__).parent / "json_exports"
    
    if not json_exports_dir.exists():
        print(f"Error: {json_exports_dir} not found")
        return 1
    
    # Get all unique order names (extract from verify_*.json files)
    json_files = list(json_exports_dir.glob("verify_*.json"))
    
    if not json_files:
        print("No verify JSON files found")
        return 1
    
    # Extract unique order names (e.g., verify_Elasticbrush01 from verify_Elasticbrush01-1.json)
    order_names = set()
    for json_file in json_files:
        # Extract base name before the dash-number
        name = json_file.stem  # e.g., "verify_Elasticbrush01-1"
        # Split by last dash to get base name
        parts = name.rsplit('-', 1)
        if len(parts) == 2 and parts[1].isdigit():
            order_names.add(parts[0])
    
    print(f"Found {len(order_names)} unique orders to regenerate")
    print("=" * 70)
    
    success_count = 0
    error_count = 0
    
    for idx, order_name in enumerate(sorted(order_names), 1):
        print(f"\n[{idx}/{len(order_names)}] Regenerating PO import for: {order_name}")
        
        try:
            po_import_path = fill_po_import_for_order(order_name, warehouse="默认仓库")
            print(f"  SUCCESS: {po_import_path}")
            success_count += 1
        except Exception as e:
            print(f"  ERROR: {e}")
            error_count += 1
    
    print("\n" + "=" * 70)
    print(f"Regeneration Complete")
    print(f"  Success: {success_count}")
    print(f"  Errors: {error_count}")
    print("=" * 70)
    
    return 0 if error_count == 0 else 1

if __name__ == "__main__":
    sys.exit(regenerate_all_po_imports())
