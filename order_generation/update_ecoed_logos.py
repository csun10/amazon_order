#!/usr/bin/env python3
"""
Update Logo field for all ECOED products in JSON templates.
This script adds "ECOED_shampoo_brush.png" to the Logo field (G14) for all EC- products.
"""

import json
from pathlib import Path

def update_ecoed_logos():
    """Update Logo field for ECOED products"""
    template_dir = Path(__file__).parent / "json_template"
    logo_value = "ECOED_shampoo_brush.png"
    
    updated_count = 0
    skipped_count = 0
    error_count = 0
    
    # Find all JSON files
    json_files = list(template_dir.glob("*.json"))
    print(f"Found {len(json_files)} JSON files in json_template/")
    print(f"Looking for ECOED products (EC- prefix or containing 'EC' in SKU)...\n")
    
    for json_file in json_files:
        try:
            # Read the JSON file
            with open(json_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # Check if this is an ECOED product
            # Look at the order number (G3 cell) or filename
            order_number = data.get("cells", {}).get("G3", {}).get("value", "")
            file_sku = json_file.stem
            
            # Check if it's an ECOED product (EC- prefix or 2EC)
            is_ecoed = (
                file_sku.upper().startswith("EC-") or 
                file_sku.upper().startswith("2EC") or
                order_number.upper().startswith("EC-") or
                order_number.upper().startswith("2EC") or
                "ECOED" in order_number.upper() or
                "ECOED" in file_sku.upper()
            )
            
            if not is_ecoed:
                skipped_count += 1
                continue
            
            # Update Logo field (G14)
            if "cells" in data and "G14" in data["cells"]:
                current_logo = data["cells"]["G14"].get("value", "")
                
                # Only update if it's empty or different
                if current_logo != logo_value:
                    data["cells"]["G14"]["value"] = logo_value
                    
                    # Write back to file
                    with open(json_file, 'w', encoding='utf-8') as f:
                        json.dump(data, f, ensure_ascii=False, indent=2)
                    
                    print(f"[OK] Updated: {json_file.name} (was: '{current_logo}')")
                    updated_count += 1
                else:
                    print(f"[i] Already set: {json_file.name}")
                    skipped_count += 1
            else:
                print(f"[!] Missing G14 cell: {json_file.name}")
                error_count += 1
                
        except Exception as e:
            print(f"[FAIL] Error processing {json_file.name}: {e}")
            error_count += 1
    
    print("\n" + "="*60)
    print("Summary:")
    print(f"  Updated: {updated_count}")
    print(f"  Skipped (non-ECOED or already set): {skipped_count}")
    print(f"  Errors: {error_count}")
    print(f"  Total processed: {len(json_files)}")

if __name__ == "__main__":
    update_ecoed_logos()
