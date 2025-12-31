#!/usr/bin/env python3
"""
Update Logo1 field (G15) for ECOED parent products based on accessory_mapping.json.
Only parent products with "ecoed" in the name get the logo.
Also reverts previous changes to G14.
"""

import json
from pathlib import Path

def update_ecoed_logos():
    """Update Logo1 field for ECOED parent products"""
    base_dir = Path(__file__).parent
    template_dir = base_dir / "json_template"
    mapping_file = base_dir / "docs" / "accessory_mapping.json"
    logo_value = "ECOED_shampoo_brush.png"
    
    # Load accessory mapping to identify parent products
    with open(mapping_file, 'r', encoding='utf-8') as f:
        mapping_data = json.load(f)
    
    # Get list of parent product SKUs with "ecoed" in the name
    ecoed_parent_skus = set()
    for sku, product_data in mapping_data.get("products", {}).items():
        product_name = product_data.get("name", "")
        if product_name and "ecoed" in product_name.lower():
            ecoed_parent_skus.add(sku)
    
    print(f"Found {len(ecoed_parent_skus)} ECOED parent products in accessory_mapping.json:")
    for sku in sorted(ecoed_parent_skus):
        name = mapping_data["products"][sku]["name"]
        print(f"  {sku}: {name}")
    print()
    
    updated_count = 0
    reverted_count = 0
    skipped_count = 0
    error_count = 0
    
    # Process all JSON files
    json_files = list(template_dir.glob("*.json"))
    print(f"Processing {len(json_files)} JSON files...\n")
    
    for json_file in json_files:
        try:
            # Read the JSON file
            with open(json_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            file_sku = json_file.stem
            modified = False
            
            # Check if this is an ECOED parent product
            is_ecoed_parent = file_sku in ecoed_parent_skus
            
            # Revert G14 (Logo) if it was set to ECOED logo
            if "cells" in data and "G14" in data["cells"]:
                current_logo = data["cells"]["G14"].get("value", "")
                if "ECOED_shampoo_brush" in current_logo:
                    data["cells"]["G14"]["value"] = ""
                    print(f"[REVERT] G14 cleared: {json_file.name}")
                    reverted_count += 1
                    modified = True
            
            # Update G15 (Logo1) only for ECOED parent products
            if "cells" in data and "G15" in data["cells"]:
                current_logo1 = data["cells"]["G15"].get("value", "")
                
                if is_ecoed_parent:
                    # This is an ECOED parent - should have the logo
                    if current_logo1 != logo_value:
                        data["cells"]["G15"]["value"] = logo_value
                        print(f"[OK] G15 updated: {json_file.name} (was: '{current_logo1}')")
                        updated_count += 1
                        modified = True
                    else:
                        print(f"[i] G15 already set: {json_file.name}")
                        skipped_count += 1
                else:
                    # Not an ECOED parent - clear if it has ECOED logo
                    if "ECOED_shampoo_brush" in current_logo1:
                        data["cells"]["G15"]["value"] = ""
                        print(f"[REVERT] G15 cleared: {json_file.name} (not an ECOED parent)")
                        reverted_count += 1
                        modified = True
            
            # Write back if modified
            if modified:
                with open(json_file, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                    
        except Exception as e:
            print(f"[FAIL] Error processing {json_file.name}: {e}")
            error_count += 1
    
    print("\n" + "="*60)
    print("Summary:")
    print(f"  G15 Updated (ECOED parents): {updated_count}")
    print(f"  Reverted (G14 or non-parent G15): {reverted_count}")
    print(f"  Skipped (already correct): {skipped_count}")
    print(f"  Errors: {error_count}")
    print(f"  Total processed: {len(json_files)}")

if __name__ == "__main__":
    update_ecoed_logos()
