#!/usr/bin/env python3
"""
One-time script to update all JSON template files with correct product names
from the accessory_mapping.json file.
"""

import json
from pathlib import Path

def load_flat_accessory_mapping():
    """Load and flatten the accessory mapping"""
    mapping_path = Path(__file__).parent / "docs" / "accessory_mapping.json"
    
    with open(mapping_path, "r", encoding="utf-8-sig") as f:
        data = json.load(f)
        products_data = data.get("products", {})
        
    # Build a flat mapping of all SKUs (main products + accessories)
    flat_mapping = {}
    
    # Add main products
    for sku, product_info in products_data.items():
        flat_mapping[sku] = product_info.get("name", "")
        
        # Add all accessories for this main product
        for accessory in product_info.get("accessories", []):
            acc_sku = accessory.get("sku", "")
            acc_name = accessory.get("name", "")
            if acc_sku:
                flat_mapping[acc_sku] = acc_name
    
    return flat_mapping

def update_json_files():
    """Update all JSON template files with correct product names"""
    template_dir = Path(__file__).parent / "json_template"
    mapping = load_flat_accessory_mapping()
    
    print(f"Loaded {len(mapping)} product/accessory names from mapping")
    print(f"Scanning JSON files in: {template_dir}")
    
    updated_count = 0
    skipped_count = 0
    not_found_count = 0
    
    for json_file in template_dir.glob("*.json"):
        try:
            with open(json_file, "r", encoding="utf-8-sig") as f:
                data = json.load(f)
            
            if not data.get("products") or len(data["products"]) == 0:
                skipped_count += 1
                continue
            
            product = data["products"][0]
            sku = product.get("产品编号", "")
            current_name = product.get("产品名称", "")
            
            if not sku:
                skipped_count += 1
                continue
            
            # Check if we have a name in the mapping
            if sku in mapping:
                mapped_name = mapping[sku]
                
                # Update if name is empty or different
                if not current_name or current_name != mapped_name:
                    product["产品名称"] = mapped_name
                    
                    # Write back to file
                    with open(json_file, "w", encoding="utf-8") as f:
                        json.dump(data, f, ensure_ascii=False, indent=2)
                    
                    print(f"[OK] Updated {sku}: '{current_name}' -> '{mapped_name}'")
                    updated_count += 1
                else:
                    skipped_count += 1
            else:
                if not current_name:
                    print(f"[WARN] No mapping found for {sku} in {json_file.name}")
                    not_found_count += 1
                else:
                    skipped_count += 1
                    
        except Exception as e:
            print(f"[ERROR] Failed to process {json_file.name}: {e}")
    
    print(f"\n=== Summary ===")
    print(f"Updated: {updated_count}")
    print(f"Skipped (already correct): {skipped_count}")
    print(f"Not found in mapping: {not_found_count}")
    print(f"Total processed: {updated_count + skipped_count + not_found_count}")

if __name__ == "__main__":
    update_json_files()
