#!/usr/bin/env python3
"""Align 采购方 (buyer) in accessory templates to match their parent products."""

import json
import openpyxl
from pathlib import Path
from excel_to_json_template import ExcelToJsonConverter

ROOT = Path(__file__).parent
MAPPING_PATH = ROOT / "docs" / "accessory_mapping.json"
EXCEL_TEMPLATE_DIR = ROOT / "PO_excel_template"
JSON_TEMPLATE_DIR = ROOT / "json_template"

# Initialize converter
converter = ExcelToJsonConverter()

def load_accessory_mapping():
    """Load accessory mapping."""
    with open(MAPPING_PATH, 'r', encoding='utf-8-sig') as f:
        data = json.load(f)
    return data.get('products', {})

def get_buyer_from_json_template(sku):
    """Get buyer from JSON template."""
    json_path = JSON_TEMPLATE_DIR / f"{sku}.json"
    if not json_path.exists():
        return None
    
    with open(json_path, 'r', encoding='utf-8-sig') as f:
        data = json.load(f)
    
    return data.get('footer', {}).get('buyer', '')

def update_excel_buyer(excel_path, buyer_value):
    """Update buyer (cell B69) in Excel template."""
    if not excel_path.exists():
        return False
    
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    
    # Update B69
    ws['B69'].value = buyer_value
    
    # Save
    wb.save(excel_path)
    print(f"  Updated Excel: {excel_path.name} -> Buyer: {buyer_value}")
    return True

def align_buyers():
    """Align all accessory buyers to match parent products."""
    
    print("Loading accessory mapping...")
    products = load_accessory_mapping()
    
    print(f"Found {len(products)} parent products\n")
    
    updates_made = 0
    errors = []
    
    for parent_sku, product_info in products.items():
        print(f"\nProcessing parent: {parent_sku}")
        
        # Get parent buyer
        parent_buyer = get_buyer_from_json_template(parent_sku)
        if not parent_buyer:
            print(f"  Warning: Parent {parent_sku} has no buyer, skipping")
            continue
        
        print(f"  Parent buyer: {parent_buyer}")
        
        # Process accessories
        accessories = product_info.get('accessories', [])
        print(f"  Accessories: {len(accessories)}")
        
        for acc in accessories:
            acc_sku = acc.get('sku')
            if not acc_sku:
                continue
            
            print(f"    Checking accessory: {acc_sku}")
            
            # Check current buyer
            current_buyer = get_buyer_from_json_template(acc_sku)
            
            if current_buyer == parent_buyer:
                print(f"      Already aligned [OK]")
                continue
            
            print(f"      Current buyer: {current_buyer or '[EMPTY]'}")
            print(f"      Target buyer: {parent_buyer}")
            
            # Find Excel template
            excel_path = EXCEL_TEMPLATE_DIR / f"{acc_sku}.xlsx"
            
            if not excel_path.exists():
                print(f"      ERROR: Excel template not found")
                errors.append(f"{acc_sku}: Excel template not found")
                continue
            
            # Update Excel
            if update_excel_buyer(excel_path, parent_buyer):
                # Convert to JSON
                try:
                    converter.convert_excel_to_json(excel_path)
                    print(f"      Updated JSON template [OK]")
                    updates_made += 1
                except Exception as e:
                    print(f"      ERROR converting to JSON: {e}")
                    errors.append(f"{acc_sku}: Failed to convert to JSON - {e}")
    
    print("\n" + "=" * 80)
    print("SUMMARY")
    print("=" * 80)
    print(f"Accessory templates updated: {updates_made}")
    print(f"Errors: {len(errors)}")
    
    if errors:
        print("\nErrors encountered:")
        for error in errors:
            print(f"  - {error}")
    
    print("\n[SUCCESS] Buyer alignment complete!")
    return updates_made, errors

if __name__ == "__main__":
    align_buyers()
