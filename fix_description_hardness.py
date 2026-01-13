import openpyxl
from pathlib import Path

po_excel_dir = Path(r"c:\Users\Cheng\Desktop\amazon_order\order_generation\PO_excel")

# The parent SKUs we updated
parent_skus = ["2EC-Blue", "2EC-Green", "2EC-Pink", "2EC-Yellow", 
               "7S-HA5T-5D0X", "EC1601", "EC404", "ZW-YI7D-KWFL"]

# Revert the description field (Cell C7) back to just "硬度55"
# Keep the specification fields with the extended text

print("Reverting 描述 field changes while keeping specification updates...")
print(f"{'='*60}")

for parent_sku in parent_skus:
    excel_file = po_excel_dir / f"{parent_sku}.xlsx"
    
    if not excel_file.exists():
        continue
    
    try:
        wb = openpyxl.load_workbook(excel_file)
        fixed = False
        
        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            
            # Check Cell C7 (description field)
            if sheet['C7'].value and isinstance(sheet['C7'].value, str):
                if "硬度55(伟氏硬度计D型测值65+-0.8)" in sheet['C7'].value:
                    # Revert back to just 硬度55
                    sheet['C7'].value = sheet['C7'].value.replace(
                        "硬度55(伟氏硬度计D型测值65+-0.8)", 
                        "硬度55"
                    )
                    print(f"✓ Reverted {excel_file.name} - Cell C7 (描述) back to '硬度55'")
                    fixed = True
        
        if fixed:
            wb.save(excel_file)
            print(f"  Saved: {excel_file.name}")
    
    except Exception as e:
        print(f"❌ Error processing {excel_file.name}: {e}")

print(f"\n{'='*60}")
print("Summary: Description fields reverted, specification fields retained")
print(f"{'='*60}")
