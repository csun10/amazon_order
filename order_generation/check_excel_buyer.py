#!/usr/bin/env python3
import openpyxl
from pathlib import Path

excel_path = Path(__file__).parent / "PO_excel_template" / "US-RB01-01.xlsx"
wb = openpyxl.load_workbook(excel_path)
ws = wb.active

print(f"Excel file: {excel_path.name}")
print(f"B69 (Buyer): {ws['B69'].value}")
