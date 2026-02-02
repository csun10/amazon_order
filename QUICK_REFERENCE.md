# Quick Reference Guide

**Essential commands and workflows for daily use**

---

## 🚀 Daily Operations

### 1. Generate New Order (GUI - Recommended)

```bash
cd order_generation
python product_search_gui.py
```

**Steps:**
1. Search product by name or SKU
2. Enter quantity
3. Click "添加到池"
4. Repeat for all products
5. Enter order name (e.g., "25AM027")
6. Check "生成采购导入" for ERP import
7. Click "生成命令" → "执行命令"

**Output:**
- `PO_excel_export/25AM027-1.xlsx`, `25AM027-2.xlsx`, etc.
- `PO_import_filled/PO_import_25AM027.xlsx` (if checkbox checked)

---

### 2. Generate New Order (CLI)

```bash
cd order_generation
python direct_sku_to_json.py --name ORDER_NAME SKU1 QTY1 SKU2 QTY2 --po-import
```

**Example:**
```bash
python direct_sku_to_json.py --name 25AM027 48-82P3-QSFG 800 Elasticbrush01 500 --po-import
```

---

### 3. Update JSON Templates from Excel

After editing Excel files in `PO_excel/`:

```bash
cd order_generation
python excel_to_json_template.py
```

Or double-click the script to launch GUI.

**What it does:**
- Reads Excel files from `PO_excel/`
- Updates corresponding JSON templates in `json_template/`
- Preserves rich text formatting

---

### 4. Update Accessory Mappings from ERP

After exporting "导出产品-按SKU-*.xlsx" from ERP:

```bash
cd order_generation
python accessory_mapping_updater_gui.py
```

**Steps:**
1. Click "浏览..." and select exported Excel
2. Click "处理文件"
3. Switch to "Preview Changes" tab
4. Click "Generate Preview" to review
5. Click "Apply Changes" to update

**Updates:** `docs/accessory_mapping.json`

---

## 📋 Core Scripts

| Script | Purpose | When to Use |
|--------|---------|-------------|
| `product_search_gui.py` | Main GUI for order generation | **Daily** - Generate new orders |
| `direct_sku_to_json.py` | CLI for order generation | **Daily** - Automated/batch orders |
| `excel_to_json_template.py` | Excel → JSON converter | **After editing** Excel files |
| `accessory_mapping_updater_gui.py` | ERP → System sync | **After ERP export** updates |
| `buyer_mapping.py` | SKU → Buyer mapping (NEW) | **(Automatic)** Auto-assigns buyers |
| `json_PO_excel.py` | JSON → Excel converter | **(Automatic)** Called by other scripts |
| `merge_json_templates.py` | Merge JSON by factory | **(Automatic)** Called by other scripts |
| `fill_po_import.py` | Generate PO import | **(Automatic)** Called by other scripts |

---

## 📁 Important Directories

| Directory | Purpose | Edit? |
|-----------|---------|-------|
| `json_template/` | Product templates (one per SKU) | ✅ Yes - via `excel_to_json_template.py` |
| `PO_excel/` | **Source** Excel files | ✅ Yes - manual edits allowed |
| `PO_excel_export/` | **Generated** Excel files | ❌ No - auto-generated |
| `json_exports/` | Temporary order JSON | ❌ No - auto-generated |
| `PO_import_filled/` | ERP import files | ❌ No - auto-generated |
| `images/products/` | Product images | ✅ Yes - add {SKU}.jpg |
| `images/accessories/` | Accessory images | ✅ Yes - add {SKU}.jpg |
| `docs/` | Config and templates | ⚠️ Careful - critical files |

---

## 🔧 Configuration Files

### `docs/accessory_mapping.json`
**When to edit:** Never manually - use `accessory_mapping_updater_gui.py`

**Purpose:** Defines which accessories go with which products

### `docs/empty_base_template.xlsx`
**When to edit:** Only for layout changes

**Purpose:** Base template for all Excel orders

### `docs/PO_import_empty.xlsx`
**When to edit:** Only if ERP import format changes

**Purpose:** Template for ERP PO import files

### `docs/Storage.txt`
**When to edit:** When warehouse list changes

**Purpose:** List of available warehouses (one per line)

### `docs/Listing20260202-876789694451576832.xlsx` (NEW)
**When to edit:** After ERP listing export

**Purpose:** Maps SKUs to buyers (JIXIU vs PINXIU)  
**Auto-applied:** Buyer field in Excel (B69) and PO import (采购方)

---

## 🎨 Rich Text Formatting

### Quick Format in Excel
1. Open Excel file
2. Select text in description cell (Column C)
3. Ctrl+B for bold
4. Home → Font Color for colors
5. Save and run `excel_to_json_template.py`

### Text Tags (Alternative)
```
[BOLD:RED]Important[/BOLD] [NORMAL:000000]Regular text[/NORMAL]
```

---

## 🐛 Common Issues & Solutions

### ❌ "Template file not found"
**Solution:** Run from `order_generation/` directory
```bash
cd order_generation
python script_name.py
```

### ❌ "No products found in Excel"
**Solution:** 
- Check product table starts around row 7
- Verify columns: 产品编号, 数量/个, 单价

### ❌ "Image not found"
**Solution:**
- Place image in `images/products/{SKU}.jpg` or `images/accessories/{SKU}.jpg`
- Filename must match SKU exactly

### ❌ "Cannot generate files into PO_excel/ folder"
**Solution:** Script auto-corrects to `PO_excel_export/`
- Keep source files in `PO_excel/`
- Generated files go to `PO_excel_export/`

### ❌ GUI won't start
**Solution:** Install dependencies
```bash
pip install -r requirements.txt
```

---

## 💡 Best Practices

### ✅ DO:
- Use GUI (`product_search_gui.py`) for most tasks
- Keep templates updated after editing Excel
- Backup `accessory_mapping.json` before updates
- Place images in correct folders with SKU names
- Test generated Excel files before sending
- **NEW:** Update listing file when product catalogs change

### ❌ DON'T:
- Edit files in `PO_excel_export/` (they're auto-generated)
- Manually edit `accessory_mapping.json` (use GUI)
- Delete files from `json_template/` without backup
- Mix up `PO_excel/` (source) with `PO_excel_export/` (output)
- **NEW:** Manually edit buyer assignments (automatic from listing)

---

## 📞 Quick Help

1. Check error messages in terminal/GUI
2. Review `README.md` for detailed guides
3. Check `docs/order_template.md` for JSON format
4. Check `docs/excel_rich_text_guide.md` for formatting

---

## 🔄 Typical Daily Workflow

```
Morning:
1. Check ERP for new SKU requirements
2. Run product_search_gui.py
3. Select products and generate orders

Afternoon:
4. Review generated Excel files
5. Make any manual edits in PO_excel/
6. Run excel_to_json_template.py to sync changes

Weekly:
7. Export product data from ERP
8. Run accessory_mapping_updater_gui.py
9. Update accessory relationships
10. Update listing file if product catalogs change
```

---

## 🆕 Buyer Mapping (NEW Feature)

**Automatic buyer assignment for dual-company orders:**

- **集秀 (JIXIU):** 宁波集秀美容科技有限公司 → JIXIUBeauty-US listings
- **品秀 (PINXIU):** 宁波品秀美容科技有限公司 → PinxiuBeautyUS-US-US-US listings

**How it works:**
1. System reads `docs/Listing20260202-876789694451576832.xlsx`
2. Maps each parent SKU to its buyer
3. Automatically fills:
   - Excel template buyer field (B69)
   - PO import 采购方 column (parent products only)
4. Accessories left blank (will be updated later)

**Test buyer mapping:**
```bash
python buyer_mapping.py
```

---

**That's it! Keep this guide handy for daily operations.** 📌
