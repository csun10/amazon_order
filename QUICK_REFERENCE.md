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

After editing Excel files in `PO_excel_template/`:

```bash
cd order_generation
python excel_to_json_template.py
```

Or double-click the script to launch GUI.

**What it does:**
- Reads Excel files from `PO_excel_template/`
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
| `json_PO_excel.py` | JSON → Excel converter | **(Automatic)** Called by other scripts |
| `merge_json_templates.py` | Merge JSON by factory | **(Automatic)** Called by other scripts |
| `fill_po_import.py` | Generate PO import | **(Automatic)** Called by other scripts |

---

## 📁 Important Directories

| Directory | Purpose | Edit? |
|-----------|---------|-------|
| `json_template/` | Product templates (one per SKU) | ✅ Yes - via `excel_to_json_template.py` |
| `PO_excel_template/` | **Source** Excel template files | ✅ Yes - manual edits allowed |
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

### `docs/PO_import_template.xlsx`
**When to edit:** Only if ERP import format changes

**Purpose:** Template for ERP PO import files

### `docs/Storage.txt`
**When to edit:** When warehouse list changes

**Purpose:** List of available warehouses (one per line)

### Excel Template Cell B69 (Buyer Field)
**When to edit:** When updating product buyer

**Purpose:** Stores buyer for each product  
**Auto-applied:** Flows to JSON, Excel output (B69), and PO import (采购方)

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

### ❌ "Cannot generate files into PO_excel_template/ folder"
**Solution:** Script auto-corrects to `PO_excel_export/`
- Keep source template files in `PO_excel_template/`
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
- Update buyer in Excel B69, then convert to JSON

### ❌ DON'T:
- Edit files in `PO_excel_export/` (they're auto-generated)
- Manually edit `accessory_mapping.json` (use GUI)
- Delete files from `json_template/` without backup
- Mix up `PO_excel_template/` (source) with `PO_excel_export/` (output)
- Manually edit JSON buyer field (edit Excel B69 instead)

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
5. Make any manual edits in PO_excel_template/
6. Run excel_to_json_template.py to sync changes

Weekly:
7. Export product data from ERP
8. Run accessory_mapping_updater_gui.py
9. Update accessory relationships

As needed:
10. Update buyer info by editing Excel B69, then convert to JSON
```

---

## 🏢 Buyer Management

**Buyer information is stored in Excel templates (cell B69):**

### **How to Update Buyer:**
1. Open `PO_excel_template/{SKU}.xlsx`
2. Edit cell B69 (采购方)
3. Save Excel file
4. Run `excel_to_json_template.py` to update JSON
5. Done! Future orders will use new buyer

### **Current Buyers:**
- **宁波集秀美容科技有限公司** (JIXIU)
- **宁波品秀美容科技有限公司** (PINXIU)
- Or any custom buyer name

### **Automatic Flow:**
```
Excel B69 → JSON footer.buyer → Generated Excel B69 + PO Import 采购方
```

---

**That's it! Keep this guide handy for daily operations.** 📌
