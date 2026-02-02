# Amazon Order Generation System

A streamlined system for generating product ordering Excel files from SKU input, with intelligent accessory management and ERP integration.

## 🎯 Core Functionality

### 1. **Generate Product Orders** (SKU → Excel + PO Import)
Input product SKUs and quantities to automatically generate:
- Factory-grouped Excel order files with rich text formatting
- Required accessories based on parent-child relationships
- PO import files for ERP synchronization
- **Automatic buyer assignment** based on product listings (JIXIU vs PINXIU)

### 2. **Template Management** (Excel ↔ JSON with Rich Text)
- Convert Excel files to JSON templates with full rich text support
- Update JSON templates from Excel files
- Preserve formatting (bold, colors) throughout the process

### 3. **Update Parent-Child Relationships** (ERP → System)
- Import accessory relationships from ERP Excel exports
- Update product-accessory mappings
- Download product images from ERP

### 4. **Buyer Management**
- Buyer information stored directly in Excel templates (cell B69)
- Automatically flows through JSON to all outputs
- Parent products get buyer, accessories left blank
- Easy to update: edit Excel B69, convert to JSON

---

## 📁 Project Structure

```
amazon_order/
├── order_generation/
│   ├── docs/
│   │   ├── accessory_mapping.json      # Product-accessory relationships
│   │   ├── empty_base_template.xlsx    # Base Excel template
│   │   ├── PO_import_empty.xlsx        # PO import template for ERP
│   │   ├── Storage.txt                 # Warehouse list
│   │   ├── order_template.md           # JSON template format guide
│   │   └── excel_rich_text_guide.md    # Rich text formatting guide
│   │
│   ├── images/
│   │   ├── products/                   # Product images (SKU.jpg)
│   │   ├── accessories/                # Accessory images (SKU.jpg)
│   │   ├── colors/                     # Color card images
│   │   └── logos/                      # Logo images
│   │
│   ├── json_template/                  # JSON templates (one per SKU)
│   ├── json_exports/                   # Generated order JSON files
│   ├── PO_excel_template/              # Source Excel template files (manual edits)
│   ├── PO_excel_export/                # Generated Excel files (output)
│   └── PO_import_filled/               # Generated ERP import files
│
├── requirements.txt                    # Python dependencies
├── setup_windows.bat                   # Windows setup script
└── 安装依赖.bat                        # Chinese setup script
```

---

## 🚀 Quick Start

### Installation

1. **Install Python dependencies:**
   ```bash
   # Windows (double-click)
   setup_windows.bat
   
   # Or manually
   pip install -r requirements.txt
   ```

2. **Required packages:**
   - `openpyxl` - Excel file processing
   - `pillow` - Image handling
   - `pyperclip` - Clipboard support (for GUI)

### Basic Usage

#### **Option A: GUI (Recommended)**

```bash
cd order_generation
python product_search_gui.py
```

Features:
- Search products by name or SKU
- Select quantities
- Auto-generate accessories
- Choose warehouse per product
- Generate Excel + PO import files

#### **Option B: Command Line**

```bash
cd order_generation
python direct_sku_to_json.py --name ORDER_NAME SKU1 QTY1 SKU2 QTY2 ...
```

Example:
```bash
python direct_sku_to_json.py --name 25AM027 48-82P3-QSFG 800 Elasticbrush01 500 --po-import
```

This generates:
- `json_exports/25AM027-1.json`, `25AM027-2.json`, ... (grouped by factory)
- `PO_excel_export/25AM027-1.xlsx`, `25AM027-2.xlsx`, ... (Excel orders)
- `PO_import_filled/PO_import_25AM027.xlsx` (ERP import file)

---

## 📋 Core Scripts

### Order Generation
| Script | Purpose |
|--------|---------|
| `product_search_gui.py` | **Main GUI** - Search products, select quantities, generate orders |
| `direct_sku_to_json.py` | **CLI tool** - Generate orders from SKU/quantity pairs |
| `json_PO_excel.py` | Convert JSON templates to Excel files |
| `merge_json_templates.py` | Merge multiple JSON templates by factory |
| `fill_po_import.py` | Generate ERP PO import Excel files |

### Template Management
| Script | Purpose |
|--------|---------|
| `excel_to_json_template.py` | Convert Excel to JSON templates (with rich text) |

### ERP Integration
| Script | Purpose |
|--------|---------|
| `accessory_mapping_updater_gui.py` | Update product-accessory mappings from ERP export |

---

## 🔄 Typical Workflow

### **Workflow 1: Generate New Order**

1. **Launch GUI:**
   ```bash
   python product_search_gui.py
   ```

2. **Select products:**
   - Search by product name or SKU
   - Input quantity
   - Add to pool
   - (Optional) Set warehouse for each product

3. **Generate order:**
   - Enter order name (e.g., "25AM027")
   - Check "生成采购导入" for ERP import
   - Click "生成命令" then "执行命令"

4. **Output files:**
   - Excel orders: `PO_excel_export/25AM027-1.xlsx`, etc.
   - PO import: `PO_import_filled/PO_import_25AM027.xlsx`

### **Workflow 2: Update Templates from Excel**

When you manually edit an Excel order file:

1. **Save edited file** in `PO_excel_template/` folder

2. **Convert to JSON template:**
   ```bash
   python excel_to_json_template.py
   # Or double-click to launch GUI
   ```

3. **Select Excel file** and it will update corresponding JSON templates in `json_template/`

### **Workflow 3: Update Accessory Mappings from ERP**

1. **Export from ERP:**
   - Export "导出产品-按SKU-*.xlsx" (includes "关联辅料" tab)

2. **Launch updater:**
   ```bash
   python accessory_mapping_updater_gui.py
   ```

3. **Process:**
   - Select exported Excel file
   - Review changes in Preview tab
   - Apply changes to update `docs/accessory_mapping.json`

---

## 📝 Data Format

### JSON Template Structure

Each SKU has a JSON template with three sections:

```json
{
  "cells": {
    "B3": {"key": "供货商：", "value": "Factory Name"},
    "G3": {"key": "订单号", "value": "ORDER-001"},
    "G4": {"key": "日期", "value": "2026年02月02日"},
    "B14": {"key": "交货时间", "value": "45天"}
  },
  "products": [
    {
      "产品编号": "SKU-001",
      "产品名称": "Product Name",
      "产品图片": "order_generation/images/products/SKU-001.jpg",
      "描述": {
        "type": "rich_text",
        "content": [
          {"text": "Regular text ", "bold": false, "color": "000000"},
          {"text": "Bold red text", "bold": true, "color": "FF0000"}
        ]
      },
      "数量/个": 1000,
      "单价": 12.50,
      "包装方式": "Carton packing"
    }
  ],
  "footer": {
    "buyer": "Buyer Name",
    "supplier": "Supplier Name"
  }
}
```

### Rich Text Support

The system fully supports rich text formatting:
- **Bold** and **colors** in Excel → JSON → Excel
- Text-based tags: `[BOLD:RED]text[/BOLD]`
- Direct Excel formatting (Ctrl+B, font colors)

See `docs/excel_rich_text_guide.md` for details.

---

## 🔧 Configuration Files

### `docs/accessory_mapping.json`
Defines product-accessory relationships:

```json
{
  "products": {
    "PARENT-SKU": {
      "name": "Product Name",
      "accessories": [
        {
          "sku": "ACC-SKU",
          "name": "Accessory Name",
          "ratio_main": "1",
          "ratio_accessory": "2"
        }
      ]
    }
  }
}
```

**Ratio Example:** If `ratio_main=1` and `ratio_accessory=2`, ordering 100 main products includes 200 accessories.

### Excel Templates (B69 - Buyer Field)
Each product template includes buyer information in cell B69:
- Edit this field to set/update buyer for a product
- Automatically extracted to JSON during conversion
- Used in generated Excel (B69) and PO import (采购方)
- Parent products get buyer, accessories left blank

### `docs/Storage.txt`
List of available warehouses (one per line):
```
义乌仓库
深圳仓库
默认仓库
```

---

## 🎨 Rich Text Formatting

### Method 1: Excel Formatting (Recommended)
1. Open Excel file
2. Select text in description cell
3. Apply formatting:
   - **Bold:** Ctrl+B
   - **Color:** Home → Font Color
4. Save and convert to JSON

### Method 2: Text Tags
```
[BOLD:RED]Important text[/BOLD] [NORMAL:000000]regular text[/NORMAL]
```

Supported colors: RED, BLUE, GREEN, BLACK, YELLOW, ORANGE, PURPLE, or hex codes (e.g., FF0000)

---

## ⚠️ Important Notes

### Directory Usage
- **`PO_excel_template/`** - Store **source** Excel template files (manual edits)
- **`PO_excel_export/`** - **Generated** Excel files (auto-created, do not edit)
- **`json_template/`** - One JSON file per SKU (product templates)
- **`json_exports/`** - Temporary order JSON files (factory-grouped)

### Order Rules
- Each order should contain products from a **single factory**
- System automatically **groups by supplier** (cell B3 in templates)
- **Accessories are auto-included** based on `accessory_mapping.json`

### File Naming
- JSON templates: `{SKU}.json` (e.g., `48-82P3-QSFG.json`)
- Order files: `{ORDER_NAME}-{N}.json` (e.g., `25AM027-1.json`, `25AM027-2.json`)
- Images: `{SKU}.jpg` (e.g., `48-82P3-QSFG.jpg`)

---

## 🐛 Troubleshooting

### "Template file not found"
- Ensure `docs/empty_base_template.xlsx` exists
- Run from `order_generation/` directory

### "No products found in Excel"
- Check if Excel file follows the standard template format
- Ensure product table starts around row 7
- Verify column headers match (产品编号, 数量/个, 单价)

### "Image not found"
- Place product images in `images/products/{SKU}.jpg`
- Place accessory images in `images/accessories/{SKU}.jpg`
- Ensure filename matches SKU exactly

### PO Import not generating
- Add `--po-import` flag when using CLI
- Check "生成采购导入" checkbox in GUI
- Verify `docs/PO_import_empty.xlsx` template exists

---

## 📚 Additional Documentation

- **`docs/order_template.md`** - Detailed JSON template format specification
- **`docs/excel_rich_text_guide.md`** - Complete rich text formatting guide
- **`.gitignore`** - Configured to ignore temporary and generated files

---

## 🔄 Updates

**Latest Changes:**
- Removed redundant utility scripts
- Fixed security issues (path traversal prevention)
- Improved error handling and validation
- Simplified project structure
- Enhanced documentation

---

## 💡 Tips

1. **Use the GUI** (`product_search_gui.py`) for most tasks - it's the easiest way
2. **Keep templates updated** - Run `excel_to_json_template.py` after editing Excel files
3. **Backup mappings** - Use "Create Backup" in `accessory_mapping_updater_gui.py`
4. **Check image paths** - Missing images will show `[图片未找到]` in Excel
5. **Test orders** - Review generated Excel files before sending to suppliers

---

## 📞 Support

For issues or questions:
1. Check `docs/` folder for detailed guides
2. Review error messages in terminal/GUI
3. Ensure all dependencies are installed: `pip install -r requirements.txt`
