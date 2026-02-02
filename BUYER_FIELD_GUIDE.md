# Buyer Field (采购方) Guide

**Date:** 2026-02-02  
**Status:** ✅ Production Ready

---

## 📋 **Overview**

The system supports dual-buyer functionality by storing buyer information **directly in the Excel templates** (cell B69). This information flows through the system automatically.

---

## 🔄 **How It Works**

### **Data Flow:**
```
Excel Template (B69)
    ↓ (excel_to_json_template.py)
JSON Template (footer.buyer)
    ↓ (direct_sku_to_json.py)
Generated Excel (B69) + PO Import (采购方 column)
```

### **Key Points:**
1. **Buyer is stored in Excel template B69**
2. **Excel→JSON conversion preserves it**
3. **Order generation uses it automatically**
4. **Parent products get buyer, accessories stay blank**

---

## 📝 **How to Set/Update Buyer Information**

### **Method: Edit Excel Template**

1. **Open the Excel template:**
   ```
   order_generation/PO_excel/{SKU}.xlsx
   ```

2. **Edit cell B69** (采购方/Buyer):
   - 宁波集秀美容科技有限公司 (JIXIU)
   - 宁波品秀美容科技有限公司 (PINXIU)
   - Or any custom buyer name

3. **Save the Excel file**

4. **Update the JSON template:**
   ```bash
   cd order_generation
   python excel_to_json_template.py
   ```
   - Select the updated Excel file
   - Converts Excel → JSON
   - Buyer info is now in `json_template/{SKU}.json`

5. **Done!** Future orders will use the updated buyer

---

## 🎯 **Buyer Assignment Rules**

### **Parent Products:**
- **Definition:** Products with their own templates in `json_template/`
- **Behavior:** Get buyer from their template's `footer.buyer` field
- **Excel output:** Buyer shown in B69
- **PO import:** Buyer filled in 采购方 column

### **Accessories:**
- **Definition:** Products without their own templates (added via accessory_mapping.json)
- **Behavior:** Buyer field left **blank**
- **Excel output:** Inherits parent's buyer in B69 (for display)
- **PO import:** Blank in 采购方 column (as designed)

---

## 📊 **Example**

### **Scenario:**
- Order ST1122-1 (parent) + accessories
- ST1122-1's template has buyer = "宁波集秀美容科技有限公司"

### **Results:**

**Excel Output (`buyer_test-1.xlsx`):**
- Cell B69: "宁波集秀美容科技有限公司"

**PO Import (`PO_import_buyer_test.xlsx`):**
| SKU | Type | 采购方 (Buyer) |
|-----|------|----------------|
| ST1122-1 | Parent | 宁波集秀美容科技有限公司 |
| US-RB01-01 | Accessory | *(blank)* |
| XLZ | Accessory | *(blank)* |

---

## 🔧 **Bulk Update Buyers**

### **To update multiple products at once:**

1. **Identify which products need updating:**
   ```bash
   # List all templates
   dir json_template\*.json
   ```

2. **Edit Excel templates:**
   - Open each `PO_excel/{SKU}.xlsx`
   - Update B69 with correct buyer
   - Save

3. **Batch convert:**
   ```bash
   python excel_to_json_template.py
   ```
   - Select all updated Excel files
   - Click "开始转换"

4. **Verify:**
   ```bash
   python -c "import json; data = json.load(open('json_template/SKU.json')); print(data.get('footer', {}).get('buyer'))"
   ```

---

## ✅ **Verification**

### **Check a template's buyer:**
```bash
python -c "import json; data = json.load(open('json_template/ST1122-1.json', encoding='utf-8')); print('Buyer:', data['footer']['buyer'])"
```

### **Test order generation:**
```bash
python direct_sku_to_json.py --name test --po-import SKU 100
```
- Check Excel B69
- Check PO import 采购方 column

---

## 🎨 **Current Buyers**

### **宁波集秀美容科技有限公司 (JIXIU):**
- Use for JIXIUBeauty-US products
- Example SKUs: ST1122-1, Elasticbrush01, EEHB-NBB

### **宁波品秀美容科技有限公司 (PINXIU):**
- Use for PinxiuBeautyUS products  
- Example SKUs: B10-MJB2-BK, EC404

### **Adding New Buyers:**
- Simply type the new company name in B69
- No code changes needed
- System will use whatever value is in B69

---

## 💡 **Design Benefits**

### **Simple & Maintainable:**
- ✅ No separate mapping files
- ✅ No lookup systems
- ✅ Buyer stored with product data
- ✅ Easy to see and update

### **Flexible:**
- ✅ Each product can have different buyer
- ✅ Add new buyers without code changes
- ✅ Update via familiar Excel interface
- ✅ Version controlled with templates

### **Reliable:**
- ✅ Data stays with template
- ✅ No sync issues
- ✅ Clear audit trail
- ✅ Automatic propagation

---

## 📁 **File Locations**

### **Source (Edit These):**
```
order_generation/PO_excel/{SKU}.xlsx        # Excel templates (edit B69)
```

### **Generated (Auto-Updated):**
```
order_generation/json_template/{SKU}.json   # JSON templates (footer.buyer)
order_generation/PO_excel_export/*.xlsx     # Generated orders (B69)
order_generation/PO_import_filled/*.xlsx    # PO import (采购方)
```

---

## 🔄 **Workflow**

### **Daily Order Generation:**
```bash
python product_search_gui.py
# or
python direct_sku_to_json.py --name ORDER SKU QTY --po-import
```
→ System automatically uses buyer from templates

### **Update Buyer Info (Rare):**
```bash
# 1. Edit Excel template B69
# 2. Run:
python excel_to_json_template.py
```
→ Select updated files → Convert → Done

---

## ⚠️ **Important Notes**

1. **Buyer info is per-template**, not per-order
2. **Accessories always have blank buyer** in PO import
3. **Excel B69 shows parent's buyer** for visual reference
4. **No external files needed** - all in templates
5. **Updates require Excel→JSON conversion**

---

## 🆚 **Previous vs Current Approach**

### **❌ Previous (Removed):**
- Separate `buyer_mapping.py` module
- External listing Excel file
- Lookup system
- Extra maintenance

### **✅ Current (Simplified):**
- Buyer in Excel template B69
- Part of product template
- Automatic propagation
- Zero extra maintenance

---

## 📞 **FAQ**

**Q: How do I change a product's buyer?**  
A: Edit the Excel template B69, then run `excel_to_json_template.py`

**Q: Can different products have different buyers?**  
A: Yes! Each template has its own buyer field

**Q: Do I need to update anything else?**  
A: No, just B69 in Excel. The system handles the rest.

**Q: Why are accessories blank?**  
A: By design - accessories don't have their own buyers

**Q: Can I add a third buyer?**  
A: Yes! Just type the name in B69. No code changes needed.

**Q: What if B69 is empty?**  
A: The buyer field will be empty in outputs. Set a default if needed.

---

## ✅ **Summary**

**The system is simple:**
1. Buyer lives in Excel template B69
2. Flows to JSON automatically
3. Used in all outputs automatically
4. Update by editing Excel + converting

**No external files, no lookups, no complexity!** 🎉

---

**For daily use:** Just generate orders normally  
**To update buyers:** Edit Excel B69 → Convert to JSON  
**That's it!** ✨
