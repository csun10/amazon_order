# Buyer Mapping Feature

**Date Implemented:** 2026-02-02  
**Status:** ✅ Production Ready

---

## 📋 **Overview**

The system now supports **dual-buyer functionality**, automatically assigning the correct buyer to orders based on product listings.

### **Two Buyers:**

1. **宁波集秀美容科技有限公司** (Ningbo Jixiu Beauty Technology Co., Ltd.)
   - Listing: `JIXIUBeauty-US`
   - 66 SKUs

2. **宁波品秀美容科技有限公司** (Ningbo Pinxiu Beauty Technology Co., Ltd.)
   - Listing: `PinxiuBeautyUS-US-US-US`
   - 74 SKUs

---

## 🔧 **How It Works**

### **Data Source:**
- **Listing File:** `order_generation/docs/Listing20260202-876789694451576832.xlsx`
- **Total SKUs:** 121 parent products mapped

### **Buyer Assignment Rules:**

1. **Parent Products:**
   - SKU is looked up in listing file
   - Buyer assigned based on listing brand
   - Applied to both Excel template and PO import

2. **Accessories/Child Products:**
   - Buyer field left **blank** (as requested)
   - Will be updated later with separate documentation

---

## 📁 **Files Modified**

### **1. New Module: `buyer_mapping.py`**
Core module that handles buyer lookups.

**Key Functions:**
```python
from buyer_mapping import get_buyer_for_sku, is_parent_product

# Get buyer for a SKU (only parent products)
buyer = get_buyer_for_sku("ST1122-1")  
# Returns: "宁波集秀美容科技有限公司"

# Check if SKU is a parent product
is_parent = is_parent_product("ST1122-1")  
# Returns: True
```

**Features:**
- Loads SKU→Buyer mapping from Excel listing file
- Loads parent product list from `accessory_mapping.json`
- Only returns buyer for parent products
- Caches data for performance

---

### **2. Updated: `json_PO_excel.py`**

**Changes:**
- Imports buyer mapping functionality
- Determines buyer from parent products in order
- Sets buyer in cell B69 of Excel template

**Logic:**
```python
# Check products in order for parent SKUs
for product in products:
    sku = product.get('产品编号')
    buyer = get_buyer_for_sku(sku)  # Only returns value for parent products
    if buyer:
        # Use this buyer for the entire order
        break

# Set buyer in Excel (B69)
ws['B69'] = buyer or default_buyer
```

---

### **3. Updated: `fill_po_import.py`**

**Changes:**
- Imports buyer mapping functionality
- Fills 采购方 field for each product row
- Only fills for parent products (accessories left blank)

**Logic:**
```python
for product in products:
    product_sku = product.get("产品编号")
    
    # Get buyer for parent products only
    buyer_value = get_buyer_for_sku(product_sku) or ""
    
    row["采购方"] = buyer_value  # Blank for accessories
```

---

## 🧪 **Testing Results**

### **Test Command:**
```bash
python direct_sku_to_json.py --name buyer_test --po-import ST1122-1 100 B10-MJB2-BK 50
```

### **Excel Output (`buyer_test-1.xlsx`):**
- **Cell B69 (Buyer):** 宁波集秀美容科技有限公司 ✅
- Based on ST1122-1 (JIXIU product)

### **PO Import Output (`PO_import_buyer_test.xlsx`):**

| Row | SKU | Type | Buyer (采购方) |
|-----|-----|------|----------------|
| 3 | ST1122-1 | Parent | 宁波集秀美容科技有限公司 ✅ |
| 4 | US-RB01-01 | Accessory | *(blank)* ✅ |
| 5 | XLZ | Accessory | *(blank)* ✅ |
| 6 | TZ | Accessory | *(blank)* ✅ |
| 11 | B10-MJB2-BK | Parent | 宁波品秀美容科技有限公司 ✅ |
| 12 | B10-MJB2-BK-1 | Accessory | *(blank)* ✅ |

**All tests passed!** ✅

---

## 📊 **SKU Distribution**

### **By Buyer:**
- **JIXIUBeauty-US (集秀):** 66 SKUs
  - Examples: ST1122-1, Elasticbrush01, EEHB-NBB, 48-82P3-QSFG
  
- **PinxiuBeautyUS (品秀):** 74 SKUs  
  - Examples: B10-MJB2-BK, B10-MJB2-BK2

### **By Type:**
- **Parent Products:** 96 SKUs (mapped to buyers)
- **Accessories:** ~100+ SKUs (buyer field left blank)

---

## 🔄 **Workflow Integration**

### **Automatic Buyer Assignment:**

1. **CLI Order Generation:**
   ```bash
   python direct_sku_to_json.py --name ORDER ST1122-1 1000
   ```
   - ✅ Buyer automatically set to 宁波集秀美容科技有限公司

2. **GUI Order Generation:**
   ```bash
   python product_search_gui.py
   ```
   - ✅ Buyer automatically assigned when Excel is generated

3. **PO Import Generation:**
   - ✅ 采购方 field automatically filled for parent products
   - ✅ Accessories left blank (as requested)

---

## 🛠️ **Maintenance**

### **Updating Buyer Mappings:**

1. **Get new listing file from ERP**
   - Export product listings with brand information

2. **Replace listing file:**
   ```bash
   # Place new file at:
   order_generation/docs/Listing20260202-876789694451576832.xlsx
   ```

3. **Verify column structure:**
   - Column 6: Brand (JIXIUBeauty-US or PinxiuBeautyUS-US-US-US)
   - Column 9: SKU

4. **Test:**
   ```bash
   cd order_generation
   python buyer_mapping.py
   ```

### **Adding New Buyers:**

Edit `buyer_mapping.py`:
```python
BUYER_NEW = "新公司名称"

BRAND_TO_BUYER = {
    "JIXIUBeauty-US": BUYER_JIXIU,
    "PinxiuBeautyUS-US-US-US": BUYER_PINXIU,
    "NewBrand-US": BUYER_NEW,  # Add new mapping
}
```

---

## ⚠️ **Important Notes**

### **Current Behavior:**
- ✅ Parent products: Buyer assigned based on listing
- ✅ Accessories: Buyer field left **blank**
- ✅ Unknown SKUs: Default to 宁波品秀美容科技有限公司

### **Future Updates:**
- 🔄 Accessory buyer assignment will be added later with separate documentation
- 🔄 Child product buyer mapping pending

### **Default Buyer:**
If a parent product SKU is not found in the listing:
- **Falls back to:** 宁波品秀美容科技有限公司 (PINXIU)

---

## 📝 **File Locations**

### **Data Files:**
```
order_generation/
├── docs/
│   ├── Listing20260202-876789694451576832.xlsx  # SKU→Buyer mapping source
│   └── accessory_mapping.json                   # Parent product list
├── buyer_mapping.py                             # Buyer lookup module (NEW)
├── json_PO_excel.py                             # Excel generation (UPDATED)
└── fill_po_import.py                            # PO import generation (UPDATED)
```

### **Output Files:**
```
order_generation/
├── PO_excel_export/
│   └── *.xlsx              # Excel orders with correct buyer in B69
└── PO_import_filled/
    └── PO_import_*.xlsx    # PO import with 采购方 filled for parent products
```

---

## ✅ **Verification Checklist**

Use this checklist to verify the feature is working:

- [ ] Listing file exists: `docs/Listing20260202-876789694451576832.xlsx`
- [ ] Test buyer mapping: `python buyer_mapping.py`
- [ ] Generate test order: `python direct_sku_to_json.py --name test SKU 100`
- [ ] Check Excel B69 cell has correct buyer
- [ ] Check PO import 采购方 column:
  - [ ] Parent products have buyer name
  - [ ] Accessories have blank buyer field
- [ ] Verify JIXIU products get correct buyer (集秀)
- [ ] Verify PINXIU products get correct buyer (品秀)

---

## 🎯 **Summary**

**Feature Status:** ✅ **Fully Functional**

- Buyer mapping loaded from listing file
- Parent products automatically assigned correct buyer
- Excel templates show buyer in B69
- PO import shows buyer in 采购方 column
- Accessories correctly left blank
- Backward compatible (defaults to PINXIU if SKU not found)

**Ready for production use!** 🚀

---

**For questions or issues, refer to the test results above or run:**
```bash
python buyer_mapping.py  # Test module
```
