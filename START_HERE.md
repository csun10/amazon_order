# 🚀 START HERE - Amazon Order System

**Last Updated:** 2026-02-02  
**Status:** ✅ Production Ready

---

## 👋 **Welcome!**

This is your streamlined Amazon Order Generation System with **dual-buyer support**.

---

## ⚡ **Quick Start (3 Steps)**

### **Step 1: Install Dependencies**
```bash
# Windows - Double click this file:
setup_windows.bat

# Or manually:
pip install -r requirements.txt
```

### **Step 2: Generate Your First Order**
```bash
cd order_generation
python product_search_gui.py
```

### **Step 3: Use the System**
- Search products by name or SKU
- Add to pool with quantities
- Click "生成命令" then "执行命令"
- Done! Excel files and PO import generated

---

## 📚 **Documentation Guide**

### **🆕 New User? Read These First:**

1. **`README.md`** (10 min read)
   - Complete system overview
   - Installation guide
   - 3 main workflows
   - Troubleshooting

2. **`QUICK_REFERENCE.md`** (5 min read)
   - Daily operations cheat sheet
   - Command examples
   - Best practices

### **📖 Feature Guides:**

3. **`BUYER_MAPPING_FEATURE.md`**
   - NEW dual-buyer support
   - How automatic assignment works
   - JIXIU vs PINXIU distinction

4. **`docs/excel_rich_text_guide.md`**
   - Rich text formatting in Excel
   - Bold and color support

5. **`docs/order_template.md`**
   - JSON template format specification

### **🔧 Technical Docs:**

6. **`CLEANUP_SUMMARY.md`**
   - What was removed and why
   - Before/after comparison

7. **`MINIMAL_FIXES_APPLIED.md`**
   - Code quality improvements
   - Security fixes

8. **`FINAL_SUMMARY.md`**
   - Complete project summary
   - All changes documented

---

## 🎯 **What This System Does**

### **Core Features:**

1. **Smart Order Generation**
   - Input: SKU + quantity
   - Output: Excel files + PO import
   - Auto-includes accessories
   - Groups by factory
   - **Auto-assigns buyer** 🆕

2. **Rich Text Support**
   - Bold and colored text
   - Excel → JSON → Excel preservation
   - Professional formatting

3. **Dual-Buyer System** 🆕
   - 宁波集秀美容科技有限公司 (JIXIU)
   - 宁波品秀美容科技有限公司 (PINXIU)
   - Automatic assignment based on SKU
   - Parent products only

4. **ERP Integration**
   - PO import file generation
   - 采购方 field auto-filled 🆕
   - Warehouse tracking
   - Parent-child relationship updates

---

## 🛠️ **System Status**

| Component | Status | Notes |
|-----------|--------|-------|
| Order Generation | ✅ Working | CLI & GUI |
| Template Management | ✅ Working | Excel ↔ JSON |
| Buyer Mapping | ✅ Working | 121 SKUs mapped |
| ERP Integration | ✅ Working | PO import |
| Input Validation | ✅ Active | Quantity & files |
| Security | ✅ Hardened | Path protection |
| Documentation | ✅ Complete | 4 main guides |

---

## 📂 **Important Directories**

### **Edit These:**
- `PO_excel/` - Source Excel files (manual edits)
- `json_template/` - Product templates (via scripts)
- `images/` - Product images

### **Don't Edit These (Auto-Generated):**
- `PO_excel_export/` - Generated orders
- `json_exports/` - Temporary JSON files
- `PO_import_filled/` - ERP import files

---

## 🎨 **Key Scripts**

### **Daily Use:**
```bash
python product_search_gui.py          # Main GUI (most used)
python direct_sku_to_json.py          # CLI for batch
```

### **Maintenance:**
```bash
python excel_to_json_template.py      # Update templates
python accessory_mapping_updater_gui.py  # Update relationships
python buyer_mapping.py               # Test buyer mappings
```

---

## 🆕 **What's New?**

### **Dual-Buyer Support (Just Added!):**
- Automatic buyer assignment from SKU listings
- 66 SKUs → JIXIU (宁波集秀美容科技有限公司)
- 74 SKUs → PINXIU (宁波品秀美容科技有限公司)
- Integrated into Excel generation and PO import
- Parent products only (accessories blank)

### **Code Quality (Just Fixed!):**
- All security issues resolved
- Input validation added
- Exception handling improved
- 27 redundant files removed

---

## 💡 **Pro Tips**

1. **Use the GUI** - Easiest for daily orders
2. **Check documentation** - Everything is documented
3. **Test small first** - Start with small quantities
4. **Review output** - Check generated files before sending
5. **Backup configs** - Before updating mappings

---

## ⚠️ **Important Notes**

### **Buyer Assignment:**
- **Automatic** - No manual configuration needed
- **Parent products** - Get buyer from listing
- **Accessories** - Left blank (as designed)
- **Default** - Falls back to PINXIU if SKU not found

### **File Organization:**
- **Source files** → `PO_excel/`
- **Generated files** → `PO_excel_export/`
- **Never edit** generated files (will be overwritten)

### **Updates:**
- **Listing file** - Update when catalog changes
- **Templates** - Sync after editing Excel files
- **Mappings** - Update from ERP exports

---

## 🎓 **Learning Path**

### **Beginner (Day 1):**
1. Read this file ✓
2. Run `product_search_gui.py`
3. Generate test order
4. Review `QUICK_REFERENCE.md`

### **Regular User (Day 2-7):**
5. Read `README.md` (complete guide)
6. Learn template updates
7. Practice workflows
8. Review `BUYER_MAPPING_FEATURE.md`

### **Advanced (Ongoing):**
9. Update accessory mappings
10. Manage buyer listings
11. Handle edge cases
12. Maintain system

---

## 📞 **Quick Help**

| Question | Answer |
|----------|--------|
| How do I generate an order? | Run `product_search_gui.py` |
| What buyer will be assigned? | Automatic from listing file |
| How do I update templates? | Run `excel_to_json_template.py` |
| Where are generated files? | `PO_excel_export/` folder |
| How do I update buyers? | Replace listing file in `docs/` |
| Need help? | Check `QUICK_REFERENCE.md` |

---

## 🎉 **You're Ready!**

**Your system is:**
- ✅ Installed and configured
- ✅ Tested and verified
- ✅ Documented and supported
- ✅ Ready for daily use

**Next step:** Run `python product_search_gui.py` and generate your first order!

---

**For complete details, see `README.md`**  
**For daily tasks, see `QUICK_REFERENCE.md`**  
**For new features, see `BUYER_MAPPING_FEATURE.md`**

**Happy ordering! 🚀**
