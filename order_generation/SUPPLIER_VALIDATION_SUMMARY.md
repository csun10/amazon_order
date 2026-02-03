# Supplier Name Validation Report

**Date:** 2026-02-02  
**Reference File:** `Supplier20260202170516-876874265653952512.xlsx`  
**Total Valid Suppliers in ERP:** 50

---

## Executive Summary

This report identifies all PO excel templates and JSON templates with supplier names that don't match the exact supplier names registered in the ERP system. ERP requires exact matching supplier names for successful PO import.

### Key Findings

| Category | PO Excel Templates | JSON Templates |
|----------|-------------------|----------------|
| **Total Files** | 149 | 146 |
| **✓ Matching** | 93 (62.4%) | 90 (61.6%) |
| **✗ Non-Matching** | 27 (18.1%) | 27 (18.5%) |
| **⚠ Blank** | 29 (19.5%) | 28 (19.2%) |
| **⚠ Errors** | 0 (0.0%) | 1 (0.7%) |

**Total Issues: 56 PO Excel Templates + 56 JSON Templates need correction**

---

## Valid Supplier Names (50 total)

These are the EXACT supplier names that must be used in templates:

1. 义乌俊曼
2. 义乌市一柠饰品有限公司
3. 义乌市喜美服饰有限公司
4. 义乌市腾裕包装制品有限公司
5. 伟翔塑业
6. 佳硕尼龙丝
7. 保定金洁源日用品制造有限公司
8. 凯源印刷
9. 吉秀制刷（已停用）
10. 和鑫制刷厂
11. 宁波吉秀美妆用品有限公司
12. 宁波市应发鼎艺美发用品制造有限公司
13. 宁波市柯莱威电子科技有限公司
14. 宁波市海曙硕丰塑料五金制品有限公司
15. 宁波泰丰机械有限公司
16. 宁波瑾秀制刷科技有限公司
17. 宁波百林斯日用品产品有限公司
18. 宁波百林斯日用品有限公司
19. 宁波鑫林达制刷有限公司
20. 宁波锐为实业有限公司
21. 常州百利基
22. 广州华辉
23. 广州市鑫唛子服装辅料有限公司
24. 库存
25. 廊坊振瑞美容美发用品有限公司
26. 明宏沐浴
27. 松溪县健洁刷业有限公司
28. 泰禾制刷
29. 浙江太平洋纸业有限公司
30. 深圳市众拓印刷
31. 深圳市怡劲手袋制品有限公司
32. 温州中鹏
33. 温州尚喜包装有限公司
34. 温州恒浩包装有限公司
35. 灿光五金制品
36. 瑞安市诚叶美发用具有限公司
37. 瑾秀制刷科技有限公司（停用
38. 瑾秀塑业
39. 盐城市国栋布艺包装材料有限公司
40. 盐城市艳丽工艺品有限公司
41. 臻亿源日化
42. 艾盛电器
43. 苍南县顺瀚纸罐有限公司
44. 菲迪印刷
45. 阳江市江城区三盛美容制品厂
46. 阳江市江城区金蚂蚁工贸有限公司
47. 阳江骏业
48. 静远科技
49. 高斯电器
50. 龙港华蓝

---

## Problem 1: Non-Matching Supplier Names (27 Excel + 27 JSON files)

### Issue Type A: Generic Placeholder "XX印刷厂" (14 files)

**Problem:** Using generic placeholder "XX印刷厂" instead of exact supplier name  
**Impact:** ERP will reject import

**Files with "XX印刷厂":**
1. AM413-3-1.xlsx / .json
2. AMSC-green.xlsx / .json
3. B10-MJQ1-PK-01.xlsx / .json
4. BB101-2-1.xlsx / .json
5. BB101-2.xlsx / .json
6. BB101-3.xlsx / .json
7. EC-D1-CYAN.xlsx / .json
8. EC-D1-GREEN.xlsx / .json
9. EC-X2-PINK.xlsx / .json
10. EC04-GREEN.xlsx / .json
11. EC04-NATURAL.xlsx / .json
12. EC404-2-4.xlsx / .json
13. ECSB-8-1.xlsx / .json
14. ST1122-4.xlsx / .json

**Recommended Action:** Replace "XX印刷厂" with actual supplier name from valid list (likely "凯源印刷", "深圳市众拓印刷", or "菲迪印刷")

---

### Issue Type B: Incorrect Company Suffix "有限公司" (4 files)

**Problem:** Supplier name includes "有限公司" but exact name in ERP is different

| File | Current Name (WRONG) | Likely Correct Name |
|------|---------------------|---------------------|
| B10-JMY802-BU.xlsx / .json | 宁波市艾盛电器有限公司 | 艾盛电器 |
| B10-JMY802-PK.xlsx / .json | 宁波市艾盛电器有限公司 | 艾盛电器 |
| B10-MJQ1-BU.xlsx / .json | 宁波市艾盛电器有限公司 | 艾盛电器 |
| B10-MJQ1-PK.xlsx / .json | 宁波市艾盛电器有限公司 | 艾盛电器 |

**Recommended Action:** Change to "艾盛电器" (exact match in ERP)

---

### Issue Type C: Incorrect Company Suffix "工贸有限公司" (3 files)

**Problem:** Supplier name includes "工贸有限公司" but exact name in ERP is different

| File | Current Name (WRONG) | Likely Correct Name |
|------|---------------------|---------------------|
| B10-MJB2-BK-1.xlsx / .json | 阳江骏业工贸有限公司 | 阳江骏业 |
| B10-MJB2-BK.xlsx / .json | 阳江骏业工贸有限公司 | 阳江骏业 |
| B10-MJB2-BK2.xlsx / .json | 阳江骏业工贸有限公司 | 阳江骏业 |

**Recommended Action:** Change to "阳江骏业" (exact match in ERP)

---

### Issue Type D: Missing Suffix "(已停用)" (2 files)

**Problem:** Supplier name is "吉秀制刷" but ERP has "吉秀制刷（已停用）"

| File | Current Name (WRONG) | Correct Name |
|------|---------------------|--------------|
| B10-ZMS3D-BR.xlsx / .json | 吉秀制刷 | 吉秀制刷（已停用） |
| B10-ZMS3D-NEW.xlsx / .json | 吉秀制刷 | 吉秀制刷（已停用） |

**Recommended Action:** Add "（已停用）" suffix OR verify if should use "宁波瑾秀制刷科技有限公司" instead

---

### Issue Type E: Different Company Name (1 file)

**Problem:** Supplier name doesn't exist in ERP

| File | Current Name (WRONG) | Possible Solution |
|------|---------------------|-------------------|
| FSC-WB01.xlsx / .json | 宁波泰友进出口有限公司 | Check if should be "宁波泰丰机械有限公司" |

**Recommended Action:** Verify correct supplier name with purchasing department

---

### Issue Type F: Placeholder "N/A" (3 files)

**Problem:** Using "N/A" as supplier name instead of actual supplier

| File | Current Name (WRONG) | Action Needed |
|------|---------------------|---------------|
| FSC-WB02.xlsx / .json | N/A | Fill in actual supplier name |
| NW-GRAY1.xlsx / .json | N/A | Fill in actual supplier name |
| NW-GRAY2.xlsx / .json | N/A | Fill in actual supplier name |

**Recommended Action:** Determine actual supplier and update

---

## Problem 2: Blank Supplier Names (29 Excel + 28 JSON files)

**Problem:** Cell B3 (供货商) is empty  
**Impact:** ERP will reject import due to missing required field

### EC-P3 Series (4 Excel + 4 JSON files)
- EC-P3-CYAN.xlsx / .json
- EC-P3-CYAN - 副本.xlsx (Excel only)
- EC-P3-GREEN.xlsx / .json
- EC-P3-NATURAL.xlsx / .json
- EC-P3-PINK.xlsx / .json

### EC-X2 Series (4 files each)
- EC-X2-CYAN.xlsx / .json
- EC-X2-GREEN.xlsx / .json
- EC-X2-NATURAL.xlsx / .json
- (EC-X2-PINK.xlsx / .json has "XX印刷厂" - see Problem 1)

### EC-D1 Series (2 files each)
- EC-D1-NATURAL.xlsx / .json
- EC-D1-PINK.xlsx / .json
- (EC-D1-CYAN.xlsx / .json has "XX印刷厂" - see Problem 1)
- (EC-D1-GREEN.xlsx / .json has "XX印刷厂" - see Problem 1)

### EC301 Series (8 files each)
- EC301-1-GR.xlsx / .json
- EC301-1-PK.xlsx / .json
- EC301-2-GR.xlsx / .json
- EC301-2-PK.xlsx / .json
- EC301-3-GR.xlsx / .json
- EC301-3-PK.xlsx / .json
- EC301-4-GR.xlsx / .json
- EC301-4-PK.xlsx / .json

### EC4xx Series (10 files each)
- EC401-3.xlsx / .json
- EC403.xlsx / .json
- EC403-3.xlsx / .json
- EC404-2-1-2.xlsx / .json
- EC404-2-2-2.xlsx / .json
- EC404-2-3-2.xlsx / .json
- EC404-2-4-2.xlsx / .json
- EC405-2-Grey.xlsx / .json
- EC405-2-Pink.xlsx / .json
- EC405-3-Blue.xlsx / .json

### AM Series (1 file each)
- AM413-2-1.xlsx / .json

**Recommended Action:** Fill in supplier name (cell B3) for all these files based on product/factory assignment

---

## Problem 3: JSON File Errors (1 file)

**File:** AM311.json  
**Error:** `Expecting value: line 2 column 3 (char 4)`  
**Impact:** File cannot be parsed, will cause system errors

**Recommended Action:** Fix JSON syntax error in AM311.json

---

## Action Plan

### Priority 1: Fix Blank Suppliers (29 Excel + 28 JSON)
These will cause immediate import failures. Assign correct suppliers to all files with blank B3 cells.

### Priority 2: Fix Generic Placeholders (14 files)
Replace "XX印刷厂" with actual supplier names.

### Priority 3: Fix Incorrect Names (10 files)
Correct supplier names with wrong suffixes or company structure:
- 宁波市艾盛电器有限公司 → 艾盛电器
- 阳江骏业工贸有限公司 → 阳江骏业
- 吉秀制刷 → 吉秀制刷（已停用）

### Priority 4: Fix Placeholders "N/A" (3 files)
Determine and assign actual suppliers for FSC-WB02, NW-GRAY1, NW-GRAY2.

### Priority 5: Fix JSON Error (1 file)
Repair syntax error in AM311.json.

### Priority 6: Verify FSC-WB01 (1 file)
Confirm correct supplier name for 宁波泰友进出口有限公司.

---

## How to Fix

### For Excel Templates:
1. Open file in Excel
2. Edit cell B3 (供货商) with exact supplier name from valid list
3. Save file
4. Run `excel_to_json_template.py` to sync to JSON

### For JSON Templates:
Option A (Recommended): Fix Excel first, then sync
Option B: Edit JSON directly - update `cells.B3.value` field

---

## Files Reference

- **Full Report:** `supplier_validation_report.json`
- **Script:** `check_supplier_names.py`
- **Reference File:** `docs/Supplier20260202170516-876874265653952512.xlsx`

---

**Report Generated:** 2026-02-02  
**Script Version:** 1.0
