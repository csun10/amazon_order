# 自动同步工作流程说明 (Auto-Sync Workflow Guide)

## 概述 (Overview)

**单向工作流程 - 避免混淆！**
**One-Way Workflow - Avoid Confusion!**

系统设计为单向流程，避免在多个文件夹间产生混淆。

The system is designed as a one-way workflow to avoid confusion between multiple folders.

---

## 📁 文件夹说明 (Folder Structure)

```
order_generation/
├── json_template/              # ✏️ 主要编辑位置 (Primary editing location)
│   └── *.json                  # JSON模板源文件
│
├── PO_excel/                   # 📋 源Excel文件 (Source Excel files - EDITABLE)
│   └── *.xlsx                  # 可编辑的Excel文件，修改后自动同步到JSON
│
└── PO_excel_export/            # 📤 生成的Excel文件 (Generated Excel files - READ ONLY)
    └── *.xlsx                  # 从JSON生成的Excel输出文件（仅用于打印/发送）
```

### 文件夹用途 (Folder Purposes):

- **`PO_excel/`**: ✏️ **主要工作位置** - 在这里编辑Excel文件
- **`json_template/`**: 🔄 中间格式 - Excel自动转换为JSON（也可直接编辑）
- **`PO_excel_export/`**: 📤 最终输出 - 从JSON生成的Excel，用于打印/发送供应商

**工作流程简述:**
`PO_excel/` (编辑) → `json_template/` (转换) → `PO_excel_export/` (输出)

- **`PO_excel/`**: ✏️ **Primary workspace** - Edit Excel files here
- **`json_template/`**: 🔄 Intermediate format - Excel auto-converts to JSON (can also edit directly)
- **`PO_excel_export/`**: 📤 Final output - Excel generated from JSON for printing/sending to suppliers

**Workflow:**
`PO_excel/` (edit) → `json_template/` (convert) → `PO_excel_export/` (output)

---

## 🔄 工作流程 (Workflow)

### ✅ 标准流程 (Standard Workflow - RECOMMENDED)

```
1. 编辑 Excel 文件
   PO_excel/*.xlsx
        ↓
2. 转换为 JSON 模板
   [运行 启动Excel转JSON转换器.bat]
        ↓
3. JSON 模板
   json_template/*.json
        ↓
4. 生成输出 Excel
   [运行 json_PO_excel.py]
        ↓
5. Excel 输出文件
   PO_excel_export/*.xlsx
```

**操作步骤 (Steps):**
1. 在 `PO_excel/` 中编辑Excel文件（这是最简单的编辑方式）
2. 运行 `启动Excel转JSON转换器.bat` 将Excel转换为JSON模板
3. JSON自动保存到 `json_template/`
4. 运行 `json_PO_excel.py` 生成最终的Excel输出到 `PO_excel_export/`
5. 交货时间自动显示为 "45天", "15天", "7天" 格式

**为什么用这个流程？**
- Excel比JSON更容易编辑
- 自动同步检测 `PO_excel/` 中的修改（最近3天）
- `PO_excel_export/` 是最终输出，用于打印/发送供应商

---

### 🔧 高级用户 (Advanced - Direct JSON Editing)

如果你熟悉JSON格式，也可以直接编辑JSON模板：

```
1. 直接编辑 JSON
   json_template/*.json
        ↓
2. 生成输出 Excel
   [运行 json_PO_excel.py]
        ↓
3. Excel 输出文件
   PO_excel_export/*.xlsx
```

---

## 可用脚本 (Available Scripts)

### 1. `自动同步Excel到JSON.bat`
**功能:** 自动同步最近3天修改的Excel文件到JSON模板
**使用:** 双击运行即可

**Function:** Auto-sync Excel files modified within last 3 days to JSON templates
**Usage:** Double-click to run

---

### 2. `查看最近修改的文件.bat`
**功能:** 仅查看哪些文件会被同步，不执行同步操作
**使用:** 双击运行查看列表

**Function:** View which files would be synced without actually syncing
**Usage:** Double-click to view list

---

### 3. `auto_sync_excel_to_json.py`
**功能:** Python脚本，可配置同步天数
**使用:** 
```bash
python auto_sync_excel_to_json.py           # 默认3天
python auto_sync_excel_to_json.py --days 5  # 自定义5天
python auto_sync_excel_to_json.py --check-all  # 同步所有文件
```

**Function:** Python script with configurable days threshold
**Usage:**
```bash
python auto_sync_excel_to_json.py           # Default 3 days
python auto_sync_excel_to_json.py --days 5  # Custom 5 days
python auto_sync_excel_to_json.py --check-all  # Sync all files
```

---

## 重要注意事项 (Important Notes)

### ⚠️ 数据传输效率 (Data Transfer Efficiency)

- ✅ **建议:** 只在需要时运行自动同步
- ✅ **推荐:** 使用3天阈值避免不必要的同步
- ❌ **不建议:** 每次生成Excel后自动运行同步（会造成循环更新）

- ✅ **Recommended:** Only run auto-sync when needed
- ✅ **Best Practice:** Use 3-day threshold to avoid unnecessary syncing
- ❌ **Not Recommended:** Auto-run sync after every Excel generation (causes circular updates)

### 📋 交货时间格式 (Delivery Time Format)

- **JSON → Excel:** 保留 "45天" 格式（数字+天）
- **Excel → JSON:** 读取Excel中的实际值（可能是日期或文本）

- **JSON → Excel:** Keeps "45天" format (number + 天)
- **Excel → JSON:** Reads actual value from Excel (may be date or text)

### 🔄 最佳实践 (Best Practices)

1. **正常工作:** 主要在JSON模板中编辑，然后生成Excel
2. **特殊情况:** 如果需要在Excel中快速修改多个字段，可以使用自动同步
3. **定期检查:** 使用"查看最近修改的文件.bat"检查哪些文件会被同步

1. **Normal Work:** Primarily edit in JSON templates, then generate Excel
2. **Special Cases:** Use auto-sync if you need to quickly edit multiple fields in Excel
3. **Regular Check:** Use "查看最近修改的文件.bat" to check which files will be synced

---

## 常见问题 (FAQ)

**Q: 为什么不在每次生成Excel后自动同步？**
A: 这会造成不必要的数据传输和循环更新。只在手动编辑Excel后同步更高效。

**Q: Why not auto-sync after every Excel generation?**
A: This causes unnecessary data transfer and circular updates. Syncing only after manual edits is more efficient.

**Q: 如果我编辑了Excel文件，需要多久同步一次？**
A: 根据需要运行即可。脚本默认只同步3天内修改的文件。

**Q: How often should I sync after editing Excel files?**
A: Run as needed. Script defaults to syncing only files modified within 3 days.

**Q: 自动同步会覆盖我的JSON文件吗？**
A: 是的，它会用Excel中的数据更新对应的JSON文件。建议使用Git等版本控制系统。

**Q: Will auto-sync overwrite my JSON files?**
A: Yes, it updates corresponding JSON files with Excel data. Recommend using version control like Git.

---

## 文件夹结构 (Folder Structure)

```
order_generation/
├── json_template/              # JSON模板源文件
│   ├── AM311.json
│   ├── EC-5001.json
│   └── ...
├── PO_excel_export/           # 生成的Excel文件
│   ├── AM311.xlsx
│   ├── EC-5001.xlsx
│   └── ...
├── json_PO_excel.py          # JSON → Excel 转换器
├── auto_sync_excel_to_json.py # Excel → JSON 自动同步
├── 自动同步Excel到JSON.bat    # 同步批处理文件
└── 查看最近修改的文件.bat      # 查看文件批处理
```

---

## 版本历史 (Version History)

- **v1.0** (2025-12-30): 初始版本，支持3天自动同步
- **v1.0** (2025-12-30): Initial version with 3-day auto-sync support
