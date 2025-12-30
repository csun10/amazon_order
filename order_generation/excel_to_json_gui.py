#!/usr/bin/env python3
"""
Enhanced Excel to JSON Template GUI

This GUI provides two main modes:
1. Manual Mode: Select specific Excel files to convert to specific JSON templates
2. Batch Mode: Process all Excel files in PO_excel folder and match to JSON templates

Features:
- Browse and select individual Excel files
- Match Excel files to existing JSON templates
- Batch process all Excel files in PO_excel directory
- Real-time progress tracking
- Detailed logging and error reporting
- Background processing to keep UI responsive
"""

import json
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, simpledialog
import threading

# Import the converter from the existing script
try:
    # Import all required modules first
    import openpyxl
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    from openpyxl.cell.rich_text import CellRichText
    
    # Now import the converter class
    from excel_to_json_template import ExcelToJsonConverter
except ImportError as e:
    print(f"Error: Required module not found: {e}")
    print("Please ensure excel_to_json_template.py and openpyxl are available")
    sys.exit(1)


class EnhancedExcelToJsonGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel ⇄ JSON 双向转换器")
        self.root.geometry("1000x800")
        
        # Initialize converter
        self.converter = ExcelToJsonConverter()
        
        # Selected files for manual mode
        self.manual_pairs = []  # List of (excel_path, json_name) tuples
        self.json_to_excel_pairs = []  # List of (json_path, excel_name) tuples
        
        # Current mode
        self.current_mode = tk.StringVar(value="excel_to_json")
        
        # Create GUI
        self._create_widgets()
        self._scan_files()
        
    def _create_widgets(self):
        """Create all GUI widgets"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Excel ⇄ JSON 双向转换器", 
                               font=("TkDefaultFont", 16, "bold"))
        title_label.grid(row=0, column=0, pady=(0, 20))
        
        # Mode selection
        mode_frame = ttk.LabelFrame(main_frame, text="转换模式", padding="10")
        mode_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Radiobutton(mode_frame, text="Excel → JSON：Excel文件转换为JSON模板", 
                       variable=self.current_mode, value="excel_to_json", 
                       command=self._mode_changed).pack(anchor=tk.W, pady=2)
        ttk.Radiobutton(mode_frame, text="JSON → Excel：JSON模板转换为Excel文件", 
                       variable=self.current_mode, value="json_to_excel", 
                       command=self._mode_changed).pack(anchor=tk.W, pady=2)
        ttk.Radiobutton(mode_frame, text="批量处理：批量转换文件夹中的所有文件", 
                       variable=self.current_mode, value="batch", 
                       command=self._mode_changed).pack(anchor=tk.W, pady=2)
        
        # Create notebook for different modes
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # Excel to JSON tab
        self.excel_to_json_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.excel_to_json_frame, text="Excel → JSON")
        self._create_excel_to_json_widgets()
        
        # JSON to Excel tab
        self.json_to_excel_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.json_to_excel_frame, text="JSON → Excel")
        self._create_json_to_excel_widgets()
        
        # Batch mode tab  
        self.batch_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.batch_frame, text="批量处理")
        self._create_batch_mode_widgets()
        
        # Progress and output section
        output_frame = ttk.LabelFrame(main_frame, text="转换进度与日志", padding="10")
        output_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        output_frame.columnconfigure(0, weight=1)
        output_frame.rowconfigure(2, weight=1)
        
        # Progress bar
        self.progress_var = tk.StringVar(value="准备转换文件...")
        progress_label = ttk.Label(output_frame, textvariable=self.progress_var)
        progress_label.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        self.progress_bar = ttk.Progressbar(output_frame, mode='determinate')
        self.progress_bar.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Output text
        self.output_text = scrolledtext.ScrolledText(output_frame, height=15, wrap=tk.WORD)
        self.output_text.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Status bar
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # Initial message
        self._log("Excel ⇄ JSON 双向转换器准备就绪！")
        self._log("选择转换模式开始转换。")
        self._log(f"JSON模板目录: {self.converter.template_dir}")
        self._log(f"PO Excel目录: {self.converter.root_dir / 'PO_excel'}")
        
    def _create_excel_to_json_widgets(self):
        """Create widgets for Excel to JSON conversion"""
        self.excel_to_json_frame.columnconfigure(0, weight=1)
        self.excel_to_json_frame.rowconfigure(1, weight=1)
        
        # Instructions
        inst_label = ttk.Label(self.excel_to_json_frame, 
                              text="选择Excel文件并指定要更新的JSON模板文件名（不含.json扩展名）")
        inst_label.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # File pairs list
        pairs_frame = ttk.LabelFrame(self.excel_to_json_frame, text="文件配对", padding="10")
        pairs_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        pairs_frame.columnconfigure(0, weight=1)
        pairs_frame.rowconfigure(0, weight=1)
        
        # Create treeview for file pairs
        self.excel_pairs_tree = ttk.Treeview(pairs_frame, columns=('excel', 'json'), show='headings', height=10)
        self.excel_pairs_tree.heading('#1', text='Excel文件')
        self.excel_pairs_tree.heading('#2', text='JSON模板名')
        self.excel_pairs_tree.column('#1', width=400)
        self.excel_pairs_tree.column('#2', width=300)
        
        excel_pairs_scrollbar = ttk.Scrollbar(pairs_frame, orient=tk.VERTICAL, command=self.excel_pairs_tree.yview)
        self.excel_pairs_tree.configure(yscrollcommand=excel_pairs_scrollbar.set)
        
        self.excel_pairs_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        excel_pairs_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Control buttons for Excel to JSON
        excel_buttons_frame = ttk.Frame(self.excel_to_json_frame)
        excel_buttons_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        
        ttk.Button(excel_buttons_frame, text="添加Excel文件", 
                  command=self._add_excel_file_pair).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(excel_buttons_frame, text="移除选中配对", 
                  command=self._remove_selected_excel_pair).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(excel_buttons_frame, text="清空全部", 
                  command=self._clear_excel_pairs).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(excel_buttons_frame, text="转换为JSON", 
                  command=self._convert_excel_to_json).pack(side=tk.LEFT, padx=(20, 0))
    
    def _create_json_to_excel_widgets(self):
        """Create widgets for JSON to Excel conversion"""
        self.json_to_excel_frame.columnconfigure(0, weight=1)
        self.json_to_excel_frame.rowconfigure(1, weight=1)
        
        # Instructions
        inst_label = ttk.Label(self.json_to_excel_frame, 
                              text="选择JSON模板文件并指定要生成的Excel文件名（不含.xlsx扩展名）")
        inst_label.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # File pairs list
        json_pairs_frame = ttk.LabelFrame(self.json_to_excel_frame, text="文件配对", padding="10")
        json_pairs_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        json_pairs_frame.columnconfigure(0, weight=1)
        json_pairs_frame.rowconfigure(0, weight=1)
        
        # Create treeview for JSON file pairs
        self.json_pairs_tree = ttk.Treeview(json_pairs_frame, columns=('json', 'excel'), show='headings', height=10)
        self.json_pairs_tree.heading('#1', text='JSON模板文件')
        self.json_pairs_tree.heading('#2', text='Excel文件名')
        self.json_pairs_tree.column('#1', width=400)
        self.json_pairs_tree.column('#2', width=300)
        
        json_pairs_scrollbar = ttk.Scrollbar(json_pairs_frame, orient=tk.VERTICAL, command=self.json_pairs_tree.yview)
        self.json_pairs_tree.configure(yscrollcommand=json_pairs_scrollbar.set)
        
        self.json_pairs_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        json_pairs_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Control buttons for JSON to Excel
        json_buttons_frame = ttk.Frame(self.json_to_excel_frame)
        json_buttons_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        
        ttk.Button(json_buttons_frame, text="添加JSON文件", 
                  command=self._add_json_file_pair).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(json_buttons_frame, text="移除选中配对", 
                  command=self._remove_selected_json_pair).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(json_buttons_frame, text="清空全部", 
                  command=self._clear_json_pairs).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(json_buttons_frame, text="转换为Excel", 
                  command=self._convert_json_to_excel).pack(side=tk.LEFT, padx=(20, 0))
        
    def _create_batch_mode_widgets(self):
        """Create widgets for batch mode"""
        self.batch_frame.columnconfigure(0, weight=1)
        self.batch_frame.rowconfigure(1, weight=1)
        
        # Instructions
        inst_label = ttk.Label(self.batch_frame, 
                              text="批量处理PO_excel文件夹中的所有Excel文件，自动匹配对应的JSON模板")
        inst_label.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # File matching preview
        preview_frame = ttk.LabelFrame(self.batch_frame, text="文件匹配预览", padding="10")
        preview_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        
        # Create treeview for file matching
        self.match_tree = ttk.Treeview(preview_frame, columns=('excel', 'json', 'status'), show='headings', height=15)
        self.match_tree.heading('#1', text='Excel文件')
        self.match_tree.heading('#2', text='JSON模板')
        self.match_tree.heading('#3', text='状态')
        self.match_tree.column('#1', width=350)
        self.match_tree.column('#2', width=350)
        self.match_tree.column('#3', width=150)
        
        match_scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.match_tree.yview)
        self.match_tree.configure(yscrollcommand=match_scrollbar.set)
        
        self.match_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        match_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Control buttons for batch mode
        batch_buttons_frame = ttk.Frame(self.batch_frame)
        batch_buttons_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        
        ttk.Button(batch_buttons_frame, text="刷新文件列表", 
                  command=self._scan_files).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(batch_buttons_frame, text="转换全部匹配文件", 
                  command=self._convert_batch).pack(side=tk.LEFT, padx=(20, 0))
        
    def _mode_changed(self):
        """Handle mode change"""
        mode = self.current_mode.get()
        if mode == "excel_to_json":
            self.notebook.select(0)  # Excel to JSON tab
        elif mode == "json_to_excel":
            self.notebook.select(1)  # JSON to Excel tab
        else:  # batch mode
            self.notebook.select(2)  # Batch tab
            self._scan_files()  # Refresh file list when switching to batch mode
    
    def _scan_files(self):
        """Scan for Excel files and existing JSON templates"""
        po_excel_dir = self.converter.root_dir / "PO_excel"
        json_template_dir = self.converter.template_dir
        
        if not po_excel_dir.exists():
            self._log(f"警告: PO_excel 目录不存在: {po_excel_dir}")
            return
        
        # Get all Excel files
        excel_files = list(po_excel_dir.glob("*.xlsx")) + list(po_excel_dir.glob("*.xls"))
        
        # Get all existing JSON templates
        json_files = {f.stem: f for f in json_template_dir.glob("*.json")}
        
        # Clear the tree
        for item in self.match_tree.get_children():
            self.match_tree.delete(item)
        
        # Populate the tree with matching information
        matched_count = 0
        for excel_file in sorted(excel_files):
            excel_name = excel_file.stem  # Name without extension
            
            if excel_name in json_files:
                status = "[OK] 已匹配"
                matched_count += 1
                tags = ('matched',)
            else:
                status = "[X] 未匹配"
                tags = ('unmatched',)
            
            json_name = f"{excel_name}.json" if excel_name in json_files else "无对应JSON"
            
            self.match_tree.insert('', 'end', values=(excel_file.name, json_name, status), tags=tags)
        
        # Configure tags
        self.match_tree.tag_configure('matched', background='lightgreen')
        self.match_tree.tag_configure('unmatched', background='lightcoral')
        
        total_files = len(excel_files)
        self._log(f"扫描完成: 找到 {total_files} 个Excel文件，{matched_count} 个已匹配JSON模板")
        self.status_var.set(f"扫描完成: {matched_count}/{total_files} 个文件已匹配")
        
    # Excel to JSON Methods
    def _add_excel_file_pair(self):
        """Add a new Excel-JSON file pair"""
        # Select Excel file
        excel_file = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
            initialdir=self.converter.root_dir / "PO_excel"
        )
        
        if not excel_file:
            return
        
        # Get JSON template name
        excel_name = Path(excel_file).stem
        json_name = simpledialog.askstring(
            "JSON模板名称", 
            f"请输入JSON模板名称（不含.json扩展名）:\n\n建议使用: {excel_name}",
            initialvalue=excel_name
        )
        
        if not json_name:
            return
        
        # Add to pairs list
        pair = (excel_file, json_name)
        if pair not in self.manual_pairs:
            self.manual_pairs.append(pair)
            self.excel_pairs_tree.insert('', 'end', values=(Path(excel_file).name, f"{json_name}.json"))
            self._log(f"添加Excel转JSON配对: {Path(excel_file).name} -> {json_name}.json")
        else:
            messagebox.showwarning("警告", "该配对已存在")
    
    def _remove_selected_excel_pair(self):
        """Remove selected Excel file pair"""
        selection = self.excel_pairs_tree.selection()
        if selection:
            for item in selection:
                index = self.excel_pairs_tree.index(item)
                removed_pair = self.manual_pairs.pop(index)
                self.excel_pairs_tree.delete(item)
                self._log(f"移除Excel转JSON配对: {Path(removed_pair[0]).name} -> {removed_pair[1]}.json")
    
    def _clear_excel_pairs(self):
        """Clear all Excel file pairs"""
        self.manual_pairs.clear()
        for item in self.excel_pairs_tree.get_children():
            self.excel_pairs_tree.delete(item)
        self._log("清空了所有Excel转JSON配对")
    
    def _convert_excel_to_json(self):
        """Convert Excel files to JSON"""
        if not self.manual_pairs:
            messagebox.showwarning("警告", "请添加要转换的Excel文件配对")
            return
        
        # Start conversion in background thread
        self._start_excel_to_json_conversion(self.manual_pairs)
    
    # JSON to Excel Methods
    def _add_json_file_pair(self):
        """Add a new JSON-Excel file pair"""
        # Select JSON file
        json_file = filedialog.askopenfilename(
            title="选择JSON模板文件",
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")],
            initialdir=self.converter.template_dir
        )
        
        if not json_file:
            return
        
        # Get Excel file name
        json_name = Path(json_file).stem
        excel_name = simpledialog.askstring(
            "Excel文件名称", 
            f"请输入Excel文件名称（不含.xlsx扩展名）:\n\n建议使用: {json_name}",
            initialvalue=json_name
        )
        
        if not excel_name:
            return
        
        # Add to pairs list
        pair = (json_file, excel_name)
        if pair not in self.json_to_excel_pairs:
            self.json_to_excel_pairs.append(pair)
            self.json_pairs_tree.insert('', 'end', values=(Path(json_file).name, f"{excel_name}.xlsx"))
            self._log(f"添加JSON转Excel配对: {Path(json_file).name} -> {excel_name}.xlsx")
        else:
            messagebox.showwarning("警告", "该配对已存在")
    
    def _remove_selected_json_pair(self):
        """Remove selected JSON file pair"""
        selection = self.json_pairs_tree.selection()
        if selection:
            for item in selection:
                index = self.json_pairs_tree.index(item)
                removed_pair = self.json_to_excel_pairs.pop(index)
                self.json_pairs_tree.delete(item)
                self._log(f"移除JSON转Excel配对: {Path(removed_pair[0]).name} -> {removed_pair[1]}.xlsx")
    
    def _clear_json_pairs(self):
        """Clear all JSON file pairs"""
        self.json_to_excel_pairs.clear()
        for item in self.json_pairs_tree.get_children():
            self.json_pairs_tree.delete(item)
        self._log("清空了所有JSON转Excel配对")
    
    def _convert_json_to_excel(self):
        """Convert JSON files to Excel"""
        if not self.json_to_excel_pairs:
            messagebox.showwarning("警告", "请添加要转换的JSON文件配对")
            return
        
        # Start conversion in background thread
        self._start_json_to_excel_conversion(self.json_to_excel_pairs)

    def _convert_batch(self):
        """Convert all matched files in batch mode"""
        po_excel_dir = self.converter.root_dir / "PO_excel"
        json_template_dir = self.converter.template_dir
        
        # Get all Excel files and find matches
        excel_files = list(po_excel_dir.glob("*.xlsx")) + list(po_excel_dir.glob("*.xls"))
        json_files = {f.stem: f for f in json_template_dir.glob("*.json")}
        
        # Build pairs for matched files
        batch_pairs = []
        for excel_file in excel_files:
            excel_name = excel_file.stem
            if excel_name in json_files:
                batch_pairs.append((str(excel_file), excel_name))
        
        if not batch_pairs:
            messagebox.showwarning("警告", "没有找到匹配的Excel和JSON文件对")
            return
        
        if messagebox.askyesno("确认", f"将转换 {len(batch_pairs)} 个匹配的文件对。是否继续？"):
            self._start_conversion(batch_pairs, "batch")
    
    def _start_excel_to_json_conversion(self, pairs: List[Tuple[str, str]]):
        """Start Excel to JSON conversion in background thread"""
        # Disable UI during processing
        self._set_ui_state(False)
        
        # Start conversion worker
        thread = threading.Thread(target=self._excel_to_json_worker, args=(pairs,))
        thread.daemon = True
        thread.start()
    
    def _start_json_to_excel_conversion(self, pairs: List[Tuple[str, str]]):
        """Start JSON to Excel conversion in background thread"""
        # Disable UI during processing
        self._set_ui_state(False)
        
        # Start conversion worker
        thread = threading.Thread(target=self._json_to_excel_worker, args=(pairs,))
        thread.daemon = True
        thread.start()
    
    def _excel_to_json_worker(self, pairs: List[Tuple[str, str]]):
        """Background worker for Excel to JSON conversion"""
        try:
            total_pairs = len(pairs)
            successful_conversions = 0
            total_generated = 0
            
            self.progress_bar.configure(maximum=total_pairs)
            
            self._log(f"\n开始Excel转JSON转换，共 {total_pairs} 个文件配对")
            self._log("="*50)
            
            for i, (excel_path, json_name) in enumerate(pairs):
                self.progress_var.set(f"处理中 {i+1}/{total_pairs}: {Path(excel_path).name}")
                self.progress_bar.configure(value=i)
                
                self._log(f"\n[{i+1}/{total_pairs}] 转换: {Path(excel_path).name} -> {json_name}.json")
                
                try:
                    # Convert Excel to JSON templates
                    generated_files = self.converter.convert_excel_to_json(Path(excel_path))
                    
                    if generated_files:
                        # Find the specific JSON file we want to update
                        target_json = self.converter.template_dir / f"{json_name}.json"
                        
                        # Check if our target was generated
                        generated_names = [f.stem for f in generated_files]
                        if json_name in generated_names:
                            successful_conversions += 1
                            total_generated += len(generated_files)
                            self._log(f"  [OK] 成功更新: {json_name}.json")
                            if len(generated_files) > 1:
                                self._log(f"  [i] 同时生成了 {len(generated_files)} 个JSON文件:")
                                for gf in generated_files:
                                    self._log(f"    - {gf.name}")
                        else:
                            self._log(f"  [!] Excel文件生成了JSON，但未包含目标文件: {json_name}.json")
                            self._log(f"  [i] 实际生成的文件:")
                            for gf in generated_files:
                                self._log(f"    - {gf.name}")
                            total_generated += len(generated_files)
                    else:
                        self._log(f"  ✗ 未生成任何JSON文件（Excel可能无有效数据）")
                        
                except Exception as e:
                    self._log(f"  ✗ 转换失败: {str(e)}")
                
                # Small delay to update UI
                self.root.after(10)
            
            # Final progress update
            self.progress_bar.configure(value=total_pairs)
            self.progress_var.set("转换完成")
            
            self._log("="*50)
            self._log(f"Excel转JSON转换完成!")
            self._log(f"成功转换: {successful_conversions}/{total_pairs} 个配对")
            self._log(f"生成JSON文件总数: {total_generated}")
            self.status_var.set(f"Excel转JSON完成: {successful_conversions}/{total_pairs} 成功")
            
        except Exception as e:
            self._log(f"转换过程出错: {str(e)}")
            self.status_var.set("转换失败")
        finally:
            # Re-enable UI
            self.root.after(100, lambda: self._set_ui_state(True))
    
    def _json_to_excel_worker(self, pairs: List[Tuple[str, str]]):
        """Background worker for JSON to Excel conversion"""
        try:
            total_pairs = len(pairs)
            successful_conversions = 0
            
            self.progress_bar.configure(maximum=total_pairs)
            
            self._log(f"\n开始JSON转Excel转换，共 {total_pairs} 个文件配对")
            self._log("="*50)
            
            for i, (json_path, excel_name) in enumerate(pairs):
                self.progress_var.set(f"处理中 {i+1}/{total_pairs}: {Path(json_path).name}")
                self.progress_bar.configure(value=i)
                
                self._log(f"\n[{i+1}/{total_pairs}] 转换: {Path(json_path).name} -> {excel_name}.xlsx")
                
                try:
                    # Use json_PO_excel.py script to convert JSON to Excel
                    import subprocess
                    import sys
                    
                    json_po_excel_script = self.converter.root_dir / "json_PO_excel.py"
                    output_dir = self.converter.root_dir / "PO_excel"
                    output_file = output_dir / f"{excel_name}.xlsx"
                    
                    # Run the JSON to Excel conversion
                    cmd = [sys.executable, str(json_po_excel_script), str(json_path), str(output_file)]
                    result = subprocess.run(cmd, capture_output=True, text=True, cwd=str(self.converter.root_dir))
                    
                    if result.returncode == 0:
                        if output_file.exists():
                            successful_conversions += 1
                            self._log(f"  [OK] 成功生成: {excel_name}.xlsx")
                        else:
                            self._log(f"  [!] 脚本执行成功但文件未生成")
                    else:
                        self._log(f"  [FAIL] 转换失败: {result.stderr}")
                        
                except Exception as e:
                    self._log(f"  [FAIL] 转换失败: {str(e)}")
                
                # Small delay to update UI
                self.root.after(10)
            
            # Final progress update
            self.progress_bar.configure(value=total_pairs)
            self.progress_var.set("转换完成")
            
            self._log("="*50)
            self._log(f"JSON转Excel转换完成!")
            self._log(f"成功转换: {successful_conversions}/{total_pairs} 个配对")
            self.status_var.set(f"JSON转Excel完成: {successful_conversions}/{total_pairs} 成功")
            
        except Exception as e:
            self._log(f"转换过程出错: {str(e)}")
            self.status_var.set("转换失败")
        finally:
            # Re-enable UI
            self.root.after(100, lambda: self._set_ui_state(True))

    def _start_conversion(self, pairs: List[Tuple[str, str]], mode: str):
        """Start conversion in background thread"""
        # Disable UI during processing
        self._set_ui_state(False)
        
        # Start conversion worker
        thread = threading.Thread(target=self._conversion_worker, args=(pairs, mode))
        thread.daemon = True
        thread.start()
    
    def _set_ui_state(self, enabled: bool):
        """Enable or disable UI elements"""
        state = 'normal' if enabled else 'disabled'
        
        def set_widget_state(widget):
            try:
                widget.configure(state=state)
            except:
                pass
            for child in widget.winfo_children():
                set_widget_state(child)
        
        set_widget_state(self.root)
        
        # Keep output text always enabled
        self.output_text.configure(state='normal')
    
    def _conversion_worker(self, pairs: List[Tuple[str, str]], mode: str):
        """Background worker for file conversion"""
        try:
            total_pairs = len(pairs)
            successful_conversions = 0
            total_generated = 0
            
            self.progress_bar.configure(maximum=total_pairs)
            
            self._log(f"\n开始{mode}转换模式，共 {total_pairs} 个文件配对")
            self._log("="*50)
            
            for i, (excel_path, json_name) in enumerate(pairs):
                self.progress_var.set(f"处理中 {i+1}/{total_pairs}: {Path(excel_path).name}")
                self.progress_bar.configure(value=i)
                
                self._log(f"\n[{i+1}/{total_pairs}] 转换: {Path(excel_path).name} -> {json_name}.json")
                
                try:
                    # Convert Excel to JSON templates
                    generated_files = self.converter.convert_excel_to_json(Path(excel_path))
                    
                    if generated_files:
                        # Find the specific JSON file we want to update
                        target_json = self.converter.template_dir / f"{json_name}.json"
                        
                        # Check if our target was generated
                        generated_names = [f.stem for f in generated_files]
                        if json_name in generated_names:
                            successful_conversions += 1
                            total_generated += len(generated_files)
                            self._log(f"  [OK] 成功更新: {json_name}.json")
                            if len(generated_files) > 1:
                                self._log(f"  [i] 同时生成了 {len(generated_files)} 个JSON文件:")
                                for gf in generated_files:
                                    self._log(f"    - {gf.name}")
                        else:
                            self._log(f"  [!] 未生成目标JSON文件: {json_name}.json")
                            self._log(f"  [i] 实际生成的文件: {', '.join(generated_names) if generated_names else '无'}")
                            total_generated += len(generated_files)
                    else:
                        self._log(f"  ✗ 未生成任何JSON文件（Excel可能无有效数据）")
                        
                except Exception as e:
                    self._log(f"  ✗ 转换错误: {e}")
            
            self.progress_bar.configure(value=total_pairs)
            self.progress_var.set("转换完成！")
            
            # Summary
            self._log(f"\n" + "="*50)
            self._log(f"转换总结 ({mode}模式)")
            self._log(f"="*50)
            self._log(f"处理的配对: {total_pairs}")
            self._log(f"成功转换: {successful_conversions}")
            self._log(f"生成的JSON文件: {total_generated}")
            self._log(f"输出目录: {self.converter.template_dir}")
            self._log(f"="*50)
            
            if successful_conversions > 0:
                messagebox.showinfo("转换完成", 
                    f"{mode}模式转换完成！\n\n"
                    f"处理的配对: {total_pairs}\n"
                    f"成功转换: {successful_conversions}\n"
                    f"生成的JSON文件: {total_generated}")
            else:
                messagebox.showwarning("转换完成", 
                    f"{mode}模式转换完成，但没有成功的转换。\n"
                    f"请检查Excel文件格式和数据。")
            
            self.status_var.set(f"完成: {successful_conversions}/{total_pairs} 成功转换")
            
            # Refresh file list if in batch mode
            if mode == "batch":
                self._scan_files()
                
        except Exception as e:
            self._log(f"\n转换过程发生意外错误: {e}")
            messagebox.showerror("错误", f"转换失败: {e}")
            self.status_var.set("发生错误")
        
        finally:
            # Re-enable UI
            self._set_ui_state(True)
    
    def _log(self, message: str):
        """Log a message to the output text"""
        self.output_text.insert(tk.END, message + "\n")
        self.output_text.see(tk.END)
        self.root.update()


def main():
    try:
        root = tk.Tk()
        app = EnhancedExcelToJsonGUI(root)
        root.mainloop()
    except Exception as e:
        print(f"GUI Error: {e}")
        print("Please ensure excel_to_json_template.py is in the same directory")
        sys.exit(1)


if __name__ == "__main__":
    main()