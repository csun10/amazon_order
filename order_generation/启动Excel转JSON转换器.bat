@echo off
chcp 65001 >nul
echo 启动Excel ⇄ JSON 双向转换器...
echo.
echo 支持功能:
echo - Excel转JSON: 将Excel文件转换为JSON模板
echo - JSON转Excel: 将JSON模板转换为Excel文件  
echo - 批量处理: 自动处理文件夹中的所有文件
echo.

cd /d "%~dp0"

python excel_to_json_gui.py

echo.
pause