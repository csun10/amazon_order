@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo.
echo ================================================================================
echo 自动同步：PO Excel → JSON 模板
echo Auto Sync: PO Excel Export → JSON Templates
echo ================================================================================
echo.
echo 此脚本将自动更新最近3天内修改过的Excel文件对应的JSON模板
echo This script will automatically update JSON templates for Excel files modified
echo within the last 3 days.
echo.
echo ================================================================================
echo.

python auto_sync_excel_to_json.py

echo.
echo ================================================================================
echo 按任意键关闭窗口...
echo Press any key to close...
echo ================================================================================
pause >nul
