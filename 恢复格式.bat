@echo off
set "SCRIPT_DIR=%~dp0"
echo ==========================================
echo Excel Formatting Restoration Tool
echo ==========================================
echo.
echo This will restore Excel formatting while preserving your content changes.
echo.
pause

cd /d "%SCRIPT_DIR%"

echo Running Python restoration script...
python restore_excel_formatting.py

echo.
echo ==========================================
echo Restoration complete!
echo.
echo Your files with restored formatting are in:
echo %SCRIPT_DIR%\order_generation\PO_excel_restored\
echo ==========================================
pause
