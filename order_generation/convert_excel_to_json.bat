@echo off
set "SCRIPT_DIR=%~dp0"
echo Excel to JSON Template Converter
echo ================================
echo.
echo This script converts Excel files in the empty_base_template.xlsx format
echo to JSON template files for use with the order generation system.
echo.

cd /d "%SCRIPT_DIR%"

if "%~1"=="" (
    echo Usage: convert_excel_to_json.bat ^<excel_file^>
    echo        convert_excel_to_json.bat *.xlsx
    echo.
    echo Examples:
    echo   convert_excel_to_json.bat my_order.xlsx
    echo   convert_excel_to_json.bat docs\*.xlsx
    echo.
    pause
    exit /b 1
)

python excel_to_json_template.py %*

echo.
echo Conversion completed!
pause
