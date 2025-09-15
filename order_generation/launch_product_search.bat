@echo off
set "SCRIPT_DIR=%~dp0"
echo Starting Amazon Order Product Search GUI...
echo.
echo Features:
echo - Search products by name or SKU
echo - Add multiple products to pool
echo - Generate commands for multiple products
echo - Direct execution from GUI
echo.
cd /d "%SCRIPT_DIR%"
python product_search_gui.py
pause
