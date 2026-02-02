@echo off
chcp 65001 >nul
echo ======================================================================
echo QUICK PIPELINE TEST
echo ======================================================================
echo.
echo Testing: Excel Template -^> JSON -^> Excel Output -^> PO Import
echo.

cd /d "%~dp0"

python quick_test_pipeline.py

pause
