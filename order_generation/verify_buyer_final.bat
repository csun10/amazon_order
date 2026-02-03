@echo off
chcp 65001 >nul
echo ================================================================
echo BUYER VERIFICATION - Elasticbrush01
echo ================================================================
echo.

cd /d "%~dp0"
python verify_buyer_final.py

pause
