@echo off
cd /d "%~dp0"
echo.
echo Accessory Mapping Updater GUI
echo ============================
echo.
echo Starting the accessory mapping updater GUI application...
echo This tool helps update the accessory_mapping.json file using Excel data.
echo.

python accessory_mapping_updater_gui.py

echo.
echo GUI application closed.
pause
