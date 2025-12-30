@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo.
echo ================================================================================
echo 查看最近修改的Excel文件 (Check Recently Modified Excel Files)
echo ================================================================================
echo.
echo 此脚本将显示最近3天内修改过的Excel文件列表，但不会执行同步。
echo This script will show Excel files modified within the last 3 days without syncing.
echo.
echo ================================================================================
echo.

python -c "from auto_sync_excel_to_json import AutoSyncManager; from datetime import datetime, timedelta; manager = AutoSyncManager(3); files = manager.get_recently_modified_excel_files(); print(f'\n找到 {len(files)} 个文件:\n'); [print(f'{i+1:3d}. {f[0].name:40s} ({(datetime.now()-f[1]).days}天前)') for i, f in enumerate(files)]"

echo.
echo ================================================================================
echo 按任意键关闭窗口...
echo Press any key to close...
echo ================================================================================
pause >nul
