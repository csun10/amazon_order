@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo.
echo ================================================================================
echo 智能同步助手 (Smart Sync Helper)
echo ================================================================================
echo.

REM First, show what files would be synced
python -c "from auto_sync_excel_to_json import AutoSyncManager; from datetime import datetime, timedelta; manager = AutoSyncManager(3); files = manager.get_recently_modified_excel_files(); print(f'\n发现 {len(files)} 个文件在最近3天内被修改:\n'); [print(f'  {i+1:2d}. {f[0].name:40s} (修改于 {f[1].strftime(\"%%Y-%%m-%%d %%H:%%M\")})') for i, f in enumerate(files)] if files else print('  没有找到最近修改的文件')"

echo.
echo ================================================================================
echo.

REM Ask user if they want to proceed
set /p choice="是否要同步这些文件到JSON模板? (Y/N): "

if /i "%choice%"=="Y" (
    echo.
    echo 开始同步...
    echo.
    python auto_sync_excel_to_json.py
) else (
    echo.
    echo 已取消同步。
    echo.
)

echo.
echo ================================================================================
echo 按任意键关闭窗口...
echo Press any key to close...
echo ================================================================================
pause >nul
