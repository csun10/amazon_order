@echo off
set "SCRIPT_DIR=%~dp0"
echo 启动亚马逊订单产品搜索界面...
echo.
echo 功能:
echo - 按名称或SKU搜索产品
echo - 将多个产品添加到池中
echo - 为多个产品生成命令
echo - 从界面直接执行
echo.
cd /d "%SCRIPT_DIR%"
python product_search_gui.py
pause
