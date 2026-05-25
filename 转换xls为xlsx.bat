@echo off
:: 强制将 CMD 编码切换为 UTF-8，解决中文乱码问题
chcp 65001 >nul

:: 切换到项目根目录
cd /d D:\xxm\textile_app

echo Checking Python version...
python --version
echo.

echo [1/2] Convert Delivery Notes...
python tools\convert_xls_to_xlsx.py "D:\xxm\送货单"
echo.

echo [2/2] Convert Templates...
python tools\convert_xls_to_xlsx.py "D:\xxm\templates"
echo.

echo Done!
pause