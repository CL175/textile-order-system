@echo off
cd /d "%~dp0"

echo ==========================================
echo   Build EXE
echo ==========================================
echo.

set PY=
if exist "C:\Python38\python.exe" set PY=C:\Python38\python.exe
if "%PY%"=="" set PY=python

echo Python: %PY%
%PY% --version
echo.

REM Clean old build
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"

%PY% -c "import PyInstaller" >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing PyInstaller...
    %PY% -m pip install pyinstaller
    echo.
)

echo Building...
%PY% -m PyInstaller --onefile --windowed --name fangzhi --add-data "data;data" --hidden-import win32com --hidden-import win32com.client --hidden-import win32gui --hidden-import win32con --hidden-import pywinauto --hidden-import openpyxl --hidden-import xlrd --hidden-import xlwt --hidden-import xlutils --hidden-import xlutils.copy --hidden-import pythoncom main.py

if %errorlevel% equ 0 (
    if exist "dist\fangzhi.exe" move /y "dist\fangzhi.exe" "dist\fangzhi_app.exe" >nul 2>&1
    echo.
    echo ==========================================
    echo   Build OK!
    echo   EXE: dist\fangzhi_app.exe
    echo ==========================================
) else (
    echo.
    echo Build FAILED!
)
pause
