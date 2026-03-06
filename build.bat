@echo off
chcp 65001 >nul
echo ============================================
echo   Build: Item Data Viewer  →  .exe
echo ============================================
echo.

:: Install dependencies first
echo [1/2] Installing Python dependencies ...
pip install -r requirements.txt --quiet

echo.
echo [2/2] Building executable with PyInstaller ...
pyinstaller ^
    --onefile ^
    --windowed ^
    --name "ItemDataViewer" ^
    --icon NONE ^
    --hidden-import psycopg2 ^
    --hidden-import pandas ^
    --hidden-import openpyxl ^
    --hidden-import xlrd ^
    app.py

echo.
if exist "dist\ItemDataViewer.exe" (
    echo  ✔  Build SUCCESS!
    echo  Output: dist\ItemDataViewer.exe
    explorer dist
) else (
    echo  ✘  Build FAILED — ตรวจสอบ error ด้านบน
)
pause
