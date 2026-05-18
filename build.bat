@echo off
cd /d "%USERPROFILE%\Desktop\FCGGrow"

echo.
echo ============================================
echo   FCG Grow - PyInstaller Build
echo   Fairy Circle Garden
echo ============================================
echo.

echo Checking for PyInstaller...
python -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo PyInstaller not found. Installing...
    pip install pyinstaller
)

echo.
echo Building FCGGrow.exe...
echo.

python -m PyInstaller --onefile --windowed --icon=fcggrow.ico --name=FCGGrow FCGGrow.pyw

echo.
if exist "dist\FCGGrow.exe" (
    echo ============================================
    echo   BUILD SUCCESSFUL
    echo   Your exe is at: dist\FCGGrow.exe
    echo   Copy FCGGrow.exe and fcggrow_help.html
    echo   into one folder to distribute.
    echo ============================================
) else (
    echo ============================================
    echo   BUILD FAILED
    echo   Check the output above for errors.
    echo ============================================
)

echo.
pause
