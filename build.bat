@echo off
echo ===============================================
echo     Building Faculty Attendance App (.exe)
echo ===============================================

REM --- Ensure current directory is this file's folder ---
cd /d "%~dp0"

echo Installing required packages...
python -m pip install --upgrade pip
python -m pip install pyinstaller python-dotenv pyodbc matplotlib numpy

echo.
echo Building .exe file (please wait)...
python -m PyInstaller --onefile --noconsole --add-data ".env:." "attendance_app.py"

if %errorlevel% neq 0 (
    echo ❌ Build failed. Check the error message above.
) else (
    echo ✅ Build complete!
    echo ===============================================
    echo The generated file is: dist\attendance_app.exe
    echo ===============================================
)

pause
