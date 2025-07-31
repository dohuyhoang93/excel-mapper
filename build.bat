@echo off
echo Building Excel Data Mapper...
echo.

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python not found in PATH
    pause
    exit /b 1
)

REM Install requirements
echo Installing requirements...
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo Error: Failed to install requirements
    pause
    exit /b 1
)

REM Install PyInstaller if not present
python -m pip install pyinstaller>=5.0

REM Build executable
echo Building executable...
python setup.py
if errorlevel 1 (
    echo Error: Build failed
    pause
    exit /b 1
)

echo.
echo Build completed successfully!
echo Executable location: dist\ExcelDataMapper.exe
echo.
pause
