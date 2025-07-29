@echo off
echo ===============================================
echo      Excel Data Mapper - Build Script
echo ===============================================
echo.

REM Check if Python is available
echo [1/5] Checking Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found in PATH
    echo Please install Python 3.9+ and add it to PATH
    echo Download from: https://www.python.org/downloads/
    pause
    exit /b 1
)
python --version
echo.

REM Check if pip is available
echo [2/5] Checking pip installation...
python -m pip --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: pip not found
    echo Please reinstall Python with pip included
    pause
    exit /b 1
)
echo pip is available
echo.

REM Install requirements
echo [3/5] Installing requirements...
if exist requirements.txt (
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo ERROR: Failed to install requirements
        echo Please check your internet connection and try again
        pause
        exit /b 1
    )
    echo Requirements installed successfully
) else (
    echo WARNING: requirements.txt not found, installing basic packages
    python -m pip install ttkbootstrap==1.10.1 openpyxl==3.1.2 Pillow>=8.0.0
)
echo.

REM Install PyInstaller
echo [4/5] Installing PyInstaller...
python -m pip install pyinstaller>=5.0
if errorlevel 1 (
    echo ERROR: Failed to install PyInstaller
    pause
    exit /b 1
)
echo PyInstaller installed successfully
echo.

REM Create icon if not exists
if not exist icon.ico (
    echo Creating default icon...
    python -c "
from PIL import Image, ImageDraw
import sys
try:
    img = Image.new('RGBA', (64, 64), (0, 120, 215, 255))
    draw = ImageDraw.Draw(img)
    for i in range(0, 64, 8):
        draw.line([(i, 0), (i, 64)], fill=(255, 255, 255, 128), width=1)
        draw.line([(0, i), (64, i)], fill=(255, 255, 255, 128), width=1)
    draw.polygon([(20, 25), (30, 32), (20, 39)], fill=(255, 255, 255, 255))
    draw.polygon([(34, 25), (44, 32), (34, 39)], fill=(255, 255, 255, 255))
    img.save('icon.ico', format='ICO')
    print('Icon created successfully')
except Exception as e:
    print(f'Could not create icon: {e}')
    with open('icon.ico', 'wb') as f:
        f.write(b'\x00\x00\x01\x00\x01\x00\x10\x10\x00\x00\x01\x00\x08\x00h\x05\x00\x00\x16\x00\x00\x00')
        f.write(b'\x00' * (0x568 - 22))
    print('Created placeholder icon')
"
)

REM Build executable
echo [5/5] Building executable...
if not exist app.py (
    echo ERROR: app.py not found
    echo Please make sure you're in the correct directory
    pause
    exit /b 1
)

pyinstaller --onefile --windowed --name=ExcelDataMapper --icon=icon.ico app.py
if errorlevel 1 (
    echo ERROR: Build failed
    echo Please check the error messages above
    pause
    exit /b 1
)

REM Check if executable was created
if exist dist\ExcelDataMapper.exe (
    echo.
    echo ===============================================
    echo           BUILD COMPLETED SUCCESSFULLY!
    echo ===============================================
    echo.
    echo Executable location: dist\ExcelDataMapper.exe
    
    REM Get file size
    for %%A in (dist\ExcelDataMapper.exe) do (
        set /a "size=%%~zA/1024/1024"
        echo File size: !size! MB
    )
    
    echo.
    echo You can now:
    echo 1. Run the executable: dist\ExcelDataMapper.exe
    echo 2. Copy the executable to any Windows computer
    echo 3. No need to install Python on target computers
    echo.
    
    REM Ask if user wants to run the executable
    set /p "run=Do you want to run the executable now? (y/n): "
    if /i "!run!"=="y" (
        echo Starting Excel Data Mapper...
        start dist\ExcelDataMapper.exe
    )
    
) else (
    echo ERROR: Executable not found in expected location
    echo The build may have failed silently
    echo Please check the PyInstaller output above
)

echo.
echo Build script completed.
pause