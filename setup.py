"""
Setup script for building executable with PyInstaller
"""
import subprocess
import sys
import os
from pathlib import Path

def build_executable():
    """Build executable using PyInstaller"""
    
    # PyInstaller command
    cmd = [
        'pyinstaller',
        '--onefile',                    # Single executable file
        '--windowed',                   # No console window
        '--paths=.',                    # Find py file in this folder
        '--name=ExcelDataMapper',       # Executable name
        '--add-data=icon.ico;.',        # Include icon file
        '--icon=icon.ico',              # Set icon for executable
        '--distpath=dist',              # Output directory
        '--workpath=build',             # Work directory
        '--specpath=.',                 # Spec file location
        'app.py'                        # Main application file
    ]
    
    print("Building executable with PyInstaller...")
    print(f"Command: {' '.join(cmd)}")
    
    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("Build successful!")
        print(result.stdout)
        
        # Check if executable was created
        exe_path = Path("dist/ExcelDataMapper.exe")
        if exe_path.exists():
            print(f"Executable created: {exe_path.absolute()}")
            print(f"File size: {exe_path.stat().st_size / 1024 / 1024:.1f} MB")
        else:
            print("Warning: Executable not found in expected location")
            
    except subprocess.CalledProcessError as e:
        print(f"Build failed with error: {e}")
        print(f"Error output: {e.stderr}")
        return False
    except FileNotFoundError:
        print("PyInstaller not found. Install it with: pip install pyinstaller")
        return False
    
    return True

def create_icon():
    """Create a simple icon file if it doesn't exist"""
    icon_path = Path("icon.ico")
    if not icon_path.exists():
        print("Creating default icon file...")
        try:
            from PIL import Image, ImageDraw
            
            # Create a simple icon
            img = Image.new('RGBA', (64, 64), (0, 120, 215, 255))  # Blue background
            draw = ImageDraw.Draw(img)
            
            # Draw Excel-like grid
            for i in range(0, 64, 8):
                draw.line([(i, 0), (i, 64)], fill=(255, 255, 255, 128), width=1)
                draw.line([(0, i), (64, i)], fill=(255, 255, 255, 128), width=1)
            
            # Draw arrow
            draw.polygon([(20, 25), (30, 32), (20, 39)], fill=(255, 255, 255, 255))
            draw.polygon([(34, 25), (44, 32), (34, 39)], fill=(255, 255, 255, 255))
            
            img.save(icon_path, format='ICO')
            print(f"Icon created: {icon_path}")
            
        except ImportError:
            print("Pillow not available. Using text-based icon placeholder.")
            # Create a dummy icon file
            with open(icon_path, 'wb') as f:
                # Minimal ICO file header (won't work but prevents build errors)
                f.write(b'\x00\x00\x01\x00\x01\x00\x10\x10\x00\x00\x01\x00\x08\x00h\x05\x00\x00\x16\x00\x00\x00')
                f.write(b'\x00' * (0x568 - 22))  # Pad to minimum size

def install_requirements():
    """Install required packages"""
    requirements = [
        'ttkbootstrap==1.10.1',
        'openpyxl==3.1.2',
        'Pillow>=8.0.0',
        'pyinstaller>=5.0'
    ]
    
    print("Installing requirements...")
    for requirement in requirements:
        try:
            subprocess.run([sys.executable, '-m', 'pip', 'install', requirement], 
                          check=True, capture_output=True)
            print(f"✓ Installed {requirement}")
        except subprocess.CalledProcessError as e:
            print(f"✗ Failed to install {requirement}: {e}")
            return False
    
    return True

def create_build_script():
    """Create build.bat script for Windows"""
    build_script = """@echo off
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
echo Executable location: dist\\ExcelDataMapper.exe
echo.
pause
"""
    
    with open("build.bat", "w", encoding='utf-8') as f:
        f.write(build_script)
    
    print("Created build.bat script")

def main():
    """Main setup function"""
    print("Excel Data Mapper - Build Setup")
    print("=" * 40)
    
    # Create icon if needed
    create_icon()
    
    # Create build script
    create_build_script()
    
    # Check if this is being run directly (not imported)
    if len(sys.argv) > 1 and sys.argv[1] == 'build':
        # Install requirements
        if not install_requirements():
            return False
        
        # Build executable
        return build_executable()
    else:
        print("\nSetup completed. To build the executable:")
        print("1. Run: python setup.py build")
        print("2. Or run: build.bat (Windows)")
        print("3. Or manually: pyinstaller --onefile --windowed --icon=icon.ico app.py")
        return True

if __name__ == "__main__":
    success = main()
    if not success:
        sys.exit(1)