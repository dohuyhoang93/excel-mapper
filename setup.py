"""
Setup script for building the ExcelDataMapper executable using a .spec file.
This script provides a clean and reliable way to run PyInstaller.
"""
import subprocess
import sys
from pathlib import Path

# The name of the spec file which controls the build process
SPEC_FILE = "ExcelDataMapper.spec"

def build_with_spec():
    """
    Builds the executable using the .spec file with PyInstaller.
    Provides detailed output and error handling.
    """
    spec_path = Path(SPEC_FILE)
    if not spec_path.exists():
        print(f"Error: Build specification file '{SPEC_FILE}' not found!")
        print("Cannot proceed with the build.")
        return False

    # Command to run PyInstaller with the spec file
    # --clean: Removes temporary files before building
    cmd = [
        'pyinstaller',
        SPEC_FILE,
        '--distpath=dist',
        '--workpath=build',
        '--clean'
    ]
    
    print(f"Building executable with command: {' '.join(cmd)}")
    
    try:
        # Execute the command, ensuring output is captured in UTF-8
        result = subprocess.run(
            cmd, 
            check=True, 
            capture_output=True, 
            text=True, 
            encoding='utf-8'
        )
        print("Build successful!")
        print(result.stdout)
        
        # Verify that the executable was created
        # The name is taken from the 'name' field in the .spec file's EXE object
        exe_name = "ExcelDataMapper.exe"
        exe_path = Path("dist") / exe_name
        if exe_path.exists():
            print(f"\nExecutable created: {exe_path.resolve()}")
            print(f"File size: {exe_path.stat().st_size / 1024 / 1024:.2f} MB")
        else:
            print(f"Warning: Executable '{exe_name}' not found in 'dist' directory.")
            
    except subprocess.CalledProcessError as e:
        print(f"Build failed with error (return code {e.returncode}):")
        print("-" * 20 + " STDOUT " + "-" * 20)
        print(e.stdout)
        print("-" * 20 + " STDERR " + "-" * 20)
        print(e.stderr)
        return False
    except FileNotFoundError:
        print("Error: 'pyinstaller' command not found.")
        print("Please ensure PyInstaller is installed and in your system's PATH.")
        print("You can install it with: pip install -r requirements.txt")
        return False
        
    return True

def main():
    """Main entry point for the build script."""
    print("="*50)
    print("Excel Data Mapper - Build Script")
    print("="*50)
    
    success = build_with_spec()
    
    print("-" * 50)
    if success:
        print("Build process completed successfully.")
    else:
        print("Build process failed.")
        
    return success

if __name__ == "__main__":
    # If the script returns False (indicating failure), exit with a non-zero status code
    if not main():
        sys.exit(1)
