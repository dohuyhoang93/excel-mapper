# Setup script for building the ExcelDataMapper executable using a .spec file.
# This script provides a clean and reliable way to run PyInstaller.

import subprocess
import sys
import os
from pathlib import Path
import argparse

# The name of the spec file which controls the build process
SPEC_FILE = "ExcelDataMapper.spec"

def build_with_spec(one_file=False):
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
        '--clean',
        '--debug=all'
    ]

    # Use an environment variable to control the build mode in the .spec file
    env = os.environ.copy()
    if one_file:
        env['BUILD_MODE'] = 'ONEFILE'
    else:
        # Default to ONEDIR if --onefile is not specified
        env['BUILD_MODE'] = 'ONEDIR'

    print(f"Building executable with command: {' '.join(cmd)}")

    try:
        # Execute the command, ensuring output is captured in UTF-8
        result = subprocess.run(
            cmd,
            check=True,
            capture_output=True,
            text=True,
            encoding='utf-8',
            env=env  # Pass the modified environment
        )
        print("Build successful!")
        print(result.stdout)

        # Verify that the executable was created
        exe_name = "ExcelDataMapper.exe"
        if one_file:
            exe_path = Path("dist") / exe_name
        else:
            # The spec file defines the output folder name in the COLLECT call
            exe_path = Path("dist") / "ExcelDataMapper" / exe_name

        if exe_path.exists():
            print(f"\nExecutable created: {exe_path.resolve()}")
            print(f"File size: {exe_path.stat().st_size / 1024 / 1024:.2f} MB")
        else:
            print(f"Warning: Executable '{exe_name}' not found in the expected directory.")

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
    parser = argparse.ArgumentParser(description="Excel Data Mapper - Build Script")
    parser.add_argument(
        '--onefile',
        action='store_true',
        help='Build a single executable file instead of a directory.'
    )
    args = parser.parse_args()

    print("="*50)
    print("Excel Data Mapper - Build Script")
    print(f"Build mode: {'One-File' if args.onefile else 'One-Dir'}")
    print("="*50)

    success = build_with_spec(one_file=args.onefile)

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
