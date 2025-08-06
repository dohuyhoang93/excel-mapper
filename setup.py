<<<<<<< HEAD
# setup.py
import sys
import os
import shutil
import subprocess
from setuptools import setup, find_packages
from setuptools.command.build_py import build_py
=======
# Setup script for building the ExcelDataMapper executable using a .spec file.
# This script provides a clean and reliable way to run PyInstaller.

import subprocess
import sys
import os
from pathlib import Path
import argparse
>>>>>>> 4b8cbb364fe49caa4def35363b0714671af8f6c7

# --- Metadata ---
NAME = 'Excel Data Mapper'
VERSION = '1.0.0'
DESCRIPTION = 'A tool to map and transfer data between Excel files.'
AUTHOR = 'dohuyhoang93'
ENTRY_POINT = 'app.py'
ICON_FILE = 'icon.ico'
REQUIREMENTS_FILE = 'requirements.txt'

<<<<<<< HEAD
# --- PyInstaller Settings ---

# List of data files and folders to include.
# Format: (source, destination_folder_in_bundle)
DATA_FILES = [
    ('icon.ico', '.'),
    ('configs', 'configs') # Keep configs as it is necessary for the app to run
]

# --- Helper Functions ---

def read_requirements():
    """Reads the requirements.txt file and returns a list of dependencies."""
    if not os.path.exists(REQUIREMENTS_FILE):
        return []
    with open(REQUIREMENTS_FILE, 'r', encoding='utf-8') as f:
        return [line.strip() for line in f if line.strip() and not line.startswith('#')]

class BuildBinaryCommand(build_py):
    """Custom command to build the application into a binary using PyInstaller."""
    
    description = 'Build the application into a binary (requires PyInstaller)'
    user_options = build_py.user_options + [
        ('onefile', None, 'Build as a single executable file.'),
        ('onedir', None, 'Build as a folder containing all dependencies (default).')
    ]

    def initialize_options(self):
        super().initialize_options()
        self.onefile = None
        self.onedir = None

    def finalize_options(self):
        super().finalize_options()
        # Default to onedir if no option is specified
        if self.onefile is None and self.onedir is None:
            self.onedir = True

    def run(self):
        # 1. Ensure pyinstaller is installed
        try:
            import PyInstaller
        except ImportError:
            print("\n[ERROR] PyInstaller is not installed. Please install it first:")
            print(f"  {sys.executable} -m pip install pyinstaller")
            sys.exit(1)

        # 2. Build the PyInstaller command
        separator = ';' if sys.platform.startswith('win') else ':'
        command = [
            'pyinstaller',
            '--noconfirm',
            '--windowed',
            '--paths', '.',
            f'--name={NAME}',
            f'--icon={ICON_FILE}',
            f'--distpath={os.path.abspath("dist")}',
            f'--workpath={os.path.abspath("build")}',
        ]

        # Add hidden imports

        # Add data files
        for src, dest in DATA_FILES:
            command.extend(['--add-data', f'{src}{separator}{dest}'])

        if self.onefile:
            print("\n--- Building a single-file executable ---")
            command.append('--onefile')
        else:
            print("\n--- Building a one-directory bundle ---")
            # No extra flag needed for one-dir, it's the default

        # Add the main script
        command.append(ENTRY_POINT)
=======
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
>>>>>>> 4b8cbb364fe49caa4def35363b0714671af8f6c7

        # 3. Run the command
        print(f"Running command: {' '.join(command)}")
        try:
            subprocess.run(command, check=True)
            print("\n--- Build successful! ---")
            print(f"The output is in the '{os.path.abspath('dist')}' directory.")
        except subprocess.CalledProcessError as e:
            print(f"\n[ERROR] PyInstaller failed with exit code {e.returncode}.")
            sys.exit(1)
        except FileNotFoundError:
            print("\n[ERROR] 'pyinstaller' command not found. Is it installed and in your PATH?")
            sys.exit(1)


# --- Setup Configuration ---
setup(
    name=NAME,
    version=VERSION,
    description=DESCRIPTION,
    author=AUTHOR,
    packages=find_packages(),
    install_requires=read_requirements(),
    cmdclass={
        'build_binary': BuildBinaryCommand,
    },
)
