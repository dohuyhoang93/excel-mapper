# setup.py
import sys
import os
import shutil
import subprocess
from setuptools import setup, find_packages
from setuptools.command.build_py import build_py

# --- Metadata ---
NAME = 'Excel Data Mapper'
VERSION = '1.0.0'
DESCRIPTION = 'A tool to map and transfer data between Excel files.'
AUTHOR = 'dohuyhoang93'
ENTRY_POINT = 'app.py'
ICON_FILE = 'icon.ico'
REQUIREMENTS_FILE = 'requirements.txt'

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
