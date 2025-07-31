import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import ttkbootstrap as ttk_boot
from ttkbootstrap.constants import *
import json
import os
import logging
from pathlib import Path
from typing import Optional
import openpyxl
from openpyxl.styles import *
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell
import subprocess
import sys
from datetime import datetime
import traceback
import shutil
from threading import Thread
import gc
import time
import psutil
from logic.parser import ExcelParser

# Cấu hình logging
logging.basicConfig(
    filename='app.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)

class FileHandleManager:
    """Manages file handles to prevent Excel file locking issues"""
    
    @staticmethod
    def force_release_handles():
        """Force garbage collection and release file handles"""
        # Multiple rounds of garbage collection
        for _ in range(3):
            gc.collect()
            time.sleep(0.05)  # Small delay between collections
    
    @staticmethod
    def is_file_locked(file_path: str) -> bool:
        """Check if a file is currently locked by another process"""
        try:
            # Try to open the file in exclusive mode
            with open(file_path, 'r+b'):
                return False
        except (IOError, OSError):
            return True
    
    @staticmethod
    def wait_for_file_release(file_path: str, max_wait_seconds: int = 5) -> bool:
        """Wait for a file to be released by other processes"""
        start_time = time.time()
        
        while time.time() - start_time < max_wait_seconds:
            if not FileHandleManager.is_file_locked(file_path):
                return True
            
            # Force garbage collection while waiting
            FileHandleManager.force_release_handles()
            time.sleep(0.2)
        
        return False
    
    @staticmethod
    def get_processes_using_file(file_path: str) -> list:
        """Get list of processes that are using the specified file"""
        processes = []
        try:
            for proc in psutil.process_iter(['pid', 'name', 'open_files']):
                try:
                    if proc.info['open_files']:
                        for file_info in proc.info['open_files']:
                            if os.path.samefile(file_info.path, file_path):
                                processes.append({
                                    'pid': proc.info['pid'],
                                    'name': proc.info['name']
                                })
                                break
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess, OSError):
                    continue
        except Exception as e:
            logging.warning(f"Error checking file usage: {e}")
        
        return processes

class ExcelDataMapper:
    def __init__(self):
        self.root = ttk_boot.Window(themename="flatly")
        self.root.title("Excel Data Mapper")
        self.root.geometry("1000x800")
        
        # Icon handling for PyInstaller
        self.icon_path = None
        try:
            if getattr(sys, 'frozen', False):
                # PyInstaller bundle
                base_path = sys._MEIPASS
            else:
                # Development
                base_path = os.path.dirname(os.path.abspath(__file__))
            
            icon_path = os.path.join(base_path, "icon.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
                self.icon_path = icon_path
        except:
            pass  # Ignore icon errors
        
        # Variables
        self.source_file = tk.StringVar()
        self.dest_file = tk.StringVar()
        self.source_header_start_row = tk.IntVar(value=1)
        self.source_header_end_row = tk.IntVar(value=1)
        self.dest_header_start_row = tk.IntVar(value=9)
        self.dest_header_end_row = tk.IntVar(value=9)
        self.sort_column = tk.StringVar()
        self.dest_stop_row = tk.IntVar(value=0)  # 0 means no limit - DEPRECATED by write_end_row
        self.current_theme = "flatly"

        # Write zone settings
        self.dest_write_start_row = tk.IntVar(value=11)
        self.dest_write_end_row = tk.IntVar(value=0) # 0 for no limit
        self.dest_skip_rows = tk.StringVar(value="") # e.g., "15, 20-25"
        self.respect_cell_protection = tk.BooleanVar(value=True)
        self.respect_formulas = tk.BooleanVar(value=True)
        self.detection_keywords = tk.StringVar(value="total,sum,cộng,tổng,thành tiền")
        
        # Data storage
        self.source_columns = {} # {name: index}
        self.dest_columns = {}   # {name: index}
        self.column_mappings = {}
        self.mapping_widgets = []
        self.mapping_combos = {}  # Khởi tạo sớm để tránh lỗi hasattr
        
        # Configuration
        self.config_dir = Path("configs")
        self.config_dir.mkdir(exist_ok=True)
        
        self.setup_menu()
        self.setup_gui()
        
        # Load last configuration if exists
        self.load_last_config()
        
    def setup_menu(self):
        # Tạo một Frame để hoạt động như thanh menu
        menubar_frame = ttk_boot.Frame(self.root)
        menubar_frame.pack(fill=X, side=TOP, padx=5, pady=(2, 0))

        # --- File Menu ---
        file_button = ttk_boot.Button(menubar_frame, text="File", bootstyle="link")
        file_button.pack(side=LEFT, padx=(5, 10))
        
        file_menu = tk.Menu(file_button, tearoff=0)
        file_menu.add_command(label="Open Destination Folder", command=self.open_dest_folder)
        file_menu.add_command(label="Force Release File Handles", command=self.force_release_excel_handles)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)

        file_button.config(command=lambda: file_menu.post(file_button.winfo_rootx(), file_button.winfo_rooty() + file_button.winfo_height()))

        # --- Settings Menu ---
        settings_button = ttk_boot.Button(menubar_frame, text="Settings", bootstyle="link")
        settings_button.pack(side=LEFT, padx=10)

        settings_menu = tk.Menu(settings_button, tearoff=0)
        settings_menu.add_command(label="Switch Theme", command=self.toggle_theme)
        settings_menu.add_command(label="Configure Detection...", command=self.open_detection_config_dialog)
        
        settings_button.config(command=lambda: settings_menu.post(settings_button.winfo_rootx(), settings_button.winfo_rooty() + settings_button.winfo_height()))

        # --- About Menu ---
        about_button = ttk_boot.Button(menubar_frame, text="About", bootstyle="link")
        about_button.pack(side=LEFT, padx=10)
        
        about_menu = tk.Menu(about_button, tearoff=0)
        about_menu.add_command(label="Info", command=self.show_about)

        about_button.config(command=lambda: about_menu.post(about_button.winfo_rootx(), about_button.winfo_rooty() + about_button.winfo_height()))
        
    def setup_gui(self):
        # Main container
        main_frame = ttk_boot.Frame(self.root, padding=10)
        main_frame.pack(fill=BOTH, expand=True, side=TOP)
        
        # File selection section
        file_frame = ttk_boot.LabelFrame(main_frame, text="File Selection", padding=10)
        file_frame.pack(fill=X, pady=(0, 10))
        
        # Source file
        ttk_boot.Label(file_frame, text="Source File:").grid(row=0, column=0, sticky=W, pady=2)
        ttk_boot.Entry(file_frame, textvariable=self.source_file, width=50).grid(row=0, column=1, padx=5, pady=5, sticky=EW)
        ttk_boot.Button(file_frame, bootstyle="outline", text="Browse", command=self.browse_source_file).grid(row=0, column=2, padx=5, pady=5)
        
        # Destination file
        ttk_boot.Label(file_frame, text="Destination File:").grid(row=1, column=0, sticky=W, pady=2)
        ttk_boot.Entry(file_frame, textvariable=self.dest_file, width=50).grid(row=1, column=1, padx=5, pady=2, sticky=EW)
        ttk_boot.Button(file_frame, bootstyle="outline", text="Browse", command=self.browse_dest_file).grid(row=1, column=2, padx=5, pady=2)
        
        file_frame.columnconfigure(1, weight=1)
        
        # Header row configuration
        header_frame = ttk_boot.LabelFrame(main_frame, text="Header Configuration", padding=10)
        header_frame.pack(fill=X, pady=(0, 10))

        ttk_boot.Label(header_frame, text="Source Header Rows:").grid(row=0, column=0, sticky=W, pady=2)
        ttk_boot.Label(header_frame, text="From:").grid(row=0, column=1, sticky=W, padx=(10, 0))
        ttk_boot.Spinbox(header_frame, from_=1, to=50, textvariable=self.source_header_start_row, width=5).grid(row=0, column=2, padx=5)
        ttk_boot.Label(header_frame, text="To:").grid(row=0, column=3, sticky=W)
        ttk_boot.Spinbox(header_frame, from_=1, to=50, textvariable=self.source_header_end_row, width=5).grid(row=0, column=4, padx=5)

        ttk_boot.Label(header_frame, text="Destination Header Rows:").grid(row=0, column=5, sticky=W, padx=(20, 0), pady=2)
        ttk_boot.Label(header_frame, text="From:").grid(row=0, column=6, sticky=W, padx=(10, 0))
        ttk_boot.Spinbox(header_frame, from_=1, to=50, textvariable=self.dest_header_start_row, width=5).grid(row=0, column=7, padx=5)
        ttk_boot.Label(header_frame, text="To:").grid(row=0, column=8, sticky=W)
        ttk_boot.Spinbox(header_frame, from_=1, to=50, textvariable=self.dest_header_end_row, width=5).grid(row=0, column=9, padx=5)
        
        self.load_cols_button = ttk_boot.Button(header_frame, text="Load Columns", command=self.safe_load_columns, bootstyle=INFO)
        self.load_cols_button.grid(row=0, column=10, padx=20)
        
        # Write zone configuration
        write_zone_frame = ttk_boot.LabelFrame(main_frame, text="Setting write zone", padding=10)
        write_zone_frame.pack(fill=X, pady=(0, 10))

        # --- Row 0: Boundaries ---
        ttk_boot.Label(write_zone_frame, text="Start Write Row:").grid(row=0, column=0, sticky=W, padx=5, pady=5)
        ttk_boot.Spinbox(write_zone_frame, from_=1, to=99999, textvariable=self.dest_write_start_row, width=8).grid(row=0, column=1, sticky=W, padx=5)
        
        ttk_boot.Label(write_zone_frame, text="End Write Row:").grid(row=0, column=2, sticky=W, padx=(20, 5), pady=5)
        ttk_boot.Spinbox(write_zone_frame, from_=0, to=99999, textvariable=self.dest_write_end_row, width=8).grid(row=0, column=3, sticky=W, padx=5)
        
        self.detect_button = ttk_boot.Button(write_zone_frame, text="Detect Zone", command=self.detect_write_zone, bootstyle="outline-info")
        self.detect_button.grid(row=0, column=4, padx=(10, 5), pady=5, sticky=E)

        # --- Row 1: Skip Rules ---
        ttk_boot.Label(write_zone_frame, text="Skip Rows (e.g., 15, 20-25):").grid(row=1, column=0, sticky=W, padx=5, pady=5)
        ttk_boot.Entry(write_zone_frame, textvariable=self.dest_skip_rows).grid(row=1, column=1, columnspan=4, padx=5, sticky=EW)

        # --- Row 2: Protection Rules ---
        ttk_boot.Checkbutton(write_zone_frame, text="Respect cell protection", variable=self.respect_cell_protection).grid(row=2, column=0, columnspan=2, sticky=W, padx=5, pady=5)
        ttk_boot.Checkbutton(write_zone_frame, text="Respect formulas", variable=self.respect_formulas).grid(row=2, column=2, columnspan=3, sticky=W, padx=(20, 5), pady=5)

        # Configure column weights for responsive resizing
        write_zone_frame.columnconfigure(1, weight=1)
        write_zone_frame.columnconfigure(3, weight=1)
        write_zone_frame.columnconfigure(4, weight=2)

        # Column mapping section
        self.mapping_frame = ttk_boot.LabelFrame(main_frame, text="Column Mapping", padding=10)
        self.mapping_frame.pack(fill=BOTH, expand=True, pady=(0, 10))
        
        # Create scrollable frame for mappings
        self.setup_mapping_scroll()
        
        # Sort configuration
        sort_frame = ttk_boot.LabelFrame(main_frame, text="Sort Configuration", padding=10)
        sort_frame.pack(fill=X, pady=(0, 10))
        
        ttk_boot.Label(sort_frame, text="Sort by Column (optional):").grid(row=0, column=0, sticky=W, pady=2)
        self.sort_combo = ttk_boot.Combobox(sort_frame, textvariable=self.sort_column, width=50)
        self.sort_combo.grid(row=0, column=1, padx=5, sticky=W)
        
        # Action buttons
        action_frame = ttk_boot.Frame(main_frame)
        action_frame.pack(fill=X, pady=(0, 10))
        
        self.save_button = ttk_boot.Button(action_frame, text="Save Configuration", command=self.save_config, bootstyle=SUCCESS)
        self.save_button.pack(side=LEFT, padx=5)
        
        self.load_button = ttk_boot.Button(action_frame, text="Load Configuration", command=self.load_config, bootstyle=INFO)
        self.load_button.pack(side=LEFT, padx=5)
        
        self.execute_button = ttk_boot.Button(action_frame, text="Execute Transfer", command=self.execute_transfer, bootstyle=PRIMARY)
        self.execute_button.pack(side=RIGHT, padx=5)
        
        self.preview_button = ttk_boot.Button(action_frame, text="Preview Transfer", command=self.preview_transfer, bootstyle="outline-secondary")
        self.preview_button.pack(side=RIGHT, padx=5)
        
        # Status bar
        self.status_frame = ttk_boot.Frame(main_frame)
        self.status_frame.pack(fill=X, pady=(5, 0))
        
        self.progress = ttk_boot.Progressbar(self.status_frame, mode='determinate', bootstyle=SUCCESS)
        self.progress.pack(side=LEFT, fill=X, expand=True, padx=(0, 10))
        
        self.status_label = ttk_boot.Label(self.status_frame, text="Ready")
        self.status_label.pack(side=RIGHT)
        
    def setup_mapping_scroll(self):
        # Create canvas and scrollbar for mapping section
        canvas = tk.Canvas(self.mapping_frame, height=200)
        scrollbar = ttk_boot.Scrollbar(self.mapping_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk_boot.Frame(canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Bind mousewheel to canvas
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind("<MouseWheel>", _on_mousewheel)
    
    def browse_source_file(self):
        filename = filedialog.askopenfilename(
            title="Select Source Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.source_file.set(filename)
            self.log_info(f"Source file selected: {filename}")
    
    def browse_dest_file(self):
        filename = filedialog.askopenfilename(
            title="Select Destination Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.dest_file.set(filename)
            self.log_info(f"Destination file selected: {filename}")
    
    def force_release_excel_handles(self):
        """Force release of all Excel file handles"""
        try:
            # Clear references to data
            if hasattr(self, 'source_columns'):
                self.source_columns = {}
            if hasattr(self, 'dest_columns'):
                self.dest_columns = {}
            
            # Force garbage collection
            FileHandleManager.force_release_handles()
            
            self.update_status("File handles released")
            self.log_info("Forced release of Excel file handles")
            messagebox.showinfo("Info", "Excel file handles have been released.")
            
        except Exception as e:
            self.log_error(f"Error forcing handle release: {str(e)}")
            messagebox.showerror("Error", f"Error releasing handles: {str(e)}")
    
    def check_file_accessibility(self, file_path: str) -> bool:
        """Check if file is accessible for reading/writing"""
        try:
            if not os.path.exists(file_path):
                return False
            
            # Wait for file to be released if locked
            if FileHandleManager.is_file_locked(file_path):
                self.update_status(f"Waiting for file to be released: {os.path.basename(file_path)}")
                
                if not FileHandleManager.wait_for_file_release(file_path, max_wait_seconds=10):
                    # Show which processes are using the file
                    processes = FileHandleManager.get_processes_using_file(file_path)
                    if processes:
                        process_names = [p['name'] for p in processes]
                        self.log_error(f"File locked by processes: {', '.join(process_names)}")
                        messagebox.showwarning(
                            "File Locked", 
                            f"File is locked by: {', '.join(process_names)}\n"
                            f"Please close these applications and try again."
                        )
                    return False
            
            return True
            
        except Exception as e:
            self.log_error(f"Error checking file accessibility: {str(e)}")
            return False
    
    def get_excel_columns(self, file_path, start_row, end_row):
        """Extracts headers using the centralized ExcelParser with guaranteed cleanup."""
        parser = None
        try:
            # Force garbage collection before opening
            FileHandleManager.force_release_handles()
            
            parser = ExcelParser(file_path)
            with parser as p:
                headers = p.get_headers(start_row, end_row)
                
                # Filter out headers that are None, empty, or whitespace-only
                filtered_headers = {
                    name: index for name, index in headers.items()
                    if name and str(name).strip()
                }
                
                # Make a copy to ensure no references to the parser remain
                result = dict(filtered_headers)
                
            # Explicit cleanup
            if parser:
                parser._cleanup()
                
            # Force garbage collection after closing
            FileHandleManager.force_release_handles()
            
            return result
            
        except Exception as e:
            self.log_error(f"Error reading Excel columns with parser: {str(e)}")
            raise
        finally:
            # Ensure cleanup even if exception occurs
            if parser:
                try:
                    parser._cleanup()
                except:
                    pass
            # Final cleanup
            FileHandleManager.force_release_handles()

    def safe_load_columns(self, saved_sort_col: Optional[str] = None, apply_suggestions: bool = True):
        """Load columns with enhanced file handle management"""
        try:
            if not self.source_file.get() or not self.dest_file.get():
                messagebox.showwarning("Warning", "Please select both source and destination files first.")
                return
            
            # Check file accessibility first
            if not self.check_file_accessibility(self.source_file.get()):
                messagebox.showerror("Error", f"Cannot access source file: {self.source_file.get()}")
                return
                
            if not self.check_file_accessibility(self.dest_file.get()):
                messagebox.showerror("Error", f"Cannot access destination file: {self.dest_file.get()}")
                return
            
            # Force release any existing handles
            self.force_release_excel_handles()
            
            self.update_status("Loading columns...")
            
            # Load source columns
            try:
                self.source_columns = self.get_excel_columns(
                    self.source_file.get(), 
                    self.source_header_start_row.get(), 
                    self.source_header_end_row.get()
                )
                # Force release after loading
                FileHandleManager.force_release_handles()
                
            except Exception as e:
                self.log_error(f"Error loading source columns: {str(e)}")
                messagebox.showerror("Error", f"Failed to load source columns: {str(e)}")
                return
            
            # Small delay between file operations
            time.sleep(0.2)
            
            # Load destination columns
            try:
                self.dest_columns = self.get_excel_columns(
                    self.dest_file.get(), 
                    self.dest_header_start_row.get(), 
                    self.dest_header_end_row.get()
                )
                # Force release after loading
                FileHandleManager.force_release_handles()
                
            except Exception as e:
                self.log_error(f"Error loading destination columns: {str(e)}")
                messagebox.showerror("Error", f"Failed to load destination columns: {str(e)}")
                return
            
            if not self.source_columns or not self.dest_columns:
                messagebox.showerror("Error", "Could not load columns. Please check file paths and header row numbers.")
                return
            
            # Update sort combo
            source_keys = list(self.source_columns.keys())
            self.sort_combo['values'] = source_keys
            self.root.update_idletasks()

            # Set the saved sort column value
            if saved_sort_col and saved_sort_col in source_keys:
                self.sort_column.set(saved_sort_col)
            
            # Create mapping widgets
            self.create_mapping_widgets(apply_suggestions=apply_suggestions)
            
            self.update_status(f"Loaded {len(self.source_columns)} unique source columns and {len(self.dest_columns)} unique destination columns")
            self.log_info(f"Columns loaded successfully")
            
        except Exception as e:
            self.log_error(f"Error loading columns: {str(e)}")
            messagebox.showerror("Error", f"Failed to load columns: {str(e)}")
            self.update_status("Error loading columns")
        finally:
            # Always force release handles at the end
            FileHandleManager.force_release_handles()
    
    # Keep original load_columns for backward compatibility but redirect to safe version
    def load_columns(self, saved_sort_col: Optional[str] = None, apply_suggestions: bool = True):
        """Backward compatibility wrapper for safe_load_columns"""
        return self.safe_load_columns(saved_sort_col, apply_suggestions)
    
    def create_mapping_widgets(self, apply_suggestions: bool = True):
        """Create dropdown widgets for column mapping using a grid layout for perfect alignment."""
        # Clear existing widgets
        for widget in self.mapping_widgets:
            widget.destroy()
        self.mapping_widgets.clear()
        
        # Configure grid columns to expand
        self.scrollable_frame.columnconfigure(0, weight=1)
        self.scrollable_frame.columnconfigure(2, weight=1)

        # Header row
        ttk_boot.Label(self.scrollable_frame, text="Source Column", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky=W, padx=5, pady=(5, 10))
        ttk_boot.Label(self.scrollable_frame, text="Destination Column", font=("Arial", 10, "bold")).grid(row=0, column=2, sticky=W, padx=5, pady=(5, 10))
        
        # Create mapping rows
        self.mapping_combos = {}
        
        for i, source_col_name in enumerate(self.source_columns.keys(), start=1):
            # Source column label
            source_label = ttk_boot.Label(self.scrollable_frame, text=source_col_name, anchor=W, width=50)
            source_label.grid(row=i, column=0, sticky=EW, padx=5, pady=2)
            
            # Arrow
            arrow_label = ttk_boot.Label(self.scrollable_frame, text="→")
            arrow_label.grid(row=i, column=1, sticky=W, padx=5)
            
            # Destination column combobox
            dest_combo = ttk_boot.Combobox(self.scrollable_frame, values=[""] + list(self.dest_columns.keys()), width=50)
            dest_combo.grid(row=i, column=2, sticky=EW, padx=5, pady=2)
            
            # Auto-suggest mapping only if enabled
            if apply_suggestions:
                suggested = self.suggest_mapping(source_col_name, list(self.dest_columns.keys()))
                if suggested:
                    dest_combo.set(suggested)
            
            self.mapping_combos[source_col_name] = dest_combo
            self.mapping_widgets.extend([source_label, arrow_label, dest_combo])
    
    def suggest_mapping(self, source_col, dest_cols):
        """Suggest best matching destination column using a scoring algorithm."""
        # Do not suggest for auto-generated column names, as they have no semantic meaning
        if str(source_col).startswith('Column_'):
            return ""

        import re

        def normalize_and_tokenize(text: str) -> set:
            # Replace full-width spaces and other separators with standard space
            text = re.sub(r'[\s　_\-]+', ' ', text) 
            # Remove punctuation
            text = re.sub(r'[()[\]{}]', '', text)
            # Convert to lower and split into words
            return set(text.lower().strip().split())

        source_tokens = normalize_and_tokenize(source_col)
        
        best_match = ""
        max_score = 0

        # Keywords mapping based on destination structure
        keywords_map = {
            'content': 'contents',
            'purpose': 'purpose', 
            'amount': 'amount',
            'vat': 'vat',
            'currency': 'currency',
            'date': 'trading date',
            'no': 'no.',
            'number': 'no.',
            'code': 'code reference',
            'total': 'sub total'
        }

        for dest_col in dest_cols:
            current_score = 0
            dest_tokens = normalize_and_tokenize(dest_col)

            if not dest_tokens:
                continue

            # Rule 1: Exact match after normalization (highest score)
            if source_tokens == dest_tokens:
                current_score = 100
            
            # Rule 2: Common tokens (high score)
            common_tokens = source_tokens.intersection(dest_tokens)
            current_score += len(common_tokens) * 50

            # Rule 3: Keyword mapping (good score)
            for key, value in keywords_map.items():
                if key in source_tokens and value in dest_tokens:
                    current_score += 40
            
            # Rule 4: Substring match (lower score)
            # Use normalized strings without spaces for better substring matching
            source_norm_str = "".join(source_tokens)
            dest_norm_str = "".join(dest_tokens)
            if source_norm_str in dest_norm_str or dest_norm_str in source_norm_str:
                current_score += 20

            # Update best match if current score is higher
            if current_score > max_score:
                max_score = current_score
                best_match = dest_col
        
        return best_match
    
    def save_config(self):
        """Save current mapping configuration to a user-selected JSON file."""
        try:
            if not hasattr(self, 'mapping_combos') or not self.mapping_combos:
                messagebox.showwarning("Warning", "No mappings to save. Please load columns first.")
                return

            # Open file dialog to choose save location
            config_file_path = filedialog.asksaveasfilename(
                title="Save Configuration As",
                defaultextension=".json",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
            )

            if not config_file_path:
                return # User cancelled

            # Get current mappings
            mappings = {source_col: combo.get() for source_col, combo in self.mapping_combos.items() if combo.get()}

            config = {
                "source_file": self.source_file.get(),
                "dest_file": self.dest_file.get(),
                "source_header_start_row": self.source_header_start_row.get(),
                "source_header_end_row": self.source_header_end_row.get(),
                "dest_header_start_row": self.dest_header_start_row.get(),
                "dest_header_end_row": self.dest_header_end_row.get(),
                "dest_write_start_row": self.dest_write_start_row.get(),
                "dest_write_end_row": self.dest_write_end_row.get(),
                "dest_skip_rows": self.dest_skip_rows.get(),
                "respect_cell_protection": self.respect_cell_protection.get(),
                "respect_formulas": self.respect_formulas.get(),
                "detection_keywords": self.detection_keywords.get(),
                "sort_column": self.sort_column.get(),
                "theme": self.current_theme,
                "mapping": mappings,
                "created_date": datetime.now().isoformat()
            }

            with open(config_file_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)

            self.update_status(f"Configuration saved to {os.path.basename(config_file_path)}")
            self.log_info(f"Configuration saved: {config_file_path}")

        except Exception as e:
            self.log_error(f"Error saving configuration: {str(e)}")
            messagebox.showerror("Error", f"Failed to save configuration: {str(e)}")
    
    def load_config(self):
        """Load mapping configuration from a user-selected JSON file."""
        try:
            config_file_path = filedialog.askopenfilename(
                title="Load Configuration",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
            )

            if not config_file_path:
                return # User cancelled

            with open(config_file_path, 'r', encoding='utf-8') as f:
                config = json.load(f)

            # --- Staged Loading Process --- #

            # Stage 1: Set file paths and header rows from config
            self.source_file.set(config.get("source_file", ""))
            self.dest_file.set(config.get("dest_file", ""))
            self.source_header_start_row.set(config.get("source_header_start_row", 1))
            self.source_header_end_row.set(config.get("source_header_end_row", 1))
            self.dest_header_start_row.set(config.get("dest_header_start_row", 9))
            self.dest_header_end_row.set(config.get("dest_header_end_row", 9))
            
            # Load write zone settings
            self.dest_write_start_row.set(config.get("dest_write_start_row", self.dest_header_end_row.get() + 1))
            self.dest_write_end_row.set(config.get("dest_write_end_row", 0))
            self.dest_skip_rows.set(config.get("dest_skip_rows", ""))
            self.respect_cell_protection.set(config.get("respect_cell_protection", True))
            self.respect_formulas.set(config.get("respect_formulas", True))
            self.detection_keywords.set(config.get("detection_keywords", "total,sum,cộng,tổng,thành tiền"))

            # Stage 2: Load and apply the theme
            new_theme = config.get("theme", "flatly")
            if new_theme != self.current_theme:
                self.root.style.theme_use(new_theme)
                self.current_theme = new_theme

                            # Stage 3: Load columns and pass the saved sort column to be set internally
            if self.source_file.get() and self.dest_file.get():
                saved_sort_col = config.get("sort_column", "")
                # Load columns WITHOUT applying suggestions, as they will be loaded from config
                self.safe_load_columns(saved_sort_col=saved_sort_col, apply_suggestions=False)

                # Stage 4: Apply the detailed column mappings
                mappings = config.get("mapping", {})
                for source_col, dest_col in mappings.items():
                    if source_col in self.mapping_combos:
                        self.mapping_combos[source_col].set(dest_col)

            self.update_status(f"Configuration loaded from {os.path.basename(config_file_path)}")
            self.log_info(f"Configuration loaded: {config_file_path}")

        except Exception as e:
            self.log_error(f"Error in load_config: {str(e)}")
            messagebox.showerror("Error", f"Failed to load configuration: {str(e)}")
    
    def load_last_config(self):
        """Load the most recent configuration automatically"""
        try:
            config_files = list(self.config_dir.glob("*.json"))
            if config_files:
                # Get most recent config file
                latest_config = max(config_files, key=os.path.getctime)
                
                with open(latest_config, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                # Only load file paths and header rows, not full mapping
                if os.path.exists(config.get("source_file", "")):
                    self.source_file.set(config.get("source_file", ""))
                if os.path.exists(config.get("dest_file", "")):
                    self.dest_file.set(config.get("dest_file", ""))
                    
                self.source_header_start_row.set(config.get("source_header_start_row", 1))
                self.source_header_end_row.set(config.get("source_header_end_row", 1))
                self.dest_header_start_row.set(config.get("dest_header_start_row", 9))
                self.dest_header_end_row.set(config.get("dest_header_end_row", 9))

                # Load write zone settings from last config
                self.dest_write_start_row.set(config.get("dest_write_start_row", self.dest_header_end_row.get() + 1))
                self.dest_write_end_row.set(config.get("dest_write_end_row", 0))
                self.dest_skip_rows.set(config.get("dest_skip_rows", ""))

                # Load and apply theme from last config
                new_theme = config.get("theme", "flatly")
                if new_theme != self.current_theme:
                    self.root.style.theme_use(new_theme)
                    self.current_theme = new_theme
                
        except Exception as e:
            self.log_error(f"Error loading last config: {str(e)}")
    
    def execute_transfer(self):
        """Starts the data transfer operation in a separate thread to keep the GUI responsive."""
        # --- Validation ---
        if not self.source_file.get() or not self.dest_file.get():
            messagebox.showwarning("Warning", "Please select both source and destination files.")
            return
        
        if not hasattr(self, 'mapping_combos') or not self.mapping_combos:
            messagebox.showwarning("Warning", "Please load columns first.")
            return
        
        mappings = {source_col: combo.get() for source_col, combo in self.mapping_combos.items() if combo.get()}
        
        if not mappings:
            messagebox.showwarning("Warning", "Please configure at least one column mapping.")
            return
        
        duplicate_destinations = []
        dest_values = list(mappings.values())
        for dest_col in set(dest_values):
            if dest_values.count(dest_col) > 1:
                duplicate_destinations.append(dest_col)
        
        if duplicate_destinations:
            messagebox.showerror("Error", f"Duplicate destination columns detected: {', '.join(duplicate_destinations)}")
            return
        
        if len(self.source_file.get()) > 255 or len(self.dest_file.get()) > 255:
            messagebox.showwarning("Warning", "File paths are very long (>255 characters). This might cause issues.")

        # Check file accessibility before starting transfer
        if not self.check_file_accessibility(self.source_file.get()):
            messagebox.showerror("Error", f"Cannot access source file: {self.source_file.get()}")
            return
            
        if not self.check_file_accessibility(self.dest_file.get()):
            messagebox.showerror("Error", f"Cannot access destination file: {self.dest_file.get()}")
            return

        # --- Disable Widgets & Start Thread ---
        self.disable_controls()
        self.update_status("Starting data transfer...")
        self.progress['value'] = 0
        self.root.update()

        # Run the actual transfer in a separate thread
        transfer_thread = Thread(target=self._execute_transfer_thread, args=(mappings,))
        transfer_thread.daemon = True
        transfer_thread.start()

    def _execute_transfer_thread(self, mappings):
        """The actual data transfer logic that runs in the background."""
        try:
            self.perform_data_transfer(mappings)
            
            self.root.after(0, self.on_transfer_success)
            
        except Exception as e:
            self.root.after(0, self.on_transfer_error, e)
    
    def on_transfer_success(self):
        """Handles successful completion of the transfer in the main thread."""
        self.progress['value'] = 100
        self.update_status("Transfer completed successfully")
        self.enable_controls()
        messagebox.showinfo("Success", "Data transfer completed successfully!")
        if messagebox.askyesno("Open Folder", "Would you like to open the destination folder?"):
            self.open_dest_folder()

    def on_transfer_error(self, error):
        """Handles errors from the transfer in the main thread."""
        self.log_error(f"Error in execute_transfer: {str(error)}")
        self.log_error(traceback.format_exc())
        self.update_status("Transfer failed")
        self.progress['value'] = 0
        self.enable_controls()
        messagebox.showerror("Error", f"Transfer failed: {str(error)}")

    def disable_controls(self):
        """Disables key controls during processing."""
        self.execute_button.config(state=DISABLED)
        self.load_button.config(state=DISABLED)
        self.save_button.config(state=DISABLED)
        self.load_cols_button.config(state=DISABLED)

    def enable_controls(self):
        """Enables key controls after processing."""
        self.execute_button.config(state=NORMAL)
        self.load_button.config(state=NORMAL)
        self.save_button.config(state=NORMAL)
        self.load_cols_button.config(state=NORMAL)
    
    def perform_data_transfer(self, mappings):
        """Perform the actual data transfer with enhanced file handle management"""
        # Create backup of destination file
        dest_path = Path(self.dest_file.get())
        backup_path = dest_path.with_suffix('.backup' + dest_path.suffix)
        
        try:
            # Make backup
            shutil.copy2(dest_path, backup_path)
            
            # Read source data
            self.update_status("Reading source data...")
            self.progress['value'] = 10
            self.root.update()
            
            source_data = self.read_source_data()
            
            if not source_data:
                raise ValueError("No data found in source file")
            
            # Sort data if specified
            if self.sort_column.get() and self.sort_column.get() in mappings:
                self.update_status("Sorting data...")
                self.progress['value'] = 30
                self.root.update()
                
                try:
                    # Tìm column index để sort
                    sort_col_key = self.sort_column.get()
                    # Sắp xếp thông minh: các dòng có giá trị rỗng/None ở cột sort sẽ bị đẩy xuống cuối
                    source_data = sorted(
                        source_data, 
                        key=lambda x: (x.get(sort_col_key) is None or str(x.get(sort_col_key, "")).strip() == "", str(x.get(sort_col_key, "")))
                    )
                except Exception as e:
                    self.log_error(f"Error sorting data: {str(e)}")
                    # Continue without sorting
            
            # Write to destination
            self.update_status("Writing to destination...")
            self.progress['value'] = 50
            self.root.update()
            
            self.write_to_destination(source_data, mappings)
            
            # Clean up backup
            if backup_path.exists():
                backup_path.unlink()
            
            self.progress['value'] = 100
            self.log_info("Data transfer completed successfully")
            
        except Exception as e:
            # Restore backup if transfer failed
            if backup_path.exists():
                try:
                    shutil.copy2(backup_path, dest_path)
                    backup_path.unlink()
                except:
                    pass  # If restore fails, at least we tried
            raise e
    
    def read_source_data(self):
        """Reads data from the source file with proper resource management."""
        workbook = None
        try:
            # Force garbage collection before opening
            FileHandleManager.force_release_handles()
            
            workbook = openpyxl.load_workbook(self.source_file.get(), data_only=True)
            worksheet = workbook.active
            
            # Data starts after the specified header end row
            start_data_row = self.source_header_end_row.get() + 1
            
            data = []
            for row_index in range(start_data_row, worksheet.max_row + 1):
                row_data = {}
                has_data = False
                
                # Iterate through the logical headers and their starting column indices
                for header_name, col_index in self.source_columns.items():
                    cell = worksheet.cell(row=row_index, column=col_index)
                    value = cell.value
                    
                    if value is not None:
                        has_data = True
                    
                    row_data[header_name] = value
                
                if has_data:
                    data.append(row_data)
            
            # Make a copy to ensure no references to workbook remain
            result = list(data)
            
            return result
            
        except Exception as e:
            self.log_error(f"Error reading source data: {str(e)}")
            raise
        finally:
            # Guaranteed cleanup
            if workbook:
                try:
                    workbook.close()
                except Exception as e:
                    self.log_error(f"Error closing source workbook: {str(e)}")
            
            # Force garbage collection
            FileHandleManager.force_release_handles()
    
    def _parse_skip_rows(self, skip_rows_str: str) -> set:
        """Parses a string like '15, 20-25, 30' into a set of integers."""
        skipped_rows = set()
        if not skip_rows_str:
            return skipped_rows
        
        parts = skip_rows_str.split(',')
        for part in parts:
            part = part.strip()
            if not part:
                continue
            if '-' in part:
                try:
                    start, end = map(int, part.split('-'))
                    if start <= end:
                        skipped_rows.update(range(start, end + 1))
                except ValueError:
                    self.log_warning(f"Could not parse range in skip_rows: {part}")
            else:
                try:
                    skipped_rows.add(int(part))
                except ValueError:
                    self.log_warning(f"Could not parse number in skip_rows: {part}")
        return skipped_rows

    def write_to_destination(self, source_data, mappings):
        """Write data to destination Excel file with protection-aware advanced write zone logic."""
        workbook = None
        try:
            FileHandleManager.force_release_handles()
            workbook = openpyxl.load_workbook(self.dest_file.get())
            worksheet = workbook.active

            dest_headers_map = self.dest_columns
            
            # Get write zone and protection settings
            start_write_row = self.dest_write_start_row.get()
            end_write_row = self.dest_write_end_row.get()
            skip_rows_str = self.dest_skip_rows.get()
            skipped_rows = self._parse_skip_rows(skip_rows_str)
            respect_protection = self.respect_cell_protection.get()
            respect_formulas = self.respect_formulas.get()
            sheet_is_protected = worksheet.protection.sheet

            if start_write_row <= self.dest_header_end_row.get():
                raise ValueError("Start Write Row must be after the destination header rows.")

            def get_writable_cell(row_idx, col_idx):
                cell = worksheet.cell(row=row_idx, column=col_idx)
                if not isinstance(cell, MergedCell):
                    return cell
                for merged_range in worksheet.merged_cells.ranges:
                    if (merged_range.min_row <= row_idx <= merged_range.max_row and
                        merged_range.min_col <= col_idx <= merged_range.max_col):
                        return worksheet.cell(row=merged_range.min_row, column=merged_range.min_col)
                return cell

            # --- 1. Clear existing data, respecting protection and skip rules ---
            clear_until_row = end_write_row if end_write_row > 0 else worksheet.max_row + 50
            
            cleared_anchors = set()
            for row_to_clear in range(start_write_row, clear_until_row + 1):
                if row_to_clear in skipped_rows:
                    continue

                # Check for protection before clearing
                if respect_protection and sheet_is_protected:
                    is_row_locked = False
                    for dest_col_num in dest_headers_map.values():
                        cell = get_writable_cell(row_to_clear, dest_col_num)
                        if cell.protection and cell.protection.locked:
                            is_row_locked = True
                            break
                    if is_row_locked:
                        continue

                for dest_col_num in dest_headers_map.values():
                    anchor_cell = get_writable_cell(row_to_clear, dest_col_num)
                    if (anchor_cell.row >= start_write_row and 
                        anchor_cell.coordinate not in cleared_anchors and
                        anchor_cell.row not in skipped_rows):
                        
                        if respect_formulas and anchor_cell.data_type == 'f':
                            continue
                        
                        anchor_cell.value = None
                        cleared_anchors.add(anchor_cell.coordinate)

            # --- 2. Write new data with protection-aware skipping logic ---
            current_write_row = start_write_row
            stop_writing = False

            for i, row_data in enumerate(source_data):
                # Find the next valid (not skipped, not protected) row to write to
                while True:
                    # Check if we've gone past the end limit
                    if end_write_row > 0 and current_write_row > end_write_row:
                        self.log_warning(f"Reached end of write zone (row {end_write_row}). Stopping data transfer. {len(source_data) - i} source rows were not written.")
                        stop_writing = True
                        break
                    
                    # Condition 1: Is the row explicitly skipped by the user?
                    is_invalid = current_write_row in skipped_rows
                    
                    # Condition 2: If not skipped, is it protected? (only check if enabled)
                    if not is_invalid and respect_protection and sheet_is_protected:
                        for dest_col_num in dest_headers_map.values():
                            cell = get_writable_cell(current_write_row, dest_col_num)
                            if cell.protection and cell.protection.locked:
                                is_invalid = True
                                break # Found a locked cell, row is invalid
                    
                    if not is_invalid:
                        break # Found a valid row, exit the 'while' loop
                    
                    current_write_row += 1 # Row is invalid, move to the next one

                if stop_writing:
                    break # Exit the main 'for' loop

                # Update progress bar
                if i > 0 and i % 10 == 0:
                    progress = 50 + (i / len(source_data)) * 45
                    self.progress['value'] = progress
                    self.root.update()

                # Write data to the found valid row
                for source_col_name, dest_col_name in mappings.items():
                    if dest_col_name in dest_headers_map:
                        dest_col_num = dest_headers_map[dest_col_name]
                        source_value = row_data.get(source_col_name, "")

                        cell_to_write = get_writable_cell(current_write_row, dest_col_num)
                        
                        if cell_to_write.row < current_write_row:
                            continue
                        
                        if respect_formulas and cell_to_write.data_type == 'f':
                            continue

                        cell_to_write.value = source_value
                
                current_write_row += 1 # Move pointer for the next source record

            if stop_writing:
                messagebox.showwarning("Write Limit Reached", f"Data transfer stopped at row {end_write_row} as configured. Not all source data may have been transferred.")

            workbook.save(self.dest_file.get())

        except Exception as e:
            self.log_error(f"Error writing to destination: {str(e)}")
            self.log_error(traceback.format_exc())
            raise
        finally:
            if workbook:
                try:
                    workbook.close()
                except Exception as e:
                    self.log_error(f"Error closing destination workbook: {str(e)}")
            FileHandleManager.force_release_handles()

    def toggle_theme(self):
        """Toggle between light and dark themes"""
        try:
            if self.current_theme == "flatly":
                new_theme = "superhero"
            else:
                new_theme = "flatly"
            
            self.root.style.theme_use(new_theme)
            self.current_theme = new_theme
            
            self.update_status(f"Theme changed to {new_theme}")
            self.log_info(f"Theme changed to {new_theme}")
            
        except Exception as e:
            self.log_error(f"Error changing theme: {str(e)}")
            messagebox.showerror("Error", f"Failed to change theme: {str(e)}")
    
    def open_dest_folder(self):
        """Open the folder containing the destination file"""
        try:
            if not self.dest_file.get():
                messagebox.showwarning("Warning", "No destination file selected.")
                return
            
            dest_path = Path(self.dest_file.get())
            if dest_path.exists():
                folder_path = dest_path.parent
                
                # Open folder based on OS
                if os.name == 'nt':  # Windows
                    os.startfile(folder_path)
                elif os.name == 'posix':  # macOS and Linux
                    subprocess.run(['open' if sys.platform == 'darwin' else 'xdg-open', folder_path])
                
                self.log_info(f"Opened destination folder: {folder_path}")
            else:
                messagebox.showwarning("Warning", "Destination file does not exist.")
                
        except Exception as e:
            self.log_error(f"Error opening destination folder: {str(e)}")
            messagebox.showerror("Error", f"Failed to open folder: {str(e)}")
    
    def show_about(self):
        """Show a custom about dialog with the application icon."""
        about_dialog = tk.Toplevel(self.root)
        about_dialog.title("About Excel Data Mapper")
        about_dialog.geometry("400x350")
        about_dialog.resizable(False, False)
        
        # Set the icon if available
        if self.icon_path:
            try:
                about_dialog.iconbitmap(self.icon_path)
            except:
                pass # Ignore icon errors on the dialog

        # Center the dialog over the main window
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 200
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 150
        about_dialog.geometry(f"+{x}+{y}")

        about_text = """Excel Data Mapper v1.1

A powerful tool for mapping and transferring data 
between Excel files while preserving formatting.

Features:
• Flexible column mapping
• Sort data before transfer
• Preserve Excel formatting and styles
• Configuration save/load
• Theme switching
• Comprehensive error handling
• Enhanced file handle management

Developed by Do Huy Hoang
https://github.com/dohuyhoang93
"""
        
        # Create a frame for the button to ensure it's not overridden
        button_frame = ttk_boot.Frame(about_dialog)
        button_frame.pack(side=BOTTOM, fill=X, pady=10)

        # Use a Label for the text content, packed to fill remaining space
        label = ttk_boot.Label(about_dialog, text=about_text, justify=LEFT, padding=(10, 10))
        label.pack(side=TOP, expand=True, fill=BOTH)

        # OK button to close the dialog, packed inside its frame
        ok_button = ttk_boot.Button(button_frame, text="OK", command=about_dialog.destroy, bootstyle=PRIMARY)
        ok_button.pack()

        # Make the dialog modal
        about_dialog.transient(self.root)
        about_dialog.grab_set()
        self.root.wait_window(about_dialog)
    
    def update_status(self, message):
        """Update status bar message"""
        self.status_label.config(text=message)
        self.root.update()
    
    def log_info(self, message):
        """Log info message"""
        logging.info(message)
    
    def log_warning(self, message):
        """Log warning message"""
        logging.warning(message)

    def log_error(self, message):
        """Log error message"""
        logging.error(message)

    def open_detection_config_dialog(self):
        """Opens a dialog to configure detection keywords."""
        DetectionConfigDialog(self)

    def detect_write_zone(self):
        """Detects start/end rows, correctly handling merged cells and scanning all columns."""
        if not self.dest_file.get() or not os.path.exists(self.dest_file.get()):
            messagebox.showwarning("Warning", "Please select a valid destination file first.")
            return

        try:
            self.update_status("Detecting write zone...")
            
            predicted_start_row = self.dest_header_end_row.get() + 1
            self.dest_write_start_row.set(predicted_start_row)
            
            predicted_end_row = 0
            reason = "No limit found (scanned to end of file)."
            parser = None
            try:
                parser = ExcelParser(self.dest_file.get())
                with parser as p:
                    ws = p.worksheet
                    merged_cell_ranges = list(ws.merged_cells.ranges)

                    def get_value_from_merged_cell(row, col):
                        """Gets a cell's value, resolving merges."""
                        cell = ws.cell(row=row, column=col)
                        if not isinstance(cell, MergedCell):
                            return cell.value
                        for cell_range in merged_cell_ranges:
                            if cell.coordinate in cell_range:
                                return ws.cell(row=cell_range.min_row, column=cell_range.min_col).value
                        return None

                    keywords_str = self.detection_keywords.get().lower()
                    total_keywords = [k.strip() for k in keywords_str.split(',') if k.strip()]
                    
                    # Priority 1: Find a "total" row, scanning ALL columns
                    for row in range(predicted_start_row, ws.max_row + 1):
                        for col in range(1, ws.max_column + 1): # Scan all columns
                            cell_val = get_value_from_merged_cell(row, col)
                            if isinstance(cell_val, str) and total_keywords:
                                if any(keyword in cell_val.lower() for keyword in total_keywords):
                                    predicted_end_row = row - 1
                                    reason = f"Detected keyword '{cell_val.strip()}' on row {row}."
                                    break
                        if predicted_end_row > 0:
                            break
                    
                    # Priority 2: Find the first blank row
                    if predicted_end_row == 0:
                        for row in range(predicted_start_row, ws.max_row + 1):
                            is_row_blank = True
                            for col in range(1, ws.max_column + 1):
                                if get_value_from_merged_cell(row, col) is not None:
                                    is_row_blank = False
                                    break
                            if is_row_blank:
                                predicted_end_row = row - 1
                                reason = f"Detected first blank row at {row}."
                                break
                
                if predicted_end_row >= predicted_start_row:
                    self.dest_write_end_row.set(predicted_end_row)
                    self.update_status(f"Detection complete. Start: {predicted_start_row}, End: {predicted_end_row}. Reason: {reason}")
                else:
                    self.dest_write_end_row.set(0)
                    self.update_status(f"Detection complete. Start: {predicted_start_row}, End: Unlimited.")

            finally:
                if parser:
                    parser._cleanup()
                FileHandleManager.force_release_handles()

        except Exception as e:
            self.log_error(f"Error detecting write zone: {str(e)}")
            messagebox.showerror("Error", f"Failed to detect write zone: {str(e)}")
            self.update_status("Detection failed")

    def preview_transfer(self):
        """Runs a detailed simulation and shows a comprehensive preview report."""
        # --- Validation ---
        if not self.source_file.get() or not os.path.exists(self.source_file.get()):
            messagebox.showerror("Error", "Please select a valid source file.")
            return
        if not self.dest_file.get() or not os.path.exists(self.dest_file.get()):
            messagebox.showerror("Error", "Please select a valid destination file.")
            return
        if not hasattr(self, 'mapping_combos') or not self.mapping_combos:
            messagebox.showwarning("Warning", "Please load columns first.")
            return
        
        self.update_status("Generating simulation report...")
        
        try:
            # --- Data Collection for Simulation ---
            report_data = {}
            
            # 1. Get source data count
            source_data = self.read_source_data()
            report_data['source_row_count'] = len(source_data)
            
            # 2. Analyze destination write zone
            parser = None
            try:
                parser = ExcelParser(self.dest_file.get())
                with parser as p:
                    ws = p.worksheet
                    start_row = self.dest_write_start_row.get()
                    end_row = self.dest_write_end_row.get()
                    
                    if start_row <= self.dest_header_end_row.get():
                        PreviewDialog(self, {"error": "Start Write Row must be after the destination header."})
                        return

                    end_limit = end_row if end_row > 0 else ws.max_row
                    if end_row > 0 and start_row > end_row:
                        PreviewDialog(self, {"error": "Start Write Row cannot be after End Write Row."})
                        return

                    report_data['start_row'] = start_row
                    report_data['end_row'] = end_row or "Unlimited"
                    report_data['total_zone_rows'] = (end_limit - start_row + 1) if end_row > 0 else "Unlimited"

                    skipped_rows_set = self._parse_skip_rows(self.dest_skip_rows.get())
                    respect_protection = self.respect_cell_protection.get()
                    respect_formulas = self.respect_formulas.get()
                    sheet_is_protected = ws.protection.sheet
                    
                    user_skipped_count = 0
                    protected_skipped_count = 0
                    
                    # Calculate skippable rows within the defined zone
                    for r in range(start_row, end_limit + 1):
                        is_user_skipped = r in skipped_rows_set
                        is_auto_skipped = False

                        if is_user_skipped:
                            user_skipped_count += 1
                            continue # No need to check further if user already skipped it

                        if respect_protection and sheet_is_protected:
                            for c in self.dest_columns.values():
                                if ws.cell(r, c).protection and ws.cell(r, c).protection.locked:
                                    is_auto_skipped = True
                                    break
                        
                        if not is_auto_skipped and respect_formulas:
                             for c in self.dest_columns.values():
                                if ws.cell(r, c).data_type == 'f':
                                    is_auto_skipped = True
                                    break
                        
                        if is_auto_skipped:
                            protected_skipped_count += 1
                    
                    report_data['user_skipped_count'] = user_skipped_count
                    report_data['protected_skipped_count'] = protected_skipped_count
                    
                    if end_row > 0:
                        available_slots = report_data['total_zone_rows'] - user_skipped_count - protected_skipped_count
                        report_data['available_slots'] = max(0, available_slots)
                    else:
                        report_data['available_slots'] = "Unlimited"

            finally:
                if parser:
                    parser._cleanup()
                FileHandleManager.force_release_handles()

            # 3. Show Preview Dialog with collected data
            report_data['settings'] = self.get_current_settings()
            PreviewDialog(self, report_data)
            self.update_status("Preview report generated.")

        except Exception as e:
            self.log_error(f"Error generating preview: {str(e)}")
            messagebox.showerror("Error", f"Failed to generate preview: {str(e)}")
            self.update_status("Preview failed")

    def get_current_settings(self) -> dict:
        """Returns a dictionary of the current settings for the preview dialog."""
        return {
            "Source File": os.path.basename(self.source_file.get()),
            "Destination File": os.path.basename(self.dest_file.get()),
            "Sort Column": self.sort_column.get() or "None",
            "---": "---", # Separator
            "Start Write Row": self.dest_write_start_row.get(),
            "End Write Row": self.dest_write_end_row.get() or "Unlimited",
            "Skip Rows": self.dest_skip_rows.get() or "None",
            "Respect Protection": "Yes" if self.respect_cell_protection.get() else "No",
            "Respect Formulas": "Yes" if self.respect_formulas.get() else "No",
        }

    def show_debug_info(self):
        pass
    
    def run(self):
        """Start the application"""
        try:
            self.root.mainloop()
        except Exception as e:
            self.log_error(f"Critical error in main loop: {str(e)}")
            messagebox.showerror("Critical Error", f"Application encountered a critical error: {str(e)}")

class PreviewDialog(tk.Toplevel):
    """A comprehensive simulation report dialog."""
    def __init__(self, parent, data: dict):
        super().__init__(parent.root)
        self.title("Transfer Simulation Report")
        self.geometry("650x750")
        self.transient(parent.root)
        self.grab_set()
        
        main_frame = ttk_boot.Frame(self, padding=10)
        main_frame.pack(fill=BOTH, expand=True)

        # --- Handle potential errors first ---
        if "error" in data:
            error_label = ttk_boot.Label(main_frame, text=f"❌ CRITICAL ERROR\n\n{data['error']}", bootstyle=DANGER, font=("Segoe UI", 12, "bold"), justify=CENTER)
            error_label.pack(pady=20, padx=10)
            ok_button = ttk_boot.Button(main_frame, text="Close", command=self.destroy, bootstyle="outline-danger")
            ok_button.pack(pady=10)
            return

        # --- 1. The Verdict ---
        verdict_frame = ttk_boot.LabelFrame(main_frame, text="Verdict", padding=10)
        verdict_frame.pack(padx=10, pady=5, fill=X)
        
        source_count = data['source_row_count']
        available_slots = data['available_slots']
        
        if available_slots == "Unlimited" or source_count <= available_slots:
            verdict_text = "✅ PERFECT"
            verdict_details = f"All {source_count} source rows will be transferred successfully."
            verdict_style = SUCCESS
        else:
            verdict_text = "⚠️ WARNING"
            verdict_details = f"Only {available_slots} of {source_count} source rows will be transferred. {source_count - available_slots} rows will be SKIPPED due to lack of space."
            verdict_style = WARNING

        ttk_boot.Label(verdict_frame, text=verdict_text, font=("Segoe UI", 14, "bold"), bootstyle=verdict_style).pack()
        ttk_boot.Label(verdict_frame, text=verdict_details, wraplength=600).pack(pady=(5,0))

        # --- 2. Data Flow Analysis ---
        analysis_frame = ttk_boot.LabelFrame(main_frame, text="Data Flow Analysis", padding=10)
        analysis_frame.pack(padx=10, pady=5, fill=X)

        ttk_boot.Label(analysis_frame, text=f"Source data to transfer:", font="-weight bold").grid(row=0, column=0, sticky=W)
        ttk_boot.Label(analysis_frame, text=f"{source_count} rows").grid(row=0, column=1, sticky=W, padx=5)
        
        ttk_boot.Separator(analysis_frame, orient=HORIZONTAL).grid(row=1, column=0, columnspan=2, sticky=EW, pady=5)

        ttk_boot.Label(analysis_frame, text=f"Destination Write Zone:", font="-weight bold").grid(row=2, column=0, sticky=W)
        ttk_boot.Label(analysis_frame, text=f"From row {data['start_row']} to {data['end_row']}").grid(row=2, column=1, sticky=W, padx=5)
        
        ttk_boot.Label(analysis_frame, text=f"  Total rows in zone:").grid(row=3, column=0, sticky=E, pady=(5,0))
        ttk_boot.Label(analysis_frame, text=f"{data['total_zone_rows']}").grid(row=3, column=1, sticky=W, padx=5, pady=(5,0))
        
        ttk_boot.Label(analysis_frame, text=f"  (-) Rows skipped by user:").grid(row=4, column=0, sticky=E)
        ttk_boot.Label(analysis_frame, text=f"{data['user_skipped_count']}").grid(row=4, column=1, sticky=W, padx=5)
        
        ttk_boot.Label(analysis_frame, text=f"  (-) Rows skipped (protected/formula):").grid(row=5, column=0, sticky=E)
        ttk_boot.Label(analysis_frame, text=f"{data['protected_skipped_count']}").grid(row=5, column=1, sticky=W, padx=5)
        
        ttk_boot.Separator(analysis_frame, orient=HORIZONTAL).grid(row=6, column=0, columnspan=2, sticky=EW, pady=5)
        
        ttk_boot.Label(analysis_frame, text=f"(=) Available Write Slots:", font="-weight bold").grid(row=7, column=0, sticky=E)
        ttk_boot.Label(analysis_frame, text=f"{data['available_slots']}", font="-weight bold").grid(row=7, column=1, sticky=W, padx=5)

        # --- 3. Settings Confirmation ---
        settings_frame = ttk_boot.LabelFrame(main_frame, text="Settings Used for Simulation", padding=10)
        settings_frame.pack(padx=10, pady=5, fill=BOTH, expand=True)

        row = 0
        for key, value in data['settings'].items():
            ttk_boot.Label(settings_frame, text=f"{key}", font="-weight bold").grid(row=row, column=0, sticky=W, padx=5, pady=1)
            ttk_boot.Label(settings_frame, text=value).grid(row=row, column=1, sticky=W, padx=5, pady=1)
            if key == "---":
                ttk_boot.Separator(settings_frame, orient=HORIZONTAL).grid(row=row, column=0, columnspan=2, sticky=EW, pady=3)
            row += 1
        settings_frame.columnconfigure(1, weight=1)
        
        # --- Close Button ---
        ok_button = ttk_boot.Button(main_frame, text="Close", command=self.destroy, bootstyle="outline-secondary")
        ok_button.pack(pady=10)

class DetectionConfigDialog(tk.Toplevel):
    """Dialog to configure end-row detection keywords."""
    def __init__(self, parent):
        super().__init__(parent.root)
        self.parent = parent
        self.title("Configure Detection Keywords")
        self.geometry("500x150")
        self.transient(parent.root)
        self.grab_set()

        # Temporary variable for editing
        self.temp_keywords = tk.StringVar(value=self.parent.detection_keywords.get())

        main_frame = ttk_boot.Frame(self, padding=15)
        main_frame.pack(fill=BOTH, expand=True)

        ttk_boot.Label(main_frame, text="Enter keywords to detect the 'total' row, separated by commas:").pack(anchor=W)
        
        entry = ttk_boot.Entry(main_frame, textvariable=self.temp_keywords)
        entry.pack(fill=X, pady=5)
        
        button_frame = ttk_boot.Frame(main_frame)
        button_frame.pack(fill=X, pady=10)

        save_button = ttk_boot.Button(button_frame, text="Save", command=self.save, bootstyle=SUCCESS)
        save_button.pack(side=RIGHT, padx=5)
        
        cancel_button = ttk_boot.Button(button_frame, text="Cancel", command=self.destroy, bootstyle="secondary")
        cancel_button.pack(side=RIGHT)

    def save(self):
        self.parent.detection_keywords.set(self.temp_keywords.get())
        self.parent.log_info(f"Detection keywords updated to: {self.temp_keywords.get()}")
        self.destroy()

if __name__ == "__main__":
    try:
        app = ExcelDataMapper()
        app.run()
    except Exception as e:
        logging.critical(f"Failed to start application: {str(e)}")
        print(f"Failed to start application: {str(e)}")
        input("Press Enter to exit...")