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
from logic.parser import ExcelParser

# Cấu hình logging
logging.basicConfig(
    filename='app.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)

class ExcelDataMapper:
    def __init__(self):
        self.root = ttk_boot.Window(themename="flatly")
        self.root.title("Excel Data Mapper")
        self.root.geometry("900x700")
        
        # Icon handling for PyInstaller
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
        self.current_theme = "flatly"
        
        # Data storage
        self.source_columns = {} # {name: index}
        self.dest_columns = {}   # {name: index}
        self.column_mappings = {}
        self.mapping_widgets = []
        self.mapping_combos = {}  # Khởi tạo sớm để tránh lỗi hasattr
        
        # Configuration
        self.config_dir = Path("configs")
        self.config_dir.mkdir(exist_ok=True)
        
        self.setup_gui()
        self.setup_menu()
        
        # Load last configuration if exists
        self.load_last_config()
        
    def setup_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open Destination Folder", command=self.open_dest_folder)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        # Settings menu
        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Settings", menu=settings_menu)
        settings_menu.add_command(label="Switch Theme", command=self.toggle_theme)
        
        # About menu
        about_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="About", menu=about_menu)
        about_menu.add_command(label="Info", command=self.show_about)
        
    def setup_gui(self):
        # Main container
        main_frame = ttk_boot.Frame(self.root, padding=10)
        main_frame.pack(fill=BOTH, expand=True)
        
        # File selection section
        file_frame = ttk_boot.LabelFrame(main_frame, text="File Selection", padding=10)
        file_frame.pack(fill=X, pady=(0, 10))
        
        # Source file
        ttk_boot.Label(file_frame, text="Source File:").grid(row=0, column=0, sticky=W, pady=2)
        ttk_boot.Entry(file_frame, textvariable=self.source_file, width=50).grid(row=0, column=1, padx=5, sticky=EW)
        ttk_boot.Button(file_frame, text="Browse", command=self.browse_source_file).grid(row=0, column=2, padx=5)
        
        # Destination file
        ttk_boot.Label(file_frame, text="Destination File:").grid(row=1, column=0, sticky=W, pady=2)
        ttk_boot.Entry(file_frame, textvariable=self.dest_file, width=50).grid(row=1, column=1, padx=5, sticky=EW)
        ttk_boot.Button(file_frame, text="Browse", command=self.browse_dest_file).grid(row=1, column=2, padx=5)
        
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
        
        ttk_boot.Button(header_frame, text="Load Columns", command=self.load_columns, bootstyle=INFO).grid(row=0, column=10, padx=20)
        
        # Column mapping section
        self.mapping_frame = ttk_boot.LabelFrame(main_frame, text="Column Mapping", padding=10)
        self.mapping_frame.pack(fill=BOTH, expand=True, pady=(0, 10))
        
        # Create scrollable frame for mappings
        self.setup_mapping_scroll()
        
        # Sort configuration
        sort_frame = ttk_boot.LabelFrame(main_frame, text="Sort Configuration", padding=10)
        sort_frame.pack(fill=X, pady=(0, 10))
        
        ttk_boot.Label(sort_frame, text="Sort by Column (optional):").grid(row=0, column=0, sticky=W, pady=2)
        self.sort_combo = ttk_boot.Combobox(sort_frame, textvariable=self.sort_column, width=30)
        self.sort_combo.grid(row=0, column=1, padx=5, sticky=W)
        
        # Action buttons
        action_frame = ttk_boot.Frame(main_frame)
        action_frame.pack(fill=X, pady=(0, 10))
        
        ttk_boot.Button(action_frame, text="Save Configuration", command=self.save_config, bootstyle=SUCCESS).pack(side=LEFT, padx=5)
        ttk_boot.Button(action_frame, text="Load Configuration", command=self.load_config, bootstyle=INFO).pack(side=LEFT, padx=5)
        ttk_boot.Button(action_frame, text="Execute Transfer", command=self.execute_transfer, bootstyle=PRIMARY).pack(side=RIGHT, padx=5)
        
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
    
    def load_columns(self, saved_sort_col: Optional[str] = None, apply_suggestions: bool = True):
        """Load column headers from both files and optionally set the sort column."""
        try:
            if not self.source_file.get() or not self.dest_file.get():
                messagebox.showwarning("Warning", "Please select both source and destination files first.")
                return
            
            # Kiểm tra file tồn tại
            if not os.path.exists(self.source_file.get()):
                messagebox.showerror("Error", f"Source file does not exist: {self.source_file.get()}")
                return
                
            if not os.path.exists(self.dest_file.get()):
                messagebox.showerror("Error", f"Destination file does not exist: {self.dest_file.get()}")
                return
            
            self.update_status("Loading columns...")
            
            # Load source columns
            self.source_columns = self.get_excel_columns(self.source_file.get(), self.source_header_start_row.get(), self.source_header_end_row.get())
            
            # Load destination columns  
            self.dest_columns = self.get_excel_columns(self.dest_file.get(), self.dest_header_start_row.get(), self.dest_header_end_row.get())
            
            if not self.source_columns or not self.dest_columns:
                messagebox.showerror("Error", "Could not load columns. Please check file paths and header row numbers.")
                return
            
            # Update sort combo
            source_keys = list(self.source_columns.keys())
            self.sort_combo['values'] = source_keys
            self.root.update_idletasks() # Force GUI update

            # Now, safely set the saved sort column value
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
    
    def _get_value_from_merged_cell(self, worksheet, row, col):
        """Helper to get value from a cell, resolving merged cells."""
        cell = worksheet.cell(row=row, column=col)
        if isinstance(cell, MergedCell):
            # It's a merged cell, find the top-left parent
            for merged_range in worksheet.merged_cells.ranges:
                if cell.coordinate in merged_range:
                    # Return the value of the top-left cell in the range
                    return worksheet.cell(row=merged_range.min_row, column=merged_range.min_col).value
        return cell.value

    def get_excel_columns(self, file_path, start_row, end_row):
        """Extracts headers using the centralized ExcelParser, filtering out empty or whitespace-only column names."""
        try:
            with ExcelParser(file_path) as parser:
                headers = parser.get_headers(start_row, end_row)
                
                # Filter out headers that are None, empty, or whitespace-only
                filtered_headers = {
                    name: index for name, index in headers.items()
                    if name and str(name).strip()
                }
                return filtered_headers
        except Exception as e:
            self.log_error(f"Error reading Excel columns with parser: {str(e)}")
            raise
    
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
            source_label = ttk_boot.Label(self.scrollable_frame, text=source_col_name, anchor=W)
            source_label.grid(row=i, column=0, sticky=EW, padx=5, pady=2)
            
            # Arrow
            arrow_label = ttk_boot.Label(self.scrollable_frame, text="→")
            arrow_label.grid(row=i, column=1, sticky=W, padx=5)
            
            # Destination column combobox
            dest_combo = ttk_boot.Combobox(self.scrollable_frame, values=[""] + list(self.dest_columns.keys()))
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

            # Stage 2: Load and apply the theme
            new_theme = config.get("theme", "flatly")
            if new_theme != self.current_theme:
                self.root.style.theme_use(new_theme)
                self.current_theme = new_theme

            # Stage 3: Load columns and pass the saved sort column to be set internally
            if self.source_file.get() and self.dest_file.get():
                saved_sort_col = config.get("sort_column", "")
                # Load columns WITHOUT applying suggestions, as they will be loaded from config
                self.load_columns(saved_sort_col=saved_sort_col, apply_suggestions=False)

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

                # Load and apply theme from last config
                new_theme = config.get("theme", "flatly")
                if new_theme != self.current_theme:
                    self.root.style.theme_use(new_theme)
                    self.current_theme = new_theme
                
        except Exception as e:
            self.log_error(f"Error loading last config: {str(e)}")
    
    def execute_transfer(self):
        """Execute the data transfer operation"""
        try:
            # Validation
            if not self.source_file.get() or not self.dest_file.get():
                messagebox.showwarning("Warning", "Please select both source and destination files.")
                return
            
            if not hasattr(self, 'mapping_combos') or not self.mapping_combos:
                messagebox.showwarning("Warning", "Please load columns first.")
                return
            
            # Get mappings
            mappings = {}
            for source_col, combo in self.mapping_combos.items():
                dest_col = combo.get()
                if dest_col:
                    mappings[source_col] = dest_col
            
            if not mappings:
                messagebox.showwarning("Warning", "Please configure at least one column mapping.")
                return
            
            # Validate mappings
            duplicate_destinations = []
            dest_values = list(mappings.values())
            for dest_col in set(dest_values):
                if dest_values.count(dest_col) > 1:
                    duplicate_destinations.append(dest_col)
            
            if duplicate_destinations:
                messagebox.showerror("Error", f"Duplicate destination columns detected: {', '.join(duplicate_destinations)}")
                return
            
            # Check file paths length
            if len(self.source_file.get()) > 255 or len(self.dest_file.get()) > 255:
                messagebox.showwarning("Warning", "File paths are very long (>255 characters). This might cause issues.")
            
            self.update_status("Starting data transfer...")
            self.progress['value'] = 0
            self.root.update()
            
            # Perform transfer
            self.perform_data_transfer(mappings)
            
            self.progress['value'] = 100
            self.update_status("Transfer completed successfully")
            
            messagebox.showinfo("Success", "Data transfer completed successfully!")
            
            # Open destination folder
            if messagebox.askyesno("Open Folder", "Would you like to open the destination folder?"):
                self.open_dest_folder()
            
        except Exception as e:
            self.log_error(f"Error in execute_transfer: {str(e)}")
            self.log_error(traceback.format_exc())
            messagebox.showerror("Error", f"Transfer failed: {str(e)}")
            self.update_status("Transfer failed")
            self.progress['value'] = 0
    
    def perform_data_transfer(self, mappings):
        """Perform the actual data transfer"""
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
        """Reads data from the source file based on the pre-parsed logical headers."""
        try:
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
            
            workbook.close()
            return data
            
        except Exception as e:
            self.log_error(f"Error reading source data: {str(e)}")
            raise
    
    def write_to_destination(self, source_data, mappings):
        """Write data to destination Excel file, handling merged cells correctly and robustly."""
        try:
            workbook = openpyxl.load_workbook(self.dest_file.get())
            worksheet = workbook.active

            dest_headers_map = self.dest_columns
            start_row = self.dest_header_end_row.get() + 1

            def get_writable_cell(row_idx, col_idx):
                """Resolves merged cells to find the top-left anchor cell which is writable."""
                cell = worksheet.cell(row=row_idx, column=col_idx)
                if not isinstance(cell, MergedCell):
                    return cell

                for merged_range in worksheet.merged_cells.ranges:
                    if (merged_range.min_row <= row_idx <= merged_range.max_row and
                        merged_range.min_col <= col_idx <= merged_range.max_col):
                        return worksheet.cell(row=merged_range.min_row, column=merged_range.min_col)
                
                return cell # Fallback

            # --- 1. Clear existing data ---
            max_data_row = worksheet.max_row
            if max_data_row >= start_row:
                cleared_anchors = set()
                for row_to_clear in range(start_row, max_data_row + 50):
                    for dest_col_num in dest_headers_map.values():
                        anchor_cell = get_writable_cell(row_to_clear, dest_col_num)
                        if anchor_cell.row >= start_row and anchor_cell.coordinate not in cleared_anchors:
                            if anchor_cell.data_type != 'f':
                                anchor_cell.value = None
                            cleared_anchors.add(anchor_cell.coordinate)

            # --- 2. Write new data ---
            for i, row_data in enumerate(source_data):
                current_row = start_row + i

                if i > 0 and i % 10 == 0:
                    progress = 50 + (i / len(source_data)) * 45
                    self.progress['value'] = progress
                    self.root.update()

                for source_col_name, dest_col_name in mappings.items():
                    if dest_col_name in dest_headers_map:
                        dest_col_num = dest_headers_map[dest_col_name]
                        source_value = row_data.get(source_col_name, "")

                        cell_to_write = get_writable_cell(current_row, dest_col_num)
                        
                        # CRITICAL FIX: If the anchor cell is above the intended row (due to a downward merge
                        # from the header), skip writing to prevent corrupting the header.
                        if cell_to_write.row < current_row:
                            continue

                        if cell_to_write.data_type == 'f':
                            continue

                        cell_to_write.value = source_value

            workbook.save(self.dest_file.get())
            workbook.close()

        except Exception as e:
            self.log_error(f"Error writing to destination: {str(e)}")
            self.log_error(traceback.format_exc())
            raise

    
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
        """Show about dialog"""
        about_text = """Excel Data Mapper v1.0

A powerful tool for mapping and transferring data between Excel files while preserving formatting.

Features:
• Flexible column mapping
• Sort data before transfer
• Preserve Excel formatting and styles
• Configuration save/load
• Theme switching
• Comprehensive error handling

Developed with Python + ttkbootstrap
"""
        messagebox.showinfo("About Excel Data Mapper", about_text)
    
    def update_status(self, message):
        """Update status bar message"""
        self.status_label.config(text=message)
        self.root.update()
    
    def log_info(self, message):
        """Log info message"""
        logging.info(message)
    
    def log_error(self, message):
        """Log error message"""
        logging.error(message)

    def show_debug_info(self):
        pass
    
    def run(self):
        """Start the application"""
        try:
            self.root.mainloop()
        except Exception as e:
            self.log_error(f"Critical error in main loop: {str(e)}")
            messagebox.showerror("Critical Error", f"Application encountered a critical error: {str(e)}")

if __name__ == "__main__":
    try:
        app = ExcelDataMapper()
        app.run()
    except Exception as e:
        logging.critical(f"Failed to start application: {str(e)}")
        print(f"Failed to start application: {str(e)}")
        input("Press Enter to exit...")
