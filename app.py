import tkinter as tk
from tkinter import ttk, filedialog, simpledialog
import ttkbootstrap as ttk_boot
from ttkbootstrap.constants import *
import json
import os
import logging
from pathlib import Path
from typing import Optional
import openpyxl
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
from gui.widgets import (ScrollableFrame, AboutDialog, PreviewDialog, 
                         DetectionConfigDialog, show_custom_info, 
                         show_custom_error, show_custom_warning, 
                         show_custom_question)

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
        for _ in range(3):
            gc.collect()
            time.sleep(0.05)
    
    @staticmethod
    def is_file_locked(file_path: str) -> bool:
        """Check if a file is currently locked by another process"""
        try:
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
                                processes.append({'pid': proc.info['pid'], 'name': proc.info['name']})
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
        self.root.geometry("1000x850")
        
        self.icon_path = None
        try:
            # Correctly determine the base path for PyInstaller
            base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
            icon_path = os.path.join(base_path, "icon.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
                self.icon_path = icon_path
        except Exception:
            # If icon loading fails, do nothing and proceed
            pass
        
        self.source_file = tk.StringVar()
        self.dest_file = tk.StringVar()
        self.source_header_start_row = tk.IntVar(value=1)
        self.source_header_end_row = tk.IntVar(value=1)
        self.dest_header_start_row = tk.IntVar(value=9)
        self.dest_header_end_row = tk.IntVar(value=9)
        self.sort_column = tk.StringVar()
        self.current_theme = "flatly"
        self.dest_write_start_row = tk.IntVar(value=11)
        self.dest_write_end_row = tk.IntVar(value=0)
        self.dest_skip_rows = tk.StringVar(value="")
        self.respect_cell_protection = tk.BooleanVar(value=True)
        self.respect_formulas = tk.BooleanVar(value=True)
        self.detection_keywords = tk.StringVar(value="total,sum,cộng,tổng,thành tiền")
        
        self.source_columns = {}
        self.dest_columns = {}
        self.mapping_combos = {}
        
        self.config_dir = Path("configs")
        self.config_dir.mkdir(exist_ok=True)
        
        self.setup_menu()
        self.setup_gui()
        self.load_last_config()
        
    def setup_menu(self):
        menubar_frame = ttk_boot.Frame(self.root)
        menubar_frame.pack(fill=X, side=TOP, padx=5, pady=(2, 0))

        def create_menu_button(text, menu):
            button = ttk_boot.Button(menubar_frame, text=text, bootstyle="link")
            button.pack(side=LEFT, padx=10)
            button.config(command=lambda: menu.post(button.winfo_rootx(), button.winfo_rooty() + button.winfo_height()))
            return button

        file_menu = tk.Menu(self.root, tearoff=0)
        file_menu.add_command(label="Open Destination Folder", command=self.open_dest_folder)
        file_menu.add_command(label="Force Release File Handles", command=self.force_release_excel_handles)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        create_menu_button("File", file_menu)

        settings_menu = tk.Menu(self.root, tearoff=0)
        settings_menu.add_command(label="Switch Theme", command=self.toggle_theme)
        settings_menu.add_command(label="Configure Detection...", command=self.open_detection_config_dialog)
        create_menu_button("Settings", settings_menu)

        about_menu = tk.Menu(self.root, tearoff=0)
        about_menu.add_command(label="Info", command=self.show_about)
        create_menu_button("About", about_menu)
        
    def setup_gui(self):
        main_frame = ttk_boot.Frame(self.root, padding=10)
        main_frame.pack(fill=BOTH, expand=True, side=TOP)
        
        file_frame = ttk_boot.LabelFrame(main_frame, text="File Selection", padding=10)
        file_frame.pack(fill=X, pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        ttk_boot.Label(file_frame, text="Source File:").grid(row=0, column=0, sticky=W, pady=2)
        ttk_boot.Entry(file_frame, textvariable=self.source_file, width=50).grid(row=0, column=1, padx=5, pady=5, sticky=EW)
        ttk_boot.Button(file_frame, bootstyle="outline", text="Browse", command=self.browse_source_file).grid(row=0, column=2, padx=5, pady=5)
        ttk_boot.Label(file_frame, text="Destination File:").grid(row=1, column=0, sticky=W, pady=2)
        ttk_boot.Entry(file_frame, textvariable=self.dest_file, width=50).grid(row=1, column=1, padx=5, pady=2, sticky=EW)
        ttk_boot.Button(file_frame, bootstyle="outline", text="Browse", command=self.browse_dest_file).grid(row=1, column=2, padx=5, pady=2)
        
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
        
        write_zone_frame = ttk_boot.LabelFrame(main_frame, text="Setting write zone", padding=10)
        write_zone_frame.pack(fill=X, pady=(0, 10))
        write_zone_frame.columnconfigure(1, weight=1)
        write_zone_frame.columnconfigure(3, weight=1)
        write_zone_frame.columnconfigure(4, weight=2)
        ttk_boot.Label(write_zone_frame, text="Start Write Row:").grid(row=0, column=0, sticky=W, padx=5, pady=5)
        ttk_boot.Spinbox(write_zone_frame, from_=1, to=99999, textvariable=self.dest_write_start_row, width=8).grid(row=0, column=1, sticky=W, padx=5)
        ttk_boot.Label(write_zone_frame, text="End Write Row:").grid(row=0, column=2, sticky=W, padx=(20, 5), pady=5)
        ttk_boot.Spinbox(write_zone_frame, from_=0, to=99999, textvariable=self.dest_write_end_row, width=8).grid(row=0, column=3, sticky=W, padx=5)
        self.detect_button = ttk_boot.Button(write_zone_frame, text="Detect Zone", command=self.detect_write_zone, bootstyle="outline-info")
        self.detect_button.grid(row=0, column=4, padx=(10, 5), pady=5, sticky=E)
        ttk_boot.Label(write_zone_frame, text="Skip Rows (e.g., 15, 20-25):").grid(row=1, column=0, sticky=W, padx=5, pady=5)
        ttk_boot.Entry(write_zone_frame, textvariable=self.dest_skip_rows).grid(row=1, column=1, columnspan=4, padx=5, sticky=EW)
        ttk_boot.Checkbutton(write_zone_frame, text="Respect cell protection", variable=self.respect_cell_protection).grid(row=2, column=0, columnspan=2, sticky=W, padx=5, pady=5)
        ttk_boot.Checkbutton(write_zone_frame, text="Respect formulas", variable=self.respect_formulas).grid(row=2, column=2, columnspan=3, sticky=W, padx=(20, 5), pady=5)

        mapping_container = ttk_boot.LabelFrame(main_frame, text="Column Mapping", padding=10)
        mapping_container.pack(fill=BOTH, expand=True, pady=(0, 10))
        self.mapping_scroll_frame = ScrollableFrame(mapping_container)
        self.mapping_scroll_frame.pack(fill=BOTH, expand=True)
        
        sort_frame = ttk_boot.LabelFrame(main_frame, text="Sort Configuration", padding=10)
        sort_frame.pack(fill=X, pady=(0, 10))
        ttk_boot.Label(sort_frame, text="Sort by Column (optional):").grid(row=0, column=0, sticky=W, pady=2)
        self.sort_combo = ttk_boot.Combobox(sort_frame, textvariable=self.sort_column, width=50)
        self.sort_combo.grid(row=0, column=1, padx=5, sticky=W)
        
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

        # Pack label first to give it priority and prevent it from being obscured
        self.status_label = ttk_boot.Label(self.status_frame, text="Ready")
        self.status_label.pack(side=RIGHT, padx=(5, 0))

        # Pack progress bar second to fill the remaining space
        self.progress = ttk_boot.Progressbar(self.status_frame, mode='determinate', bootstyle=SUCCESS)
        self.progress.pack(side=LEFT, fill=X, expand=True, padx=(0, 5))
    
    def browse_source_file(self):
        filename = filedialog.askopenfilename(title="Select Source Excel file", filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            self.source_file.set(filename)
            self.log_info(f"Source file selected: {filename}")
    
    def browse_dest_file(self):
        filename = filedialog.askopenfilename(title="Select Destination Excel file", filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            self.dest_file.set(filename)
            self.log_info(f"Destination file selected: {filename}")
    
    def force_release_excel_handles(self):
        try:
            self.source_columns, self.dest_columns = {}, {}
            FileHandleManager.force_release_handles()
            self.update_status("File handles released")
            self.log_info("Forced release of Excel file handles")
            show_custom_info(self.root, self, "Info", "Excel file handles have been released.")
        except Exception as e:
            self.log_error(f"Error forcing handle release: {str(e)}")
            show_custom_error(self.root, self, "Error", f"Error releasing handles: {str(e)}")
    
    def check_file_accessibility(self, file_path: str) -> bool:
        try:
            if not os.path.exists(file_path): return False
            if FileHandleManager.is_file_locked(file_path):
                self.update_status(f"Waiting for file to be released: {os.path.basename(file_path)}")
                if not FileHandleManager.wait_for_file_release(file_path, max_wait_seconds=10):
                    processes = FileHandleManager.get_processes_using_file(file_path)
                    if processes:
                        process_names = [p['name'] for p in processes]
                        self.log_error(f"File locked by processes: {', '.join(process_names)}")
                        show_custom_warning(self.root, self, "File Locked", f"File is locked by: {', '.join(process_names)}\nPlease close these applications and try again.")
                    return False
            return True
        except Exception as e:
            self.log_error(f"Error checking file accessibility: {str(e)}")
            return False
    
    def get_excel_columns(self, file_path, start_row, end_row):
        try:
            FileHandleManager.force_release_handles()
            with ExcelParser(file_path) as parser:
                headers = parser.get_headers(start_row, end_row)
                return {name: index for name, index in headers.items() if name and str(name).strip()}
        except Exception as e:
            self.log_error(f"Error reading Excel columns with parser: {str(e)}")
            raise
        finally:
            FileHandleManager.force_release_handles()

    def safe_load_columns(self, saved_sort_col: Optional[str] = None, apply_suggestions: bool = True):
        try:
            if not self.source_file.get() or not self.dest_file.get():
                show_custom_warning(self.root, self, "Warning", "Please select both source and destination files first.")
                return
            if not self.check_file_accessibility(self.source_file.get()):
                show_custom_error(self.root, self, "Error", f"Cannot access source file: {self.source_file.get()}")
                return
            if not self.check_file_accessibility(self.dest_file.get()):
                show_custom_error(self.root, self, "Error", f"Cannot access destination file: {self.dest_file.get()}")
                return
            
            self.force_release_excel_handles()
            self.update_status("Loading columns...")
            
            self.source_columns = self.get_excel_columns(self.source_file.get(), self.source_header_start_row.get(), self.source_header_end_row.get())
            time.sleep(0.1)
            self.dest_columns = self.get_excel_columns(self.dest_file.get(), self.dest_header_start_row.get(), self.dest_header_end_row.get())
            
            if not self.source_columns or not self.dest_columns:
                show_custom_error(self.root, self, "Error", "Could not load columns. Please check file paths and header row numbers.")
                return
            
            source_keys = list(self.source_columns.keys())
            self.sort_combo['values'] = source_keys
            if saved_sort_col and saved_sort_col in source_keys:
                self.sort_column.set(saved_sort_col)
            
            self.create_mapping_widgets(apply_suggestions=apply_suggestions)
            self.update_status(f"Loaded {len(self.source_columns)} source and {len(self.dest_columns)} destination columns")
            self.log_info("Columns loaded successfully")
        except Exception as e:
            self.log_error(f"Error loading columns: {str(e)}")
            show_custom_error(self.root, self, "Error", f"Failed to load columns: {str(e)}")
            self.update_status("Error loading columns")
        finally:
            FileHandleManager.force_release_handles()
    
    def create_mapping_widgets(self, apply_suggestions: bool = True):
        for widget in self.mapping_scroll_frame.scrollable_frame.winfo_children():
            widget.destroy()
        
        self.mapping_scroll_frame.scrollable_frame.columnconfigure(0, weight=1)
        self.mapping_scroll_frame.scrollable_frame.columnconfigure(2, weight=1)
        ttk_boot.Label(self.mapping_scroll_frame.scrollable_frame, text="Source Column", font="-weight bold").grid(row=0, column=0, sticky=W, padx=5, pady=(5, 10))
        ttk_boot.Label(self.mapping_scroll_frame.scrollable_frame, text="Destination Column", font="-weight bold").grid(row=0, column=2, sticky=W, padx=5, pady=(5, 10))
        
        self.mapping_combos = {}
        for i, source_col_name in enumerate(self.source_columns.keys(), start=1):
            ttk_boot.Label(self.mapping_scroll_frame.scrollable_frame, text=source_col_name, anchor=W).grid(row=i, column=0, sticky=EW, padx=5, pady=2)
            ttk_boot.Label(self.mapping_scroll_frame.scrollable_frame, text="→").grid(row=i, column=1, sticky=W, padx=5)
            dest_combo = ttk_boot.Combobox(self.mapping_scroll_frame.scrollable_frame, values=[""] + list(self.dest_columns.keys()))
            dest_combo.grid(row=i, column=2, sticky=EW, padx=5, pady=2)
            if apply_suggestions:
                suggested = self.suggest_mapping(source_col_name, list(self.dest_columns.keys()))
                if suggested: dest_combo.set(suggested)
            self.mapping_combos[source_col_name] = dest_combo
    
    def suggest_mapping(self, source_col, dest_cols):
        if str(source_col).startswith('Column_'): return ""
        import re
        def normalize_and_tokenize(text: str) -> set:
            text = re.sub(r'[\s\u3000_\-]+', ' ', text)
            text = re.sub(r'[()[\\]{}]', '', text)
            return set(text.lower().strip().split())
        source_tokens = normalize_and_tokenize(source_col)
        best_match, max_score = "", 0
        keywords_map = {'content': 'contents', 'purpose': 'purpose', 'amount': 'amount', 'vat': 'vat', 'currency': 'currency', 'date': 'trading date', 'no': 'no.', 'number': 'no.', 'code': 'code reference', 'total': 'sub total'}
        for dest_col in dest_cols:
            current_score = 0
            dest_tokens = normalize_and_tokenize(dest_col)
            if not dest_tokens: continue
            if source_tokens == dest_tokens: current_score = 100
            common_tokens = source_tokens.intersection(dest_tokens)
            current_score += len(common_tokens) * 50
            for key, value in keywords_map.items():
                if key in source_tokens and value in dest_tokens: current_score += 40
            source_norm_str, dest_norm_str = "".join(source_tokens), "".join(dest_tokens)
            if source_norm_str in dest_norm_str or dest_norm_str in source_norm_str: current_score += 20
            if current_score > max_score: max_score, best_match = current_score, dest_col
        return best_match
    
    def save_config(self):
        try:
            if not hasattr(self, 'mapping_combos') or not self.mapping_combos:
                show_custom_warning(self.root, self, "Warning", "No mappings to save. Please load columns first.")
                return
            config_file_path = filedialog.asksaveasfilename(title="Save Configuration As", defaultextension=".json", filetypes=[("JSON files", "*.json")])
            if not config_file_path: return
            mappings = {source_col: combo.get() for source_col, combo in self.mapping_combos.items() if combo.get()}
            config = {
                "source_file": self.source_file.get(), "dest_file": self.dest_file.get(),
                "source_header_start_row": self.source_header_start_row.get(), "source_header_end_row": self.source_header_end_row.get(),
                "dest_header_start_row": self.dest_header_start_row.get(), "dest_header_end_row": self.dest_header_end_row.get(),
                "dest_write_start_row": self.dest_write_start_row.get(), "dest_write_end_row": self.dest_write_end_row.get(),
                "dest_skip_rows": self.dest_skip_rows.get(), "respect_cell_protection": self.respect_cell_protection.get(),
                "respect_formulas": self.respect_formulas.get(), "detection_keywords": self.detection_keywords.get(),
                "sort_column": self.sort_column.get(), "theme": self.current_theme, "mapping": mappings,
                "created_date": datetime.now().isoformat()
            }
            with open(config_file_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            self.update_status(f"Configuration saved to {os.path.basename(config_file_path)}")
            self.log_info(f"Configuration saved: {config_file_path}")
        except Exception as e:
            self.log_error(f"Error saving configuration: {str(e)}")
            show_custom_error(self.root, self, "Error", f"Failed to save configuration: {str(e)}")
    
    def load_config(self):
        try:
            config_file_path = filedialog.askopenfilename(title="Load Configuration", filetypes=[("JSON files", "*.json")])
            if not config_file_path: return
            with open(config_file_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            self.source_file.set(config.get("source_file", ""))
            self.dest_file.set(config.get("dest_file", ""))
            self.source_header_start_row.set(config.get("source_header_start_row", 1))
            self.source_header_end_row.set(config.get("source_header_end_row", 1))
            self.dest_header_start_row.set(config.get("dest_header_start_row", 9))
            self.dest_header_end_row.set(config.get("dest_header_end_row", 9))
            self.dest_write_start_row.set(config.get("dest_write_start_row", self.dest_header_end_row.get() + 1))
            self.dest_write_end_row.set(config.get("dest_write_end_row", 0))
            self.dest_skip_rows.set(config.get("dest_skip_rows", ""))
            self.respect_cell_protection.set(config.get("respect_cell_protection", True))
            self.respect_formulas.set(config.get("respect_formulas", True))
            self.detection_keywords.set(config.get("detection_keywords", "total,sum,cộng,tổng,thành tiền"))
            new_theme = config.get("theme", "flatly")
            if new_theme != self.current_theme:
                self.root.style.theme_use(new_theme)
                self.current_theme = new_theme
            if self.source_file.get() and self.dest_file.get():
                saved_sort_col = config.get("sort_column", "")
                self.safe_load_columns(saved_sort_col=saved_sort_col, apply_suggestions=False)
                mappings = config.get("mapping", {})
                for source_col, dest_col in mappings.items():
                    if source_col in self.mapping_combos:
                        self.mapping_combos[source_col].set(dest_col)
            self.update_status(f"Configuration loaded from {os.path.basename(config_file_path)}")
            self.log_info(f"Configuration loaded: {config_file_path}")
        except Exception as e:
            self.log_error(f"Error in load_config: {str(e)}")
            show_custom_error(self.root, self, "Error", f"Failed to load configuration: {str(e)}")
    
    def load_last_config(self):
        try:
            config_files = list(self.config_dir.glob("*.json"))
            if not config_files: return
            latest_config = max(config_files, key=os.path.getctime)
            with open(latest_config, 'r', encoding='utf-8') as f:
                config = json.load(f)
            if os.path.exists(config.get("source_file", "")):
                self.source_file.set(config.get("source_file", ""))
            if os.path.exists(config.get("dest_file", "")):
                self.dest_file.set(config.get("dest_file", ""))
            self.source_header_start_row.set(config.get("source_header_start_row", 1))
            self.source_header_end_row.set(config.get("source_header_end_row", 1))
            self.dest_header_start_row.set(config.get("dest_header_start_row", 9))
            self.dest_header_end_row.set(config.get("dest_header_end_row", 9))
            self.dest_write_start_row.set(config.get("dest_write_start_row", self.dest_header_end_row.get() + 1))
            self.dest_write_end_row.set(config.get("dest_write_end_row", 0))
            self.dest_skip_rows.set(config.get("dest_skip_rows", ""))
            new_theme = config.get("theme", "flatly")
            if new_theme != self.current_theme:
                self.root.style.theme_use(new_theme)
                self.current_theme = new_theme
        except Exception as e:
            self.log_error(f"Error loading last config: {str(e)}")
    
    def execute_transfer(self):
        if not self.source_file.get() or not self.dest_file.get():
            show_custom_warning(self.root, self, "Warning", "Please select both source and destination files.")
            return
        if not hasattr(self, 'mapping_combos') or not self.mapping_combos:
            show_custom_warning(self.root, self, "Warning", "Please load columns first.")
            return
        mappings = {s: c.get() for s, c in self.mapping_combos.items() if c.get()}
        if not mappings:
            show_custom_warning(self.root, self, "Warning", "Please configure at least one column mapping.")
            return
        dest_values = list(mappings.values())
        duplicates = [d for d in set(dest_values) if dest_values.count(d) > 1]
        if duplicates:
            show_custom_error(self.root, self, "Error", f"Duplicate destination columns detected: {', '.join(duplicates)}")
            return
        if not self.check_file_accessibility(self.source_file.get()) or not self.check_file_accessibility(self.dest_file.get()):
            return

        self.disable_controls()
        self.update_status("Starting data transfer...")
        self.progress['value'] = 0
        transfer_thread = Thread(target=self._execute_transfer_thread, args=(mappings,))
        transfer_thread.daemon = True
        transfer_thread.start()

    def _execute_transfer_thread(self, mappings):
        try:
            self.perform_data_transfer(mappings)
            self.root.after(0, self.on_transfer_success)
        except Exception as e:
            self.root.after(0, self.on_transfer_error, e)
    
    def on_transfer_success(self):
        self.progress['value'] = 100
        self.update_status("Transfer completed successfully")
        self.enable_controls()
        show_custom_info(self.root, self, "Success", "Data transfer completed successfully!")
        if show_custom_question(self.root, self, "Open Folder", "Would you like to open the destination folder?"):
            self.open_dest_folder()

    def on_transfer_error(self, error):
        self.log_error(f"Error in execute_transfer: {str(error)}\n{traceback.format_exc()}")
        self.update_status("Transfer failed")
        self.progress['value'] = 0
        self.enable_controls()
        show_custom_error(self.root, self, "Error", f"Transfer failed: {str(error)}")

    def disable_controls(self):
        for widget in [self.execute_button, self.load_button, self.save_button, self.load_cols_button, self.preview_button]:
            widget.config(state=DISABLED)

    def enable_controls(self):
        for widget in [self.execute_button, self.load_button, self.save_button, self.load_cols_button, self.preview_button]:
            widget.config(state=NORMAL)
    
    def perform_data_transfer(self, mappings):
        dest_path = Path(self.dest_file.get())
        backup_path = dest_path.with_suffix('.backup' + dest_path.suffix)
        try:
            shutil.copy2(dest_path, backup_path)
            self.update_status("Reading source data..."); self.progress['value'] = 10; self.root.update()
            source_data = self.read_source_data()
            if not source_data: raise ValueError("No data found in source file")
            
            if self.sort_column.get() in mappings:
                self.update_status("Sorting data..."); self.progress['value'] = 30; self.root.update()
                sort_col_key = self.sort_column.get()
                source_data.sort(key=lambda x: (x.get(sort_col_key) is None or str(x.get(sort_col_key, "")).strip() == "", str(x.get(sort_col_key, ""))))
            
            self.update_status("Writing to destination..."); self.progress['value'] = 50; self.root.update()
            self.write_to_destination(source_data, mappings)
            if backup_path.exists(): backup_path.unlink()
            self.progress['value'] = 100
            self.log_info("Data transfer completed successfully")
        except Exception as e:
            if backup_path.exists():
                try:
                    shutil.copy2(backup_path, dest_path)
                    backup_path.unlink()
                except Exception as backup_e:
                    self.log_error(f"Failed to restore backup: {backup_e}")
            raise e
    
    def read_source_data(self):
        FileHandleManager.force_release_handles()
        workbook = None
        try:
            workbook = openpyxl.load_workbook(self.source_file.get(), data_only=True)
            worksheet = workbook.active
            start_data_row = self.source_header_end_row.get() + 1
            data = []
            for row_index in range(start_data_row, worksheet.max_row + 1):
                row_data, has_data = {}, False
                for header_name, col_index in self.source_columns.items():
                    value = worksheet.cell(row=row_index, column=col_index).value
                    if value is not None: has_data = True
                    row_data[header_name] = value
                if has_data: data.append(row_data)
            return data
        finally:
            if workbook:
                workbook.close()
            FileHandleManager.force_release_handles()
    
    def _parse_skip_rows(self, skip_rows_str: str) -> set:
        skipped_rows = set()
        if not skip_rows_str: return skipped_rows
        for part in skip_rows_str.split(','):
            part = part.strip()
            if not part: continue
            if '-' in part:
                try:
                    start, end = map(int, part.split('-'))
                    if start <= end: skipped_rows.update(range(start, end + 1))
                except ValueError: self.log_warning(f"Could not parse range in skip_rows: {part}")
            else:
                try: skipped_rows.add(int(part))
                except ValueError: self.log_warning(f"Could not parse number in skip_rows: {part}")
        return skipped_rows

    def write_to_destination(self, source_data, mappings):
        FileHandleManager.force_release_handles()
        workbook = None
        try:
            workbook = openpyxl.load_workbook(self.dest_file.get())
            worksheet = workbook.active
            start_write_row, end_write_row = self.dest_write_start_row.get(), self.dest_write_end_row.get()
            skipped_rows = self._parse_skip_rows(self.dest_skip_rows.get())
            respect_protection, respect_formulas = self.respect_cell_protection.get(), self.respect_formulas.get()
            if start_write_row <= self.dest_header_end_row.get():
                raise ValueError("Start Write Row must be after the destination header rows.")

            def get_writable_cell(row_idx, col_idx):
                cell = worksheet.cell(row=row_idx, column=col_idx)
                if isinstance(cell, MergedCell):
                    for merged_range in worksheet.merged_cells.ranges:
                        if cell.coordinate in merged_range:
                            return worksheet.cell(row=merged_range.min_row, column=merged_range.min_col)
                return cell

            clear_until_row = end_write_row if end_write_row > 0 else worksheet.max_row + 50
            cleared_anchors = set()
            for row_to_clear in range(start_write_row, clear_until_row + 1):
                if row_to_clear in skipped_rows: continue
                for dest_col_num in self.dest_columns.values():
                    anchor_cell = get_writable_cell(row_to_clear, dest_col_num)
                    if (anchor_cell.row >= start_write_row and anchor_cell.coordinate not in cleared_anchors and anchor_cell.row not in skipped_rows):
                        if not (respect_formulas and anchor_cell.data_type == 'f'):
                            anchor_cell.value = None
                        cleared_anchors.add(anchor_cell.coordinate)

            current_write_row = start_write_row
            for i, row_data in enumerate(source_data):
                while True:
                    if end_write_row > 0 and current_write_row > end_write_row:
                        self.log_warning(f"Reached end of write zone (row {end_write_row}). Stopping data transfer.")
                        show_custom_warning(self.root, self, "Write Limit Reached", f"Data transfer stopped at row {end_write_row} as configured.")
                        workbook.save(self.dest_file.get())
                        return
                    
                    is_invalid = current_write_row in skipped_rows
                    if not is_invalid and respect_protection and worksheet.protection.sheet:
                        for dest_col_num in self.dest_columns.values():
                            if get_writable_cell(current_write_row, dest_col_num).protection.locked:
                                is_invalid = True; break
                    if not is_invalid: break
                    current_write_row += 1

                if i > 0 and i % 10 == 0:
                    self.progress['value'] = 50 + (i / len(source_data)) * 45; self.root.update()

                for source_col, dest_col in mappings.items():
                    if dest_col in self.dest_columns:
                        dest_col_num = self.dest_columns[dest_col]
                        cell_to_write = get_writable_cell(current_write_row, dest_col_num)
                        if cell_to_write.row >= current_write_row and not (respect_formulas and cell_to_write.data_type == 'f'):
                            cell_to_write.value = row_data.get(source_col, "")
                current_write_row += 1
            workbook.save(self.dest_file.get())
        finally:
            if workbook:
                workbook.close()
            FileHandleManager.force_release_handles()

    def toggle_theme(self):
        new_theme = "superhero" if self.current_theme == "flatly" else "flatly"
        self.root.style.theme_use(new_theme)
        self.current_theme = new_theme
        self.update_status(f"Theme changed to {new_theme}")
    
    def open_dest_folder(self):
        if not self.dest_file.get():
            show_custom_warning(self.root, self, "Warning", "No destination file selected.")
            return
        dest_path = Path(self.dest_file.get())
        if dest_path.exists():
            folder_path = dest_path.parent
            try:
                if os.name == 'nt':
                    os.startfile(folder_path)
                elif sys.platform == 'darwin':
                    subprocess.run(['open', folder_path])
                else:
                    subprocess.run(['xdg-open', folder_path])
            except Exception as e:
                self.log_error(f"Could not open folder: {e}")
                show_custom_error(self.root, self, "Error", f"Could not open folder:\n{e}")
        else:
            show_custom_warning(self.root, self, "Warning", "Destination file does not exist.")
    
    def show_about(self):
        AboutDialog(self.root, self)
    
    def update_status(self, message):
        self.status_label.config(text=message)
        self.root.update_idletasks()
    
    def log_info(self, message): logging.info(message)
    def log_warning(self, message): logging.warning(message)
    def log_error(self, message): logging.error(message)

    def open_detection_config_dialog(self):
        DetectionConfigDialog(self.root, self)

    def detect_write_zone(self):
        if not self.dest_file.get() or not os.path.exists(self.dest_file.get()):
            show_custom_warning(self.root, self, "Warning", "Please select a valid destination file first.")
            return
        try:
            self.update_status("Detecting write zone...")
            predicted_start_row = self.dest_header_end_row.get() + 1
            self.dest_write_start_row.set(predicted_start_row)
            predicted_end_row = 0
            with ExcelParser(self.dest_file.get()) as p:
                ws = p.worksheet
                keywords = [k.strip() for k in self.detection_keywords.get().lower().split(',') if k.strip()]
                for row in range(predicted_start_row, ws.max_row + 1):
                    for col in range(1, ws.max_column + 1):
                        cell_val = ws.cell(row, col).value
                        if isinstance(cell_val, str) and any(k in cell_val.lower() for k in keywords):
                            predicted_end_row = row - 1; break
                    if predicted_end_row > 0: break
                if predicted_end_row == 0:
                    for row in range(predicted_start_row, ws.max_row + 1):
                        if all(ws.cell(row, col).value is None for col in range(1, ws.max_column + 1)):
                            predicted_end_row = row - 1; break
            if predicted_end_row >= predicted_start_row:
                self.dest_write_end_row.set(predicted_end_row)
                self.update_status(f"Detection complete. Start: {predicted_start_row}, End: {predicted_end_row}.")
            else:
                self.dest_write_end_row.set(0)
                self.update_status(f"Detection complete. Start: {predicted_start_row}, End: Unlimited.")
        except Exception as e:
            self.log_error(f"Error detecting write zone: {str(e)}")
            show_custom_error(self.root, self, "Error", f"Failed to detect write zone: {str(e)}")
            self.update_status("Detection failed")

    def _run_preview_simulation(self):
        report = {}
        try:
            with ExcelParser(self.source_file.get()) as p:
                report['source_row_count'] = p.count_data_rows(self.source_header_end_row.get())
            with ExcelParser(self.dest_file.get()) as p:
                ws = p.worksheet
                start_row, end_row = self.dest_write_start_row.get(), self.dest_write_end_row.get()
                if start_row <= self.dest_header_end_row.get():
                    report["error"] = "Start Write Row must be after the destination header."
                    return report
                end_limit = end_row if end_row > 0 else ws.max_row
                if end_row > 0 and start_row > end_row:
                    report["error"] = "Start Write Row cannot be after End Write Row."
                    return report
                report.update({'start_row': start_row, 'end_row': end_row or "Unlimited", 'total_zone_rows': (end_limit - start_row + 1) if end_row > 0 else "Unlimited"})
                skipped_rows_set = self._parse_skip_rows(self.dest_skip_rows.get())
                user_skipped, protected_skipped = 0, 0
                for r in range(start_row, end_limit + 1):
                    if r in skipped_rows_set: user_skipped += 1; continue
                    is_auto_skipped = False
                    if self.respect_cell_protection.get() and ws.protection.sheet:
                        if any(ws.cell(r, c).protection.locked for c in self.dest_columns.values()): is_auto_skipped = True
                    if not is_auto_skipped and self.respect_formulas.get():
                        if any(ws.cell(r, c).data_type == 'f' for c in self.dest_columns.values()): is_auto_skipped = True
                    if is_auto_skipped: protected_skipped += 1
                report.update({'user_skipped_count': user_skipped, 'protected_skipped_count': protected_skipped})
                if end_row > 0:
                    report['available_slots'] = max(0, report['total_zone_rows'] - user_skipped - protected_skipped)
                else:
                    report['available_slots'] = "Unlimited"
            return report
        finally:
            FileHandleManager.force_release_handles()

    def preview_transfer(self):
        if not all([self.source_file.get(), os.path.exists(self.source_file.get()), self.dest_file.get(), os.path.exists(self.dest_file.get())]):
            show_custom_error(self.root, self, "Error", "Please select valid source and destination files.")
            return
        if not hasattr(self, 'mapping_combos') or not self.source_columns:
            show_custom_warning(self.root, self, "Warning", "Please load columns first.")
            return
        self.update_status("Generating simulation report...")
        try:
            report_data = self._run_preview_simulation()
            if "error" in report_data:
                PreviewDialog(self.root, self, report_data, [], {}); return
            mappings = {s: c.get() for s, c in self.mapping_combos.items() if c.get()}
            if not mappings:
                show_custom_warning(self.root, self, "Warning", "Please configure at least one mapping for a meaningful preview.")
                return
            with ExcelParser(self.source_file.get()) as p:
                preview_data = p.read_data_preview(self.source_columns, self.source_header_end_row.get(), 10)
            report_data['settings'] = self.get_current_settings()
            PreviewDialog(self.root, self, report_data, preview_data, mappings)
            self.update_status("Preview report generated.")
        except Exception as e:
            self.log_error(f"Error generating preview: {str(e)}\n{traceback.format_exc()}")
            show_custom_error(self.root, self, "Error", f"Failed to generate preview: {str(e)}")
            self.update_status("Preview failed")

    def get_current_settings(self) -> dict:
        return {
            "Source File": os.path.basename(self.source_file.get()), "Destination File": os.path.basename(self.dest_file.get()),
            "Sort Column": self.sort_column.get() or "None", "---": "---",
            "Start Write Row": self.dest_write_start_row.get(), "End Write Row": self.dest_write_end_row.get() or "Unlimited",
            "Skip Rows": self.dest_skip_rows.get() or "None", "Respect Protection": "Yes" if self.respect_cell_protection.get() else "No",
            "Respect Formulas": "Yes" if self.respect_formulas.get() else "No",
        }
    
    def run(self):
        try:
            self.root.mainloop()
        except Exception as e:
            self.log_error(f"Critical error in main loop: {str(e)}")
            # Fallback to console message if GUI fails early
            try:
                root = tk.Tk()
                root.withdraw()
                show_custom_error(None, None, "Critical Error", f"Application encountered a critical error:\n\n{e}")
            except tk.TclError:
                print(f"CRITICAL ERROR: {e}")
                input("Press Enter to exit...")

if __name__ == "__main__":
    try:
        app = ExcelDataMapper()
        app.run()
    except Exception as e:
        logging.critical(f"Failed to start application: {str(e)}\n{traceback.format_exc()}")
        # Fallback to console message if GUI fails early
        try:
            root = tk.Tk()
            root.withdraw()
            # Use a basic messagebox as a last resort if the custom one fails
            from tkinter import messagebox
            messagebox.showerror("Critical Error", f"Application encountered a critical error:\n\n{e}")
        except tk.TclError:
            print(f"CRITICAL ERROR: {e}")
            input("Press Enter to exit...")