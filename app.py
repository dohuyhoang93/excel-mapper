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
from logic.config_manager import ConfigurationManager
from logic.mapper import ColumnMapper
from logic.transfer import ExcelTransferEngine, parse_skip_rows_string
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
        self.root.geometry("1000x700")

        # Center the window on the screen
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
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
        
        self.config_manager = ConfigurationManager()
        self.column_mapper = ColumnMapper()

        self.setup_menu()
        self.setup_gui()
        self.load_app_settings()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
    def setup_menu(self):
        menubar_frame = ttk_boot.Frame(self.root)
        menubar_frame.pack(fill=X, side=TOP, padx=5, pady=(1, 0))

        def create_menu_button(text, menu):
            button = ttk_boot.Button(menubar_frame, text=text, bootstyle="link")
            button.pack(side=LEFT, padx=10, pady=1)
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
        main_frame = ttk_boot.Frame(self.root, padding=5)
        main_frame.pack(fill=BOTH, expand=True, side=TOP)
        
        file_frame = ttk_boot.LabelFrame(main_frame, text="File Selection", padding=5)
        file_frame.pack(fill=X, pady=(0, 1))
        file_frame.columnconfigure(1, weight=1)
        ttk_boot.Label(file_frame, text="Source File:").grid(row=0, column=0, sticky=W, pady=2)
        ttk_boot.Entry(file_frame, textvariable=self.source_file, width=50).grid(row=0, column=1, padx=5, pady=2, sticky=EW)
        ttk_boot.Button(file_frame, bootstyle="outline", text="Browse", command=self.browse_source_file).grid(row=0, column=2, padx=5, pady=2)
        ttk_boot.Label(file_frame, text="Destination File:").grid(row=1, column=0, sticky=W, pady=2)
        ttk_boot.Entry(file_frame, textvariable=self.dest_file, width=50).grid(row=1, column=1, padx=5, pady=2, sticky=EW)
        ttk_boot.Button(file_frame, bootstyle="outline", text="Browse", command=self.browse_dest_file).grid(row=1, column=2, padx=5, pady=2)
        
        header_frame = ttk_boot.LabelFrame(main_frame, text="Header Configuration", padding=5)
        header_frame.pack(fill=X, pady=(0, 2))
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
        
        write_zone_frame = ttk_boot.LabelFrame(main_frame, text="Setting write zone", padding=5)
        write_zone_frame.pack(fill=X, pady=(0, 2))
        write_zone_frame.columnconfigure(1, weight=1)
        write_zone_frame.columnconfigure(3, weight=1)
        write_zone_frame.columnconfigure(4, weight=2)
        ttk_boot.Label(write_zone_frame, text="Start Write Row:").grid(row=0, column=0, sticky=W, padx=5, pady=2)
        ttk_boot.Spinbox(write_zone_frame, from_=1, to=99999, textvariable=self.dest_write_start_row, width=8).grid(row=0, column=1, sticky=W, padx=5)
        ttk_boot.Label(write_zone_frame, text="End Write Row (0 = unlimited):").grid(row=0, column=2, sticky=W, padx=(20, 5), pady=2)
        ttk_boot.Spinbox(write_zone_frame, from_=0, to=99999, textvariable=self.dest_write_end_row, width=8).grid(row=0, column=3, sticky=W, padx=5)
        self.detect_button = ttk_boot.Button(write_zone_frame, text="Detect Zone", command=self.detect_write_zone, bootstyle="outline-info")
        self.detect_button.grid(row=0, column=4, padx=(10, 5), pady=2, sticky=E)
        ttk_boot.Label(write_zone_frame, text="Skip Rows (e.g., 15, 20-25):").grid(row=1, column=0, sticky=W, padx=5, pady=2)
        ttk_boot.Entry(write_zone_frame, textvariable=self.dest_skip_rows).grid(row=1, column=1, columnspan=4, padx=5, sticky=EW)
        ttk_boot.Checkbutton(write_zone_frame, text="Respect cell protection", variable=self.respect_cell_protection).grid(row=2, column=0, columnspan=2, sticky=W, padx=5, pady=2)
        ttk_boot.Checkbutton(write_zone_frame, text="Respect formulas", variable=self.respect_formulas).grid(row=2, column=2, columnspan=3, sticky=W, padx=(20, 5), pady=2)

        mapping_container = ttk_boot.LabelFrame(main_frame, text="Column Mapping", padding=3)
        mapping_container.pack(fill=BOTH, expand=True, pady=(0, 2))
        self.mapping_scroll_frame = ScrollableFrame(mapping_container)
        self.mapping_scroll_frame.pack(fill=BOTH, expand=True)
        
        sort_frame = ttk_boot.LabelFrame(main_frame, text="Sort Configuration", padding=5)
        sort_frame.pack(fill=X, pady=(0, 2))
        ttk_boot.Label(sort_frame, text="Sort by Column (optional):").grid(row=0, column=0, sticky=W, pady=(0,2))
        self.sort_combo = ttk_boot.Combobox(sort_frame, textvariable=self.sort_column, width=50)
        self.sort_combo.grid(row=0, column=1, padx=5, sticky=W)
        
        action_frame = ttk_boot.Frame(main_frame)
        action_frame.pack(fill=X, pady=(0, 2))
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
        self.status_frame.pack(fill=X, pady=(2, 0))
        self.status_frame.columnconfigure(0, weight=1)  # Column for progress bar will expand
        self.status_frame.columnconfigure(1, weight=0)  # Column for label will not expand

        self.progress = ttk_boot.Progressbar(self.status_frame, mode='determinate', bootstyle=SUCCESS)
        self.progress.grid(row=0, column=0, sticky='we')
        
        self.status_label = ttk_boot.Label(self.status_frame, text="Ready")
        self.status_label.grid(row=0, column=1, sticky='we')
    
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
        ttk_boot.Label(self.mapping_scroll_frame.scrollable_frame, text="Source Column", font="-weight bold").grid(row=0, column=0, sticky=W, padx=5, pady=(2, 2))
        ttk_boot.Label(self.mapping_scroll_frame.scrollable_frame, text="Destination Column", font="-weight bold").grid(row=0, column=2, sticky=W, padx=5, pady=(2, 2))
        
        self.mapping_combos = {}
        for i, source_col_name in enumerate(self.source_columns.keys(), start=1):
            ttk_boot.Label(self.mapping_scroll_frame.scrollable_frame, text=source_col_name, anchor=W).grid(row=i, column=0, sticky=EW, padx=5, pady=2)
            ttk_boot.Label(self.mapping_scroll_frame.scrollable_frame, text="→").grid(row=i, column=1, sticky=W, padx=5)
            dest_combo = ttk_boot.Combobox(self.mapping_scroll_frame.scrollable_frame, values=[""] + list(self.dest_columns.keys()), width=60)
            dest_combo.grid(row=i, column=2, sticky=EW, padx=5, pady=2)
            if apply_suggestions:
                suggested = self.column_mapper.suggest_mapping(source_col_name, list(self.dest_columns.keys()))
                if suggested: dest_combo.set(suggested)
            self.mapping_combos[source_col_name] = dest_combo
    
    def save_config(self):
        try:
            if not hasattr(self, 'mapping_combos') or not self.mapping_combos:
                show_custom_warning(self.root, self, "Warning", "No mappings to save. Please load columns first.")
                return
            
            config_file_path = filedialog.asksaveasfilename(
                title="Save Job Configuration As",
                defaultextension=".json",
                filetypes=[("JSON files", "*.json")],
                initialdir=self.config_manager.config_dir
            )
            if not config_file_path: return

            mappings = {source_col: combo.get() for source_col, combo in self.mapping_combos.items() if combo.get()}
            
            job_config = {
                #"source_file": self.source_file.get(), "dest_file": self.dest_file.get(),
                "source_header_start_row": self.source_header_start_row.get(), "source_header_end_row": self.source_header_end_row.get(),
                "dest_header_start_row": self.dest_header_start_row.get(), "dest_header_end_row": self.dest_header_end_row.get(),
                "dest_write_start_row": self.dest_write_start_row.get(), "dest_write_end_row": self.dest_write_end_row.get(),
                "dest_skip_rows": self.dest_skip_rows.get(), 
                "respect_cell_protection": self.respect_cell_protection.get(),
                "respect_formulas": self.respect_formulas.get(), 
                "sort_column": self.sort_column.get(), 
                "mapping": mappings,
            }
            
            self.config_manager.save_job_config(job_config, config_file_path)
            self.save_app_settings() # Update last used files in app settings
            
            self.update_status(f"Job configuration saved to {os.path.basename(config_file_path)}")
            self.log_info(f"Job configuration saved: {config_file_path}")
        except Exception as e:
            self.log_error(f"Error saving job configuration: {str(e)}")
            show_custom_error(self.root, self, "Error", f"Failed to save job configuration: {str(e)}")
    
    def load_config(self):
        try:
            config_file_path = filedialog.askopenfilename(
                title="Load Job Configuration", 
                filetypes=[("JSON files", "*.json")],
                initialdir=self.config_manager.config_dir
            )
            if not config_file_path: return

            config = self.config_manager.load_job_config(config_file_path)

            #self.source_file.set(config.get("source_file", ""))
            #self.dest_file.set(config.get("dest_file", ""))
            self.source_header_start_row.set(config.get("source_header_start_row", 1))
            self.source_header_end_row.set(config.get("source_header_end_row", 1))
            self.dest_header_start_row.set(config.get("dest_header_start_row", 9))
            self.dest_header_end_row.set(config.get("dest_header_end_row", 9))
            self.dest_write_start_row.set(config.get("dest_write_start_row", self.dest_header_end_row.get() + 1))
            self.dest_write_end_row.set(config.get("dest_write_end_row", 0))
            self.dest_skip_rows.set(config.get("dest_skip_rows", ""))
            self.respect_cell_protection.set(config.get("respect_cell_protection", True))
            self.respect_formulas.set(config.get("respect_formulas", True))
            
            if self.source_file.get() and self.dest_file.get():
                saved_sort_col = config.get("sort_column", "")
                self.safe_load_columns(saved_sort_col=saved_sort_col, apply_suggestions=False)
                mappings = config.get("mapping", {})
                for source_col, dest_col in mappings.items():
                    if source_col in self.mapping_combos:
                        self.mapping_combos[source_col].set(dest_col)

            self.save_app_settings() # Update last used files
            self.update_status(f"Job configuration loaded from {os.path.basename(config_file_path)}")
            self.log_info(f"Job configuration loaded: {config_file_path}")
        except Exception as e:
            self.log_error(f"Error in load_config: {str(e)}")
            show_custom_error(self.root, self, "Error", f"Failed to load job configuration: {str(e)}")
    
    def load_app_settings(self):
        """Loads global application settings at startup."""
        try:
            settings = self.config_manager.load_app_settings()
            
            # Restore last used theme
            new_theme = settings.get("theme", "flatly")
            if new_theme != self.current_theme:
                self.root.style.theme_use(new_theme)
                self.current_theme = new_theme
            
            # Restore last used keywords
            self.detection_keywords.set(settings.get("detection_keywords", "total,sum,cộng,tổng,thành tiền"))

            # Restore last used file paths if they exist
            last_source = settings.get("last_source_file", "")
            if os.path.exists(last_source):
                self.source_file.set(last_source)
            
            last_dest = settings.get("last_dest_file", "")
            if os.path.exists(last_dest):
                self.dest_file.set(last_dest)

            self.log_info("Application settings loaded.")

        except Exception as e:
            self.log_error(f"Error loading application settings: {str(e)}")
            # Don't show a popup for this, just log it.
            
    def save_app_settings(self):
        """Saves global application settings."""
        try:
            settings = {
                "theme": self.current_theme,
                "detection_keywords": self.detection_keywords.get(),
                "last_source_file": self.source_file.get(),
                "last_dest_file": self.dest_file.get()
            }
            self.config_manager.save_app_settings(settings)
            self.log_info("Application settings saved.")
        except Exception as e:
            self.log_error(f"Error saving application settings: {str(e)}")

    def on_closing(self):
        """Handles window closing event."""
        self.save_app_settings()
        self.root.destroy()
    
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
            # 1. Collect all settings for the engine
            settings = {
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
                "sort_column": self.sort_column.get(),
                "mappings": mappings,
                "source_columns": self.source_columns,
                "dest_columns": self.dest_columns,
            }

            # 2. Create and run the engine
            engine = ExcelTransferEngine(settings, self.update_progress_callback)
            engine.run_transfer()
            
            # 3. Update UI on success
            self.root.after(0, self.on_transfer_success)
        except Exception as e:
            # 4. Update UI on error
            self.root.after(0, self.on_transfer_error, e)

    def update_progress_callback(self, value: int, message: str):
        """Callback function for the engine to update the GUI's progress."""
        self.progress['value'] = value
        self.update_status(message)
        self.root.update() # Use update, not update_idletasks, to force redraw
    
    def on_transfer_success(self):
        self.progress['value'] = 100
        self.update_status("Transfer completed successfully")
        self.enable_controls()
        show_custom_info(self.root, self, "Success", "Data transfer completed successfully!")
        if show_custom_question(self.root, self, "Open Folder", "Would you like to open the destination folder?"):
            self.open_dest_folder()

    def on_transfer_error(self, error):
        self.log_error(f"Error during transfer thread: {str(error)}\n{traceback.format_exc()}")
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
                
                skipped_rows_set = parse_skip_rows_string(self.dest_skip_rows.get())
                mappings = {s: c.get() for s, c in self.mapping_combos.items() if c.get()}
                mapped_dest_indices = {self.dest_columns[name] for name in mappings.values() if name in self.dest_columns}

                user_skipped, protected_skipped = 0, 0
                for r in range(start_row, end_limit + 1):
                    if r in skipped_rows_set:
                        user_skipped += 1
                        continue
                    
                    # Check for protected cells or formulas in the row
                    is_row_auto_skipped = False
                    if self.respect_cell_protection.get() and ws.protection.sheet:
                        # A row is considered skipped if ANY of its destination cells are locked
                        if any(ws.cell(r, c_idx).protection.locked for c_idx in mapped_dest_indices):
                            is_row_auto_skipped = True
                    
                    # A row is also considered skipped if ALL of its mapped destination cells contain formulas
                    if self.respect_formulas.get():
                        if mapped_dest_indices and all(ws.cell(r, c_idx).data_type == 'f' for c_idx in mapped_dest_indices):
                             is_row_auto_skipped = True

                    if is_row_auto_skipped:
                        protected_skipped += 1

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