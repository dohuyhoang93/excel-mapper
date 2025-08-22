"Reusable GUI widgets for the Excel Data Mapper application"
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import ttkbootstrap as ttk_boot
from ttkbootstrap.constants import *
from typing import List, Dict, Callable, Optional, Any
import tkinter.font as tkFont

class ScrollableFrame(ttk_boot.Frame):
    """A scrollable frame widget"""
    
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        
        self.canvas = tk.Canvas(self, highlightthickness=0)
        self.scrollbar = ttk_boot.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk_boot.Frame(self.canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        self._bind_mousewheel()
    
    def _bind_mousewheel(self):
        def _on_mousewheel(event):
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        # Bind directly to the canvas and scrollable frame, not the entire application
        self.canvas.bind("<MouseWheel>", _on_mousewheel)
        self.scrollable_frame.bind("<MouseWheel>", _on_mousewheel)

# --- DIALOGS ---

class BaseDialog(tk.Toplevel):
    """A base class for consistent dialog windows."""
    def __init__(self, parent, app_instance, title=""):
        super().__init__(parent)
        self.parent = parent
        self.app = app_instance
        self.title(title)
        
        self.withdraw()

        if hasattr(self.app, 'icon_path') and self.app.icon_path:
            try:
                self.iconbitmap(self.app.icon_path)
            except tk.TclError:
                pass
                
        self.transient(parent)
        self.grab_set()
        self.bind("<Escape>", lambda e: self.destroy())

    def center_on_parent(self):
        self.update_idletasks()
        self.deiconify()
        
        parent_x = self.parent.winfo_x()
        parent_y = self.parent.winfo_y()
        parent_w = self.parent.winfo_width()
        parent_h = self.parent.winfo_height()
        dialog_w = self.winfo_width()
        dialog_h = self.winfo_height()
        
        x = parent_x + (parent_w // 2) - (dialog_w // 2)
        y = parent_y + (parent_h // 2) - (dialog_h // 2)
        
        self.geometry(f"+{x}+{y}")
        self.lift()
        self.focus_set()

class AboutDialog(BaseDialog):
    """The 'About' dialog window."""
    def __init__(self, parent, app_instance):
        super().__init__(parent, app_instance, title="About Excel Data Mapper")
        self.geometry("450x380")
        self.resizable(False, False)
        about_text = """Excel Data Mapper v1.2

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
        ttk_boot.Label(self, text=about_text, justify=LEFT, padding=(20, 20)).pack(expand=True, fill=BOTH)
        ttk_boot.Button(self, text="OK", command=self.destroy, bootstyle=PRIMARY, width=10).pack(pady=15)
        self.center_on_parent()

class PreviewDialog(BaseDialog):
    """A comprehensive, multi-tab simulation report dialog."""
    def __init__(self, parent, app_instance, report_data: dict, preview_data: list, mappings: dict):
        super().__init__(parent, app_instance, title="Transfer Simulation Report")
        self.geometry("950x650")
        self.result = None
        self.excluded_groups = []

        main_frame = ttk_boot.Frame(self, padding=10)
        main_frame.pack(fill=BOTH, expand=True)

        if "error" in report_data:
            self._create_error_view(main_frame, report_data["error"])
            self.center_on_parent()
            return
            
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=BOTH, expand=True, pady=5)
        
        self._create_summary_tab(self.notebook, report_data, mappings)
        self._create_groups_tab(self.notebook, report_data)
        self._create_validation_tab(self.notebook, report_data)
        self._create_data_preview_tab(self.notebook, preview_data, mappings)

        button_frame = ttk_boot.Frame(main_frame)
        button_frame.pack(fill=X, pady=(10,0))
        ttk_boot.Button(button_frame, text="Run Transfer", command=self.on_run_transfer, bootstyle=SUCCESS).pack(side=RIGHT)
        ttk_boot.Button(button_frame, text="Cancel", command=self.destroy, bootstyle="secondary").pack(side=RIGHT, padx=5)
        
        self.center_on_parent()
        self.wait_window()

    def on_run_transfer(self):
        self.result = self.excluded_groups
        self.destroy()

    def _adjust_column_widths(self, tree, cols, anchors):
        self.update_idletasks()
        for idx, col_name in enumerate(cols):
            header_width = tkFont.Font().measure(tree.heading(col_name)["text"])
            max_width = header_width
            for item in tree.get_children():
                try:
                    cell_text = tree.item(item)["values"][idx]
                    cell_width = tkFont.Font().measure(str(cell_text))
                    if cell_width > max_width: max_width = cell_width
                except (IndexError, KeyError):
                    continue
            tree.column(col_name, width=max_width + 25, anchor=anchors[idx])

    def _create_error_view(self, parent, error_message):
        ttk_boot.Label(parent, text=f"❌ CRITICAL ERROR\n\n{error_message}", bootstyle=DANGER, font="-size 12 -weight bold", justify=CENTER).pack(pady=20, padx=10, fill=BOTH, expand=True)
        ttk_boot.Button(parent, text="Close", command=self.destroy, bootstyle="outline-danger").pack(pady=10)

    def _create_summary_tab(self, notebook, data, mappings):
        summary_frame = ttk_boot.Frame(notebook, padding=10)
        notebook.add(summary_frame, text="📊 Summary")

        stats_frame = ttk_boot.LabelFrame(summary_frame, text="Simulation Summary", padding=10)
        stats_frame.pack(padx=10, pady=5, fill=X)
        stats_frame.columnconfigure(1, weight=1)

        row_limit = data.get('row_limit', 0)
        limit_text = f"{row_limit} rows" if row_limit > 0 else "All rows"

        summary_data = [
            ("Preview based on:", limit_text),
            ("Total source rows found:", data.get('total_rows', 'N/A')),
            ("Number of unique groups found:", data.get('group_count', 'N/A')),
            ("Active column mappings:", f"{len(mappings)}"),
            ("New sheets to be created:", data.get('group_count', 'N/A')),
            ("Potential validation errors:", len(data.get('validation_errors', []))),
        ]
        for i, (label, value) in enumerate(summary_data):
            ttk_boot.Label(stats_frame, text=label, anchor=W).grid(row=i, column=0, sticky=EW, pady=2, padx=5)
            ttk_boot.Label(stats_frame, text=str(value), anchor=W, bootstyle="info").grid(row=i, column=1, sticky=EW, pady=2, padx=5)

        settings_frame = ttk_boot.LabelFrame(summary_frame, text="Current Settings", padding=10)
        settings_frame.pack(padx=10, pady=5, fill=X, expand=True)
        settings_frame.columnconfigure(1, weight=1)
        
        settings_data = data.get('settings', {})
        for i, (key, value) in enumerate(settings_data.items()):
            ttk_boot.Label(settings_frame, text=f"{key}:", anchor=W).grid(row=i, column=0, sticky=EW, pady=2, padx=5)
            ttk_boot.Label(settings_frame, text=str(value), anchor=W, bootstyle="info").grid(row=i, column=1, sticky=EW, pady=2, padx=5)

    def _create_groups_tab(self, notebook, data):
        tab_frame = ttk_boot.Frame(notebook, padding=10)
        notebook.add(tab_frame, text="📋 Group Details")
        
        button = ttk_boot.Button(tab_frame, text="Exclude Groups...", command=self._exclude_groups_dialog, bootstyle="outline-danger")
        button.pack(anchor=E, pady=5)

        container = ttk_boot.LabelFrame(tab_frame, text=f"All Groups ({data.get('group_count', 0)})")
        container.pack(fill=BOTH, expand=True, padx=5, pady=5)

        cols = ("Group Name", "Row Count")
        self.groups_tree = ttk.Treeview(container, columns=cols, show='headings', bootstyle="info")
        self.groups_tree.column("Group Name", anchor=W, width=400)
        self.groups_tree.column("Row Count", anchor=E, width=100)
        self.groups_tree.heading("Group Name", text="Group Name", anchor=W)
        self.groups_tree.heading("Row Count", text="Row Count", anchor=E)

        groups_data = data.get('groups', [])
        for group_name, row_count in groups_data:
            self.groups_tree.insert("", "end", values=(group_name, row_count))

        self.groups_tree.pack(side=LEFT, fill=BOTH, expand=True, padx=5, pady=5)
        vsb = ttk.Scrollbar(container, orient="vertical", command=self.groups_tree.yview, bootstyle="info-round")
        vsb.pack(side=RIGHT, fill='y')
        self.groups_tree.configure(yscrollcommand=vsb.set)
        
        self._adjust_column_widths(self.groups_tree, cols, [W, E])

    def _exclude_groups_dialog(self):
        current_exclusions = ", ".join(self.excluded_groups)
        to_exclude = simpledialog.askstring("Exclude Groups", 
                                            "Enter group names to exclude, separated by commas:",
                                            initialvalue=current_exclusions,
                                            parent=self)
        if to_exclude is not None:
            self.excluded_groups = [name.strip() for name in to_exclude.split(",") if name.strip()]
            show_custom_info(self, self.app, "Info", f"{len(self.excluded_groups)} groups marked for exclusion.")

    def _create_validation_tab(self, notebook, data):
        tab_frame = ttk_boot.Frame(notebook, padding=10)
        validation_errors = data.get('validation_errors', [])
        tab_text = f"⚠️ Validation ({len(validation_errors)})"
        notebook.add(tab_frame, text=tab_text)

        if not validation_errors:
            ttk_boot.Label(tab_frame, text="✅ No data validation errors detected.", bootstyle="success", font="-size 12").pack(pady=20)
            return

        container = ttk_boot.LabelFrame(tab_frame, text="Potential Data Validation Errors")
        container.pack(fill=BOTH, expand=True, padx=5, pady=5)

        cols = ("Source Row", "Dest Column", "Invalid Value", "Rule")
        self.validation_tree = ttk.Treeview(container, columns=cols, show='headings', bootstyle="danger")
        
        for col in cols:
            self.validation_tree.heading(col, text=col)
        
        for error in validation_errors:
            self.validation_tree.insert("", "end", values=(error['row'], error['column'], error['value'], error['rule']))

        self.validation_tree.pack(side=LEFT, fill=BOTH, expand=True, padx=5, pady=5)
        vsb = ttk.Scrollbar(container, orient="vertical", command=self.validation_tree.yview, bootstyle="danger-round")
        vsb.pack(side=RIGHT, fill='y')
        self.validation_tree.configure(yscrollcommand=vsb.set)
        
        self._adjust_column_widths(self.validation_tree, cols, [E, W, W, W])

    def _create_data_preview_tab(self, notebook, preview_data, mappings):
        tab_frame = ttk_boot.Frame(notebook, padding=10)
        notebook.add(tab_frame, text="📄 Data Preview")
        container = ttk_boot.LabelFrame(tab_frame, text="Preview of First 10 Rows to be Transferred")
        container.pack(fill=BOTH, expand=True)
        if not preview_data:
            ttk_boot.Label(container, text="No source data found to preview.", bootstyle=INFO).pack(padx=10, pady=10)
            return
        
        dest_cols = [v for v in mappings.values() if v]
        if not dest_cols:
            ttk_boot.Label(container, text="No columns are mapped for preview.", bootstyle=INFO).pack(padx=10, pady=10)
            return

        self.data_tree = ttk.Treeview(container, columns=dest_cols, show='headings', bootstyle="info", height=10)
        
        for col in dest_cols:
            self.data_tree.column(col, anchor='center')
            self.data_tree.heading(col, text=col, anchor='center')

        dest_to_source = {v: k for k, v in mappings.items()}
        for row_data in preview_data:
            values = [str(row_data.get(dest_to_source.get(dc, ""), ""))[:100] for dc in dest_cols]
            self.data_tree.insert("", "end", values=values)
            
        vsb = ttk.Scrollbar(container, orient="vertical", command=self.data_tree.yview, bootstyle="info-round")
        hsb = ttk.Scrollbar(container, orient="horizontal", command=self.data_tree.xview, bootstyle="info-round")
        self.data_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
        self.data_tree.pack(side=LEFT, fill='both', expand=True, padx=5, pady=5)
        
        self._adjust_column_widths(self.data_tree, dest_cols, ['center'] * len(dest_cols))

class DetectionConfigDialog(BaseDialog):
    """Dialog to configure end-row detection keywords."""
    def __init__(self, parent, app_instance):
        super().__init__(parent, app_instance, title="Configure Detection Keywords")
        self.geometry("500x150")
        self.resizable(False, False)
        self.parent_app = app_instance
        self.temp_keywords = tk.StringVar(value=self.parent_app.detection_keywords.get())
        main_frame = ttk_boot.Frame(self, padding=15)
        main_frame.pack(fill=BOTH, expand=True)
        ttk_boot.Label(main_frame, text="Enter keywords to detect the 'total' row, separated by commas:").pack(anchor=W)
        ttk_boot.Entry(main_frame, textvariable=self.temp_keywords).pack(fill=X, pady=5)
        button_frame = ttk_boot.Frame(main_frame)
        button_frame.pack(fill=X, pady=10)
        ttk_boot.Button(button_frame, text="Save", command=self.save, bootstyle=SUCCESS).pack(side=RIGHT, padx=5)
        ttk_boot.Button(button_frame, text="Cancel", command=self.destroy, bootstyle="secondary").pack(side=RIGHT)
        self.center_on_parent()

    def save(self):
        self.parent_app.detection_keywords.set(self.temp_keywords.get())
        self.parent_app.log_info(f"Detection keywords updated to: {self.temp_keywords.get()}")
        self.destroy()

class CustomMessageDialog(BaseDialog):
    """A custom messagebox replacement that respects app theme and icon."""
    def __init__(self, parent, app_instance, title, message, dialog_type="info"):
        super().__init__(parent, app_instance, title)
        self.result = None
        self.dialog_type = dialog_type
        self.resizable(False, False)

        main_frame = ttk_boot.Frame(self, padding=(20, 20, 20, 10))
        main_frame.pack(fill=BOTH, expand=True)

        icon_label = ttk_boot.Label(main_frame, font="-size 28")
        icon_label.pack(side=LEFT, padx=(0, 15), anchor=N)

        text_label = ttk_boot.Label(main_frame, text=message, wraplength=350, justify=LEFT)
        text_label.pack(side=LEFT, fill=BOTH, expand=True, pady=4)

        button_frame = ttk_boot.Frame(self, padding=(10, 0, 10, 10))
        button_frame.pack(fill=X)

        if self.dialog_type == "info":
            icon_label.config(text="ⓘ", bootstyle=INFO)
            ok_button = ttk_boot.Button(button_frame, text="OK", command=self.on_ok, bootstyle="info")
            ok_button.pack(side=RIGHT)
            ok_button.focus_set()
            self.bind("<Return>", lambda e: self.on_ok())
        elif self.dialog_type == "error":
            icon_label.config(text="❌", bootstyle=DANGER)
            ok_button = ttk_boot.Button(button_frame, text="OK", command=self.on_ok, bootstyle="danger")
            ok_button.pack(side=RIGHT)
            ok_button.focus_set()
            self.bind("<Return>", lambda e: self.on_ok())
        elif self.dialog_type == "warning":
            icon_label.config(text="⚠️", bootstyle=WARNING)
            ok_button = ttk_boot.Button(button_frame, text="OK", command=self.on_ok, bootstyle="warning")
            ok_button.pack(side=RIGHT)
            ok_button.focus_set()
            self.bind("<Return>", lambda e: self.on_ok())
        elif self.dialog_type == "question":
            icon_label.config(text="?", bootstyle=INFO)
            no_button = ttk_boot.Button(button_frame, text="No", command=self.on_no, bootstyle="secondary")
            no_button.pack(side=RIGHT, padx=(0, 5))
            yes_button = ttk_boot.Button(button_frame, text="Yes", command=self.on_yes, bootstyle="success")
            yes_button.pack(side=RIGHT, padx=5)
            yes_button.focus_set()
            self.bind("<Return>", lambda e: self.on_yes())

        self.center_on_parent()
        self.wait_window()

    def on_ok(self):
        self.result = True
        self.destroy()

    def on_yes(self):
        self.result = True
        self.destroy()

    def on_no(self):
        self.result = False
        self.destroy()

# --- Utility Functions for Dialogs ---

def show_custom_info(parent, app_instance, title: str, message: str):
    """Shows a themed info dialog."""
    CustomMessageDialog(parent, app_instance, title, message, "info")

def show_custom_error(parent, app_instance, title: str, message: str):
    """Shows a themed error dialog."""
    CustomMessageDialog(parent, app_instance, title, message, "error")

def show_custom_warning(parent, app_instance, title: str, message: str):
    """Shows a themed warning dialog."""
    CustomMessageDialog(parent, app_instance, title, message, "warning")

def show_custom_question(parent, app_instance, title: str, message: str) -> bool:
    """Shows a themed question dialog and returns the boolean result."""
    dialog = CustomMessageDialog(parent, app_instance, title, message, "question")
    return dialog.result
