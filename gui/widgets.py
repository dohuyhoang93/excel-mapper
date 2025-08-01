"""
Reusable GUI widgets for the Excel Data Mapper application
"""
import tkinter as tk
from tkinter import ttk, messagebox
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
        
        def _bind_to_mousewheel(event):
            self.canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        def _unbind_from_mousewheel(event):
            self.canvas.unbind_all("<MouseWheel>")
        
        self.canvas.bind('<Enter>', _bind_to_mousewheel)
        self.canvas.bind('<Leave>', _unbind_from_mousewheel)

# --- DIALOGS ---

class BaseDialog(tk.Toplevel):
    """A base class for consistent dialog windows."""
    def __init__(self, parent, app_instance, title=""):
        super().__init__(parent)
        self.parent = parent
        self.app = app_instance
        self.title(title)
        
        # Hide the window until it's fully configured and centered
        self.withdraw()

        if hasattr(self.app, 'icon_path') and self.app.icon_path:
            try:
                self.iconbitmap(self.app.icon_path)
            except tk.TclError:
                pass # May fail on some systems/configurations
                
        self.transient(parent)
        self.grab_set()
        self.bind("<Escape>", lambda e: self.destroy())

    def center_on_parent(self):
        # Force update of widget sizes
        self.update_idletasks()
        
        parent_x = self.parent.winfo_x()
        parent_y = self.parent.winfo_y()
        parent_w = self.parent.winfo_width()
        parent_h = self.parent.winfo_height()
        dialog_w = self.winfo_width()
        dialog_h = self.winfo_height()
        
        x = parent_x + (parent_w // 2) - (dialog_w // 2)
        y = parent_y + (parent_h // 2) - (dialog_h // 2)
        
        self.geometry(f"+{x}+{y}")
        
        # Show the window now that it's ready
        self.deiconify()

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
        
        # This must be called at the end of the child __init__
        self.center_on_parent()

class PreviewDialog(BaseDialog):
    """A comprehensive, multi-tab simulation report dialog."""
    def __init__(self, parent, app_instance, report_data: dict, preview_data: list, mappings: dict):
        super().__init__(parent, app_instance, title="Transfer Simulation Report")
        self.geometry("950x650")
        main_frame = ttk_boot.Frame(self, padding=10)
        main_frame.pack(fill=BOTH, expand=True)
        if "error" in report_data:
            self._create_error_view(main_frame, report_data["error"])
            self.center_on_parent() # Center even on error
            return
            
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=BOTH, expand=True, pady=5)
        self._create_summary_tab(self.notebook, report_data)
        self._create_mappings_tab(self.notebook, mappings)
        self._create_data_preview_tab(self.notebook, preview_data, mappings)
        ttk_boot.Button(main_frame, text="Close", command=self.destroy, bootstyle="outline-secondary").pack(pady=(10, 0))
        
        # This must be called at the end of the child __init__
        self.center_on_parent()

    def _adjust_column_widths(self, tree, anchors):
        self.update_idletasks()
        for idx, col in enumerate(tree["columns"]):
            header_width = tkFont.Font().measure(tree.heading(col)["text"])
            max_width = header_width
            for item in tree.get_children():
                try:
                    cell_text = tree.item(item)["values"][idx]
                    cell_width = tkFont.Font().measure(str(cell_text))
                    if cell_width > max_width: max_width = cell_width
                except (IndexError, KeyError):
                    continue
            tree.column(col, width=max_width + 25, anchor=anchors[idx])

    def _create_error_view(self, parent, error_message):
        ttk_boot.Label(parent, text=f"❌ CRITICAL ERROR\n\n{error_message}", bootstyle=DANGER, font="-size 12 -weight bold", justify=CENTER).pack(pady=20, padx=10, fill=BOTH, expand=True)
        ttk_boot.Button(parent, text="Close", command=self.destroy, bootstyle="outline-danger").pack(pady=10)

    def _create_summary_tab(self, notebook, data):
        summary_frame = ttk_boot.Frame(notebook, padding=10)
        notebook.add(summary_frame, text="📊 Summary")
        verdict_frame = ttk_boot.LabelFrame(summary_frame, text="Verdict", padding=10)
        verdict_frame.pack(padx=10, pady=5, fill=X)
        source_count = data.get('source_row_count', 0)
        available_slots = data.get('available_slots', 0)
        if available_slots == "Unlimited" or source_count <= available_slots:
            verdict_text, details, style = "✅ PERFECT", f"All {source_count} source rows will be transferred.", SUCCESS
        else:
            verdict_text, details, style = "⚠️ WARNING", f"Only {available_slots} of {source_count} source rows will be transferred.", WARNING
        ttk_boot.Label(verdict_frame, text=verdict_text, font="-size 14 -weight bold", bootstyle=style).pack()
        ttk_boot.Label(verdict_frame, text=details, wraplength=800).pack(pady=(5,0))
        
        bottom_frame = ttk_boot.Frame(summary_frame)
        bottom_frame.pack(fill=BOTH, expand=True, pady=5)
        
        analysis_frame = ttk_boot.LabelFrame(bottom_frame, text="Data Flow Analysis", padding=10)
        analysis_frame.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 5))
        analysis_frame.columnconfigure(1, weight=1)
        analysis_data = [
            ("Source rows to transfer:", data.get('source_row_count', 'N/A')),
            ("Destination write zone:", f"Row {data.get('start_row', '?')} to {data.get('end_row', '?')}"),
            ("Total rows in zone:", data.get('total_zone_rows', 'N/A')),
            ("User-skipped rows:", data.get('user_skipped_count', 'N/A')),
            ("Protected/Formula rows skipped:", data.get('protected_skipped_count', 'N/A')),
            ("Available rows for writing:", data.get('available_slots', 'N/A'))
        ]
        for i, (label, value) in enumerate(analysis_data):
            ttk_boot.Label(analysis_frame, text=label, anchor=W).grid(row=i, column=0, sticky=EW, pady=2, padx=5)
            ttk_boot.Label(analysis_frame, text=str(value), anchor=W, bootstyle="info").grid(row=i, column=1, sticky=EW, pady=2, padx=5)

        settings_frame = ttk_boot.LabelFrame(bottom_frame, text="Settings Used", padding=10)
        settings_frame.pack(side=LEFT, fill=BOTH, expand=True, padx=(5, 0))
        settings_frame.columnconfigure(1, weight=1)
        settings_data = data.get('settings', {})
        for i, (key, value) in enumerate(settings_data.items()):
            ttk_boot.Label(settings_frame, text=f"{key}:", anchor=W).grid(row=i, column=0, sticky=EW, pady=2, padx=5)
            style = "info" if str(value).lower() not in ["no", "none", ""] else "secondary"
            ttk_boot.Label(settings_frame, text=str(value), anchor=W, bootstyle=style).grid(row=i, column=1, sticky=EW, pady=2, padx=5)

    def _create_mappings_tab(self, notebook, mappings):
        tab_frame = ttk_boot.Frame(notebook, padding=10)
        notebook.add(tab_frame, text="🔗 Column Mappings")
        container = ttk_boot.LabelFrame(tab_frame, text=f"Active Mappings ({len(mappings)})")
        container.pack(fill=BOTH, expand=True)
        cols, anchors = ("Source Column", "Destination Column"), [W, W]
        tree = ttk.Treeview(container, columns=cols, show='headings', bootstyle="info", height=15)
        for i, col in enumerate(cols):
            tree.heading(col, text=col, anchor=anchors[i])
        for source, dest in sorted(mappings.items()):
            tree.insert("", "end", values=(source, dest))
        tree.pack(side=LEFT, fill=BOTH, expand=True, padx=5, pady=5)
        vsb = ttk.Scrollbar(container, orient="vertical", command=tree.yview, bootstyle="info-round")
        vsb.pack(side=RIGHT, fill='y')
        tree.configure(yscrollcommand=vsb.set)
        self.after(100, lambda: self._adjust_column_widths(tree, anchors))

    def _create_data_preview_tab(self, notebook, preview_data, mappings):
        tab_frame = ttk_boot.Frame(notebook, padding=10)
        notebook.add(tab_frame, text="📄 Data Preview")
        container = ttk_boot.LabelFrame(tab_frame, text="Preview of First 10 Rows to be Transferred")
        container.pack(fill=BOTH, expand=True)
        if not preview_data:
            ttk_boot.Label(container, text="No source data found to preview.", bootstyle=INFO).pack(padx=10, pady=10)
            return
        dest_cols = list(mappings.values())
        tree = ttk.Treeview(container, columns=dest_cols, show='headings', bootstyle="info", height=10)
        
        # Set alignment for both header and column content to center
        for col in dest_cols:
            tree.column(col, anchor='center')
            tree.heading(col, text=col, anchor='center')

        dest_to_source = {v: k for k, v in mappings.items()}
        for row_data in preview_data:
            values = [str(row_data.get(dest_to_source.get(dc), ""))[:100] for dc in dest_cols]
            tree.insert("", "end", values=values)
            
        vsb = ttk.Scrollbar(container, orient="vertical", command=tree.yview, bootstyle="info-round")
        hsb = ttk.Scrollbar(container, orient="horizontal", command=tree.xview, bootstyle="info-round")
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
        tree.pack(side=LEFT, fill='both', expand=True, padx=5, pady=5)
        
        # Adjust widths after a short delay
        self.after(100, lambda: self._adjust_column_widths(tree, ['center'] * len(dest_cols)))

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
        
        # Make the dialog blocking
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