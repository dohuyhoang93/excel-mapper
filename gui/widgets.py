"""
Reusable GUI widgets for the Excel Data Mapper application
"""
import tkinter as tk
from tkinter import ttk, messagebox
import ttkbootstrap as ttk_boot
from ttkbootstrap.constants import *
from typing import List, Dict, Callable, Optional, Any

class ScrollableFrame(ttk_boot.Frame):
    """A scrollable frame widget"""
    
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        
        # Create canvas and scrollbar
        self.canvas = tk.Canvas(self, highlightthickness=0)
        self.scrollbar = ttk_boot.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk_boot.Frame(self.canvas)
        
        # Configure scrolling
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Pack components
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # Bind mousewheel
        self._bind_mousewheel()
    
    def _bind_mousewheel(self):
        """Bind mousewheel to canvas"""
        def _on_mousewheel(event):
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def _bind_to_mousewheel(event):
            self.canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        def _unbind_from_mousewheel(event):
            self.canvas.unbind_all("<MouseWheel>")
        
        self.canvas.bind('<Enter>', _bind_to_mousewheel)
        self.canvas.bind('<Leave>', _unbind_from_mousewheel)

class MappingWidget(ttk_boot.Frame):
    """Widget for column mapping configuration"""
    
    def __init__(self, parent, source_columns: List[str], dest_columns: List[str], 
                 mapping_changed_callback: Optional[Callable] = None, **kwargs):
        super().__init__(parent, **kwargs)
        
        self.source_columns = source_columns
        self.dest_columns = dest_columns
        self.mapping_changed_callback = mapping_changed_callback
        self.mapping_rows = []
        
        self.setup_ui()
        self.create_mapping_rows()
    
    def setup_ui(self):
        """Setup the basic UI structure"""
        # Header
        header_frame = ttk_boot.Frame(self)
        header_frame.pack(fill=X, pady=(0, 10))
        
        ttk_boot.Label(header_frame, text="Source Column", 
                      font=("Arial", 10, "bold")).pack(side=LEFT, padx=(0, 20))
        ttk_boot.Label(header_frame, text="→", 
                      font=("Arial", 12, "bold")).pack(side=LEFT, padx=10)
        ttk_boot.Label(header_frame, text="Destination Column", 
                      font=("Arial", 10, "bold")).pack(side=LEFT, padx=(20, 0))
        ttk_boot.Label(header_frame, text="Confidence", 
                      font=("Arial", 10, "bold")).pack(side=RIGHT, padx=(20, 0))
        
        # Scrollable content
        self.scroll_frame = ScrollableFrame(self)
        self.scroll_frame.pack(fill=BOTH, expand=True)
    
    def create_mapping_rows(self):
        """Create mapping rows for each source column"""
        from logic.mapper import ColumnMapper
        
        mapper = ColumnMapper()
        
        for source_col in self.source_columns:
            row_frame = ttk_boot.Frame(self.scroll_frame.scrollable_frame)
            row_frame.pack(fill=X, pady=2)
            
            # Source column label
            source_label = ttk_boot.Label(row_frame, text=source_col, width=25, anchor=W)
            source_label.pack(side=LEFT, padx=5)
            
            # Arrow
            ttk_boot.Label(row_frame, text="→").pack(side=LEFT, padx=10)
            
            # Destination column combobox
            dest_var = tk.StringVar()
            dest_combo = ttk_boot.Combobox(row_frame, textvariable=dest_var, 
                                          values=[""] + self.dest_columns, width=35)
            dest_combo.pack(side=LEFT, padx=5)
            
            # Auto-suggest
            suggested = mapper.suggest_mapping(source_col, self.dest_columns)
            if suggested:
                dest_var.set(suggested)
            
            # Confidence indicator
            confidence_var = tk.StringVar()
            confidence_label = ttk_boot.Label(row_frame, textvariable=confidence_var, 
                                             width=10, anchor=E)
            confidence_label.pack(side=RIGHT, padx=5)
            
            # Update confidence when selection changes
            def update_confidence(event, src=source_col, dest_var=dest_var, conf_var=confidence_var):
                dest_col = dest_var.get()
                if dest_col:
                    confidence = mapper.get_mapping_confidence(src, dest_col)
                    conf_var.set(f"{confidence:.1%}")
                    
                    # Color code confidence
                    if confidence >= 0.8:
                        confidence_label.configure(bootstyle=SUCCESS)
                    elif confidence >= 0.5:
                        confidence_label.configure(bootstyle=WARNING)
                    else:
                        confidence_label.configure(bootstyle=DANGER)
                else:
                    conf_var.set("")
                    confidence_label.configure(bootstyle=DEFAULT)
                
                if self.mapping_changed_callback:
                    self.mapping_changed_callback()
            
            dest_combo.bind('<<ComboboxSelected>>', update_confidence)
            
            # Initial confidence update
            if suggested:
                update_confidence(None)
            
            self.mapping_rows.append({
                'source_column': source_col,
                'dest_var': dest_var,
                'combo': dest_combo,
                'confidence_var': confidence_var,
                'frame': row_frame
            })
    
    def get_mappings(self) -> Dict[str, str]:
        """Get current column mappings"""
        mappings = {}
        for row in self.mapping_rows:
            dest_col = row['dest_var'].get()
            if dest_col:
                mappings[row['source_column']] = dest_col
        return mappings
    
    def set_mappings(self, mappings: Dict[str, str]):
        """Set column mappings"""
        for row in self.mapping_rows:
            source_col = row['source_column']
            if source_col in mappings:
                row['dest_var'].set(mappings[source_col])
    
    def validate_mappings(self) -> tuple[bool, List[str]]:
        """Validate current mappings"""
        mappings = self.get_mappings()
        
        errors = []
        dest_values = list(mappings.values())
        
        # Check for duplicates
        duplicates = [col for col in set(dest_values) if dest_values.count(col) > 1]
        if duplicates:
            errors.append(f"Duplicate destination columns: {', '.join(duplicates)}")
        
        # Check for missing mappings (optional - warn only)
        unmapped = [row['source_column'] for row in self.mapping_rows 
                   if not row['dest_var'].get()]
        if unmapped:
            errors.append(f"Unmapped source columns: {', '.join(unmapped)}")
        
        return len(errors) == 0, errors

class ProgressDialog(tk.Toplevel):
    """Progress dialog for long-running operations"""
    
    def __init__(self, parent, title="Processing...", message="Please wait..."):
        super().__init__(parent)
        
        self.parent = parent
        self.cancelled = False
        
        # Configure window
        self.title(title)
        self.geometry("400x150")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        
        # Center on parent
        self.geometry(f"+{parent.winfo_rootx() + 50}+{parent.winfo_rooty() + 50}")
        
        self.setup_ui(message)
        
    def setup_ui(self, message):
        """Setup progress dialog UI"""
        main_frame = ttk_boot.Frame(self, padding=20)
        main_frame.pack(fill=BOTH, expand=True)
        
        # Message
        self.message_label = ttk_boot.Label(main_frame, text=message, 
                                           font=("Arial", 10))
        self.message_label.pack(pady=(0, 10))
        
        # Progress bar
        self.progress = ttk_boot.Progressbar(main_frame, mode='determinate', 
                                            length=300, bootstyle=SUCCESS)
        self.progress.pack(pady=(0, 5))
        
        # Progress text
        self.progress_text = ttk_boot.Label(main_frame, text="0%", 
                                           font=("Arial", 9))
        self.progress_text.pack(pady=(0, 15))
        
        # Cancel button
        self.cancel_button = ttk_boot.Button(main_frame, text="Cancel", 
                                            command=self.cancel, bootstyle=SECONDARY)
        self.cancel_button.pack()
        
        # Handle window close
        self.protocol("WM_DELETE_WINDOW", self.cancel)
    
    def update_progress(self, value: int, message: Optional[str] = None):
        """Update progress value and message"""
        self.progress['value'] = value
        self.progress_text.config(text=f"{value}%")
        
        if message:
            self.message_label.config(text=message)
        
        self.update()
    
    def cancel(self):
        """Cancel the operation"""
        self.cancelled = True
        self.destroy()
    
    def is_cancelled(self) -> bool:
        """Check if operation was cancelled"""
        return self.cancelled

class FileInfoWidget(ttk_boot.LabelFrame):
    """Widget to display file information"""
    
    def __init__(self, parent, title="File Information", **kwargs):
        super().__init__(parent, text=title, padding=10, **kwargs)
        
        self.info_vars = {}
        self.setup_ui()
    
    def setup_ui(self):
        """Setup file info display"""
        info_frame = ttk_boot.Frame(self)
        info_frame.pack(fill=BOTH, expand=True)
        
        # Create info labels
        self.info_labels = {}
        info_items = [
            ('file_name', 'File Name:'),
            ('sheet_name', 'Sheet Name:'),
            ('max_rows', 'Total Rows:'),
            ('max_columns', 'Total Columns:'),
            ('header_row', 'Header Row:'),
            ('data_rows', 'Data Rows:')
        ]
        
        for i, (key, label) in enumerate(info_items):
            row = i // 2
            col = (i % 2) * 2
            
            ttk_boot.Label(info_frame, text=label, font=("Arial", 9, "bold")).grid(
                row=row, column=col, sticky=W, padx=(0, 5), pady=2)
            
            self.info_vars[key] = tk.StringVar(value="N/A")
            ttk_boot.Label(info_frame, textvariable=self.info_vars[key], 
                          font=("Arial", 9)).grid(
                row=row, column=col+1, sticky=W, padx=(0, 20), pady=2)
    
    def update_info(self, info_dict: Dict[str, Any]):
        """Update file information display"""
        for key, value in info_dict.items():
            if key in self.info_vars:
                self.info_vars[key].set(str(value))

class ValidationPanel(ttk_boot.LabelFrame):
    """Panel to display validation results"""
    
    def __init__(self, parent, **kwargs):
        super().__init__(parent, text="Validation Results", padding=10, **kwargs)
        
        self.setup_ui()
    
    def setup_ui(self):
        """Setup validation panel"""
        # Status frame
        status_frame = ttk_boot.Frame(self)
        status_frame.pack(fill=X, pady=(0, 10))
        
        self.status_label = ttk_boot.Label(status_frame, text="Ready for validation", 
                                          font=("Arial", 10, "bold"))
        self.status_label.pack(side=LEFT)
        
        self.status_icon = ttk_boot.Label(status_frame, text="●", 
                                         font=("Arial", 12), foreground="gray")
        self.status_icon.pack(side=RIGHT)
        
        # Issues frame
        self.issues_frame = ttk_boot.Frame(self)
        self.issues_frame.pack(fill=BOTH, expand=True)
        
        # Issues text widget with scrollbar
        text_frame = ttk_boot.Frame(self.issues_frame)
        text_frame.pack(fill=BOTH, expand=True)
        
        self.issues_text = tk.Text(text_frame, height=6, wrap=tk.WORD, 
                                  font=("Arial", 9), state=tk.DISABLED)
        scrollbar = ttk_boot.Scrollbar(text_frame, orient="vertical", 
                                      command=self.issues_text.yview)
        
        self.issues_text.configure(yscrollcommand=scrollbar.set)
        
        self.issues_text.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)
    
    def update_validation(self, is_valid: bool, issues: List[str]):
        """Update validation display"""
        self.issues_text.config(state=tk.NORMAL)
        self.issues_text.delete(1.0, tk.END)
        
        if is_valid:
            self.status_label.config(text="✓ Validation Passed")
            self.status_icon.config(foreground="green", text="●")
            self.issues_text.insert(tk.END, "No issues found. Ready to proceed.")
        else:
            self.status_label.config(text="✗ Validation Failed")
            self.status_icon.config(foreground="red", text="●")
            
            for i, issue in enumerate(issues, 1):
                self.issues_text.insert(tk.END, f"{i}. {issue}\n")
        
        self.issues_text.config(state=tk.DISABLED)

# Utility functions for common dialogs
def show_error_dialog(parent, title: str, message: str, details: Optional[str] = None):
    """Show enhanced error dialog with optional details"""
    if details:
        # Create custom dialog with details
        dialog = tk.Toplevel(parent)
        dialog.title(title)
        dialog.geometry("500x300")
        dialog.transient(parent)
        dialog.grab_set()
        
        # Center on parent
        dialog.geometry(f"+{parent.winfo_rootx() + 50}+{parent.winfo_rooty() + 50}")
        
        main_frame = ttk_boot.Frame(dialog, padding=20)
        main_frame.pack(fill=BOTH, expand=True)
        
        # Message
        ttk_boot.Label(main_frame, text=message, font=("Arial", 10, "bold")).pack(pady=(0, 10))
        
        # Details
        details_frame = ttk_boot.LabelFrame(main_frame, text="Details", padding=10)
        details_frame.pack(fill=BOTH, expand=True, pady=(0, 10))
        
        text_widget = tk.Text(details_frame, wrap=tk.WORD, font=("Consolas", 9))
        scrollbar = ttk_boot.Scrollbar(details_frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        text_widget.insert(tk.END, details)
        text_widget.config(state=tk.DISABLED)
        
        text_widget.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)
        
        # OK button
        ttk_boot.Button(main_frame, text="OK", command=dialog.destroy, 
                       bootstyle=PRIMARY).pack()
    else:
        messagebox.showerror(title, message, parent=parent)

def show_confirmation_dialog(parent, title: str, message: str) -> bool:
    """Show confirmation dialog"""
    return messagebox.askyesno(title, message, parent=parent)