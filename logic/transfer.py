"""
Data transfer engine for copying data between Excel files while preserving formatting.
This module encapsulates the core business logic of the data transfer process.
"""
import openpyxl
import shutil
from pathlib import Path
import logging
from typing import List, Dict, Any, Optional, Callable, Set
from openpyxl.cell.cell import MergedCell

def parse_skip_rows_string(skip_rows_str: str) -> Set[int]:
    """
    Parses a user-provided string of rows to skip into a set of integers.
    This is a public utility function that can be used by other modules.

    Args:
        skip_rows_str: A string like "15, 22, 30-35".

    Returns:
        A set of integers representing the rows to be skipped.
    """
    skipped_rows = set()
    if not skip_rows_str:
        return skipped_rows
    for part in skip_rows_str.split(','):
        part = part.strip()
        if not part:
            continue
        if '-' in part:
            try:
                start, end = map(int, part.split('-'))
                if start <= end:
                    skipped_rows.update(range(start, end + 1))
            except ValueError:
                logging.warning(f"Could not parse range in skip_rows: {part}")
        else:
            try:
                skipped_rows.add(int(part))
            except ValueError:
                logging.warning(f"Could not parse number in skip_rows: {part}")
    return skipped_rows

"""
Data transfer engine for copying data between Excel files while preserving formatting.
This module encapsulates the core business logic of the data transfer process.
"""
import openpyxl
import shutil
from pathlib import Path
import logging
from typing import List, Dict, Any, Optional, Callable, Set
from openpyxl.cell.cell import MergedCell
import re

def _sanitize_sheet_name(name: str) -> str:
    """Sanitizes a string to be a valid Excel sheet name."""
    if not name:
        return "Untitled"
    # Remove invalid characters
    name = re.sub(r'[\\/*?:[\\]]', '', name)
    # Truncate to 31 characters (Excel's limit)
    return name[:31]

def parse_skip_rows_string(skip_rows_str: str) -> Set[int]:
    """
    Parses a user-provided string of rows to skip into a set of integers.
    This is a public utility function that can be used by other modules.

    Args:
        skip_rows_str: A string like "15, 22, 30-35".

    Returns:
        A set of integers representing the rows to be skipped.
    """
    skipped_rows = set()
    if not skip_rows_str:
        return skipped_rows
    for part in skip_rows_str.split(','):
        part = part.strip()
        if not part:
            continue
        if '-' in part:
            try:
                start, end = map(int, part.split('-'))
                if start <= end:
                    skipped_rows.update(range(start, end + 1))
            except ValueError:
                logging.warning(f"Could not parse range in skip_rows: {part}")
        else:
            try:
                skipped_rows.add(int(part))
            except ValueError:
                logging.warning(f"Could not parse number in skip_rows: {part}")
    return skipped_rows

class ExcelTransferEngine:
    """
    Handles the entire data transfer process from a source to a destination Excel file.
    It takes a comprehensive settings dictionary to control its behavior.
    """
    def __init__(self, settings: Dict[str, Any], progress_callback: Optional[Callable[[int, str], None]] = None):
        self.source_path = Path(settings["source_file"])
        self.dest_path = Path(settings["dest_file"])
        self.source_header_end_row = settings["source_header_end_row"]
        self.dest_header_end_row = settings["dest_header_end_row"]
        self.dest_write_start_row = settings["dest_write_start_row"]
        self.dest_write_end_row = settings["dest_write_end_row"]
        self.dest_skip_rows_str = settings["dest_skip_rows"]
        self.respect_cell_protection = settings["respect_cell_protection"]
        self.respect_formulas = settings["respect_formulas"]
        self.group_by_column = settings.get("group_by_column")
        self.master_sheet_name = settings.get("master_sheet")
        self.mappings = settings["mappings"]
        self.source_columns = settings["source_columns"]
        self.dest_columns = settings["dest_columns"]

        self.progress_callback = progress_callback
        self.backup_path = None

    def _update_progress(self, value: int, message: str):
        """Safely calls the progress callback function if it exists."""
        if self.progress_callback:
            self.progress_callback(value, message)

    def run_transfer(self):
        """
        Executes the grouped data transfer process.
        """
        if not self.group_by_column:
            raise ValueError("'Group by Column' must be selected for this operation.")
        if not self.master_sheet_name:
            raise ValueError("'Master Sheet' must be selected for this operation.")

        self.backup_path = self.dest_path.with_suffix(f'.{self.dest_path.suffix}.backup')
        try:
            shutil.copy2(self.dest_path, self.backup_path)
            logging.info(f"Backup created at {self.backup_path}")

            self._update_progress(5, "Reading source data...")
            source_data = self._read_source_data()
            if not source_data:
                raise ValueError("No data found in source file.")

            self._update_progress(15, "Grouping data...")
            grouped_data = self._group_data(source_data)
            
            self._update_progress(25, "Processing groups...")
            self._write_grouped_data(grouped_data)

            if self.backup_path.exists():
                self.backup_path.unlink()
            self._update_progress(100, "Transfer completed successfully")
            logging.info("Grouped data transfer completed successfully")

        except Exception as e:
            logging.error(f"Transfer failed: {e}", exc_info=True)
            if self.backup_path and self.backup_path.exists():
                try:
                    shutil.copy2(self.backup_path, self.dest_path)
                    self.backup_path.unlink()
                    logging.info("Restored destination file from backup.")
                except Exception as backup_e:
                    logging.error(f"CRITICAL: Failed to restore backup: {backup_e}", exc_info=True)
            raise e

    def _read_source_data(self) -> List[Dict[str, Any]]:
        """Reads all data rows from the source file."""
        # This function is reused from the old implementation
        workbook = None
        try:
            workbook = openpyxl.load_workbook(self.source_path, data_only=True)
            worksheet = workbook.active
            start_data_row = self.source_header_end_row + 1
            data = []
            for row_index in range(start_data_row, worksheet.max_row + 1):
                row_data, has_data = {}, False
                for header_name, col_index in self.source_columns.items():
                    value = worksheet.cell(row=row_index, column=col_index).value
                    if value is not None:
                        has_data = True
                    row_data[header_name] = value
                if has_data:
                    data.append(row_data)
            return data
        finally:
            if workbook:
                workbook.close()

    def _group_data(self, source_data: List[Dict[str, Any]]) -> Dict[str, List[Dict[str, Any]]]:
        """Groups source data by the selected column."""
        grouped = {}
        for row in source_data:
            key = str(row.get(self.group_by_column, "Uncategorized"))
            if key not in grouped:
                grouped[key] = []
            grouped[key].append(row)
        return grouped

    def _write_grouped_data(self, grouped_data: Dict[str, List[Dict[str, Any]]]):
        """
        Writes the grouped data to the destination file, creating a new sheet for each group.
        """
        workbook = None
        try:
            workbook = openpyxl.load_workbook(self.dest_path)
            
            if self.master_sheet_name not in workbook.sheetnames:
                raise ValueError(f"Master sheet '{self.master_sheet_name}' not found in the destination file.")
            master_sheet = workbook[self.master_sheet_name]

            total_groups = len(grouped_data)
            for i, (group_name, data_rows) in enumerate(grouped_data.items()):
                self._update_progress(25 + int((i / total_groups) * 70), f"Processing group {i+1}/{total_groups}: {group_name}")

                # 1. Create new sheet from master
                new_sheet_name = _sanitize_sheet_name(group_name)
                if new_sheet_name in workbook.sheetnames:
                    # Handle duplicate sheet names if necessary
                    new_sheet_name = _sanitize_sheet_name(f"{group_name}_{i+1}")
                
                new_sheet = workbook.copy_worksheet(master_sheet)
                new_sheet.title = new_sheet_name

                # 2. Insert rows if necessary
                skipped_rows = parse_skip_rows_string(self.dest_skip_rows_str)
                available_rows = 0
                if self.dest_write_end_row > 0:
                    for r in range(self.dest_write_start_row, self.dest_write_end_row + 1):
                        if r not in skipped_rows:
                            available_rows += 1
                
                rows_to_write = len(data_rows)
                if self.dest_write_end_row > 0 and rows_to_write > available_rows:
                    rows_to_insert = rows_to_write - available_rows
                    # Insert rows just before the end write row to preserve any footers
                    new_sheet.insert_rows(self.dest_write_end_row, amount=rows_to_insert)
                    logging.info(f"Inserted {rows_to_insert} rows into sheet '{new_sheet_name}'.")
                    # Adjust the end write row for the current sheet processing
                    current_end_row = self.dest_write_end_row + rows_to_insert
                else:
                    current_end_row = self.dest_write_end_row

                # 3. Write data to the new sheet
                self._write_single_sheet(new_sheet, data_rows, skipped_rows, current_end_row)

            # Optional: remove the master sheet after all groups are processed
            # del workbook[self.master_sheet_name]

            workbook.save(self.dest_path)
        finally:
            if workbook:
                workbook.close()

    def _write_single_sheet(self, worksheet, data_rows, skipped_rows, current_end_row):
        """Writes data rows to a single sheet, respecting write zone rules."""
        
        def get_writable_cell(row_idx, col_idx):
            cell = worksheet.cell(row=row_idx, column=col_idx)
            if isinstance(cell, MergedCell):
                for merged_range in worksheet.merged_cells.ranges:
                    if cell.coordinate in merged_range:
                        return worksheet.cell(row=merged_range.min_row, column=merged_range.min_col)
            return cell

        current_write_row = self.dest_write_start_row
        EXCEL_MAX_ROW = 1048576

        for row_data in data_rows:
            while True:
                if current_end_row > 0 and current_write_row > current_end_row:
                    logging.warning(f"Reached end of write zone on sheet '{worksheet.title}'.")
                    return

                if current_write_row > EXCEL_MAX_ROW:
                    raise RuntimeError(f"Reached maximum Excel row limit on sheet '{worksheet.title}'.")
                
                is_invalid_row = current_write_row in skipped_rows
                if not is_invalid_row and self.respect_cell_protection and worksheet.protection.sheet:
                    if any(get_writable_cell(current_write_row, c).protection.locked for c in self.dest_columns.values()):
                        is_invalid_row = True
                
                if not is_invalid_row:
                    break
                current_write_row += 1

            for source_col, dest_col in self.mappings.items():
                if dest_col in self.dest_columns:
                    dest_col_num = self.dest_columns[dest_col]
                    cell_to_write = get_writable_cell(current_write_row, dest_col_num)
                    if cell_to_write.row >= current_write_row and not (self.respect_formulas and cell_to_write.data_type == 'f'):
                        cell_to_write.value = row_data.get(source_col)
            current_write_row += 1
