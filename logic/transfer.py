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
from copy import copy

def _sanitize_sheet_name(name: str) -> str:
    """Sanitizes a string to be a valid Excel sheet name."""
    if not name:
        return "Untitled"
    name = str(name) # Ensure name is a string
    # Remove invalid characters
    name = re.sub(r'[\\/*?:[\\]', '', name)
    # Truncate to 31 characters (Excel's limit)
    return name[:31]

def parse_skip_rows_string(skip_rows_str: str) -> Set[int]:
    """
    Parses a user-provided string of rows to skip into a set of integers.
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
        if self.progress_callback:
            self.progress_callback(value, message)

    def _get_writable_cell(self, worksheet, row_idx, col_idx):
        """Gets the top-left cell of a potentially merged range."""
        cell = worksheet.cell(row=row_idx, column=col_idx)
        if isinstance(cell, MergedCell):
            for merged_range in worksheet.merged_cells.ranges:
                if cell.coordinate in merged_range:
                    return worksheet.cell(row=merged_range.min_row, column=merged_range.min_col)
        return cell

    def run_transfer(self):
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
                    if isinstance(value, str):
                        value = value.strip() # ISSUE 1 FIX: Strip whitespace
                    if value is not None and str(value).strip() != '':
                        has_data = True
                    row_data[header_name] = value
                if has_data:
                    data.append(row_data)
            return data
        finally:
            if workbook:
                workbook.close()

    def _group_data(self, source_data: List[Dict[str, Any]]) -> Dict[str, List[Dict[str, Any]]]:
        grouped = {}
        for row in source_data:
            key = str(row.get(self.group_by_column, "Uncategorized"))
            if key not in grouped:
                grouped[key] = []
            grouped[key].append(row)
        return grouped

    def _write_group_identifier(self, worksheet, group_name: str, group_by_header: str):
        """Finds the cell with the group-by header and writes the group name next to it."""
        # Search in a reasonable range
        for row in worksheet.iter_rows(min_row=1, max_row=50, max_col=50):
            for cell in row:
                cell_value = str(cell.value).strip() if cell.value is not None else ""
                if cell_value.lower() == group_by_header.lower():
                    try:
                        # ISSUE 3 FIX: Use _get_writable_cell to handle merged cells
                        target_cell = self._get_writable_cell(worksheet, cell.row, cell.column + 1)
                        target_cell.value = group_name
                        logging.info(f"Wrote group name '{group_name}' to cell {target_cell.coordinate}")
                        return
                    except Exception as e:
                        logging.error(f"Error writing group identifier for '{group_name}': {e}")
                        return

    def _write_grouped_data(self, grouped_data: Dict[str, List[Dict[str, Any]]]):
        workbook = None
        try:
            workbook = openpyxl.load_workbook(self.dest_path)
            
            if self.master_sheet_name not in workbook.sheetnames:
                raise ValueError(f"Master sheet '{self.master_sheet_name}' not found.")
            master_sheet = workbook[self.master_sheet_name]

            # Remove default sheet if it exists and is not the master
            if "Sheet" in workbook.sheetnames and len(workbook.sheetnames) > 1 and workbook["Sheet"] != master_sheet:
                workbook.remove(workbook["Sheet"])

            total_groups = len(grouped_data)
            created_sheets = []
            for i, (group_name, data_rows) in enumerate(grouped_data.items()):
                self._update_progress(25 + int((i / total_groups) * 70), f"Processing group {i+1}/{total_groups}: {group_name}")

                new_sheet_name = _sanitize_sheet_name(group_name)
                if new_sheet_name in workbook.sheetnames:
                    new_sheet_name = _sanitize_sheet_name(f"{group_name}_{i+1}")
                
                new_sheet = workbook.copy_worksheet(master_sheet)
                new_sheet.title = new_sheet_name
                created_sheets.append(new_sheet_name)

                self._write_group_identifier(new_sheet, group_name, self.group_by_column)

                skipped_rows = parse_skip_rows_string(self.dest_skip_rows_str)
                rows_to_write = len(data_rows)
                current_end_row = self.dest_write_end_row

                if self.dest_write_end_row > 0:
                    available_rows = sum(1 for r in range(self.dest_write_start_row, self.dest_write_end_row + 1) if r not in skipped_rows)
                    if rows_to_write > available_rows:
                        rows_to_insert = rows_to_write - available_rows
                        insertion_point = self.dest_write_end_row
                        new_sheet.insert_rows(insertion_point, amount=rows_to_insert)
                        logging.info(f"Inserted {rows_to_insert} rows into '{new_sheet_name}'.")
                        current_end_row += rows_to_insert

                        # ISSUE 2 FIX: Copy styles to new rows
                        style_source_row = insertion_point - 1 if insertion_point > 1 else 1
                        for row_idx in range(insertion_point, insertion_point + rows_to_insert):
                            for col_idx in range(1, new_sheet.max_column + 1):
                                source_cell = new_sheet.cell(row=style_source_row, column=col_idx)
                                new_cell = new_sheet.cell(row=row_idx, column=col_idx)
                                if source_cell.has_style:
                                    new_cell.font = copy(source_cell.font)
                                    new_cell.border = copy(source_cell.border)
                                    new_cell.fill = copy(source_cell.fill)
                                    new_cell.number_format = source_cell.number_format
                                    new_cell.protection = copy(source_cell.protection)
                                    new_cell.alignment = copy(source_cell.alignment)
                
                self._write_single_sheet(new_sheet, data_rows, skipped_rows, current_end_row)

            if self.master_sheet_name in workbook.sheetnames and len(workbook.sheetnames) > 1:
                del workbook[self.master_sheet_name]
            
            # ISSUE 4 FIX: Set active sheet before saving
            if created_sheets and created_sheets[0] in workbook.sheetnames:
                workbook.active = workbook[created_sheets[0]]

            workbook.save(self.dest_path)
        finally:
            if workbook:
                workbook.close()

    def _write_single_sheet(self, worksheet, data_rows, skipped_rows, current_end_row):
        current_write_row = self.dest_write_start_row
        EXCEL_MAX_ROW = 1048576

        for row_data in data_rows:
            while True:
                if current_end_row > 0 and current_write_row > current_end_row:
                    logging.warning(f"Reached end of write zone on sheet '{worksheet.title}'.")
                    return
                if current_write_row > EXCEL_MAX_ROW:
                    raise RuntimeError(f"Reached max Excel row limit on sheet '{worksheet.title}'.")
                
                is_invalid_row = current_write_row in skipped_rows
                if not is_invalid_row and self.respect_cell_protection and worksheet.protection.sheet:
                    if any(self._get_writable_cell(worksheet, current_write_row, c).protection.locked for c in self.dest_columns.values()):
                        is_invalid_row = True
                
                if not is_invalid_row:
                    break
                current_write_row += 1

            for source_col, dest_col in self.mappings.items():
                if dest_col in self.dest_columns:
                    dest_col_num = self.dest_columns[dest_col]
                    cell_to_write = self._get_writable_cell(worksheet, current_write_row, dest_col_num)
                    if cell_to_write.row >= current_write_row and not (self.respect_formulas and cell_to_write.data_type == 'f'):
                        cell_to_write.value = row_data.get(source_col)
            current_write_row += 1
