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
        self.sort_column = settings.get("sort_column")
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
        Executes the entire data transfer process: backup, read, sort, write, and cleanup.
        """
        self.backup_path = self.dest_path.with_suffix(f'.{self.dest_path.suffix}.backup')
        try:
            shutil.copy2(self.dest_path, self.backup_path)
            logging.info(f"Backup created at {self.backup_path}")

            self._update_progress(10, "Reading source data...")
            source_data = self._read_source_data()
            if not source_data:
                raise ValueError("No data found in source file. Please check the file and header settings.")

            if self.sort_column:
                self._update_progress(30, "Sorting data...")
                source_data.sort(key=lambda x: (
                    x.get(self.sort_column) is None or str(x.get(self.sort_column, "")).strip() == "",
                    str(x.get(self.sort_column, ""))
                ))

            self._update_progress(50, "Writing to destination...")
            self._write_to_destination(source_data)

            if self.backup_path.exists():
                self.backup_path.unlink()
            self._update_progress(100, "Transfer completed successfully")
            logging.info("Data transfer completed successfully")

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
        """Reads all data rows from the source file based on the source column definitions."""
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

    def _write_to_destination(self, source_data: List[Dict[str, Any]]):
        """Writes the processed data to the destination file, respecting all write zone rules."""
        workbook = None
        try:
            workbook = openpyxl.load_workbook(self.dest_path)
            worksheet = workbook.active
            
            skipped_rows = parse_skip_rows_string(self.dest_skip_rows_str)
            
            if self.dest_write_start_row <= self.dest_header_end_row:
                raise ValueError("Start Write Row must be after the destination header rows.")

            def get_writable_cell(row_idx, col_idx):
                cell = worksheet.cell(row=row_idx, column=col_idx)
                if isinstance(cell, MergedCell):
                    for merged_range in worksheet.merged_cells.ranges:
                        if cell.coordinate in merged_range:
                            return worksheet.cell(row=merged_range.min_row, column=merged_range.min_col)
                return cell

            clear_until_row = self.dest_write_end_row if self.dest_write_end_row > 0 else worksheet.max_row + 50
            cleared_anchors = set()
            for row_to_clear in range(self.dest_write_start_row, clear_until_row + 1):
                if row_to_clear in skipped_rows:
                    continue
                for dest_col_num in self.dest_columns.values():
                    anchor_cell = get_writable_cell(row_to_clear, dest_col_num)
                    if (anchor_cell.row >= self.dest_write_start_row and 
                            anchor_cell.coordinate not in cleared_anchors and 
                            anchor_cell.row not in skipped_rows):
                        if not (self.respect_formulas and anchor_cell.data_type == 'f'):
                            anchor_cell.value = None
                        cleared_anchors.add(anchor_cell.coordinate)

            current_write_row = self.dest_write_start_row
            EXCEL_MAX_ROW = 1048576
            total_source_rows = len(source_data)

            for i, row_data in enumerate(source_data):
                while True:
                    if self.dest_write_end_row > 0 and current_write_row > self.dest_write_end_row:
                        logging.warning(f"Reached end of write zone (row {self.dest_write_end_row}). Stopping data transfer.")
                        workbook.save(self.dest_path)
                        return

                    if current_write_row > EXCEL_MAX_ROW:
                        logging.error(f"Reached absolute maximum Excel row limit ({EXCEL_MAX_ROW}). Stopping.")
                        workbook.save(self.dest_path)
                        raise RuntimeError(f"Reached maximum Excel row limit ({EXCEL_MAX_ROW}).")
                    
                    is_invalid_row = current_write_row in skipped_rows
                    if not is_invalid_row and self.respect_cell_protection and worksheet.protection.sheet:
                        for dest_col_num in self.dest_columns.values():
                            if get_writable_cell(current_write_row, dest_col_num).protection.locked:
                                is_invalid_row = True
                                break
                    
                    if not is_invalid_row:
                        break
                    current_write_row += 1

                progress_value = 50 + int((i / total_source_rows) * 45)
                self._update_progress(progress_value, f"Writing row {i+1}/{total_source_rows}")

                for source_col, dest_col in self.mappings.items():
                    if dest_col in self.dest_columns:
                        dest_col_num = self.dest_columns[dest_col]
                        cell_to_write = get_writable_cell(current_write_row, dest_col_num)
                        if cell_to_write.row >= current_write_row and not (self.respect_formulas and cell_to_write.data_type == 'f'):
                            cell_to_write.value = row_data.get(source_col)
                current_write_row += 1
            
            workbook.save(self.dest_path)
        finally:
            if workbook:
                workbook.close()
