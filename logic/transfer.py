"""
Data transfer engine for copying data between Excel files while preserving formatting.
This module implements a 'constructive' approach: it builds a new workbook from scratch
to avoid the limitations of modifying an existing file, especially with complex templates.
"""
import openpyxl
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
import shutil
from pathlib import Path
import logging
from typing import List, Dict, Any, Optional, Callable, Set
from openpyxl.utils import get_column_letter, range_boundaries, column_index_from_string
from openpyxl.utils.cell import coordinate_from_string
from openpyxl.formula.translate import Translator
from openpyxl.cell.cell import MergedCell
import re
from copy import copy

# Dedicated logger for transfer debugging
transfer_logger = logging.getLogger('transfer_debug')
if not transfer_logger.handlers:
    transfer_logger.setLevel(logging.DEBUG)
    fh = logging.FileHandler('log.txt', mode='w', encoding='utf-8')
    fh.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - [%(funcName)s]: %(message)s')
    fh.setFormatter(formatter)
    transfer_logger.addHandler(fh)
    transfer_logger.propagate = False

def _sanitize_sheet_name(name: str) -> str:
    if not name:
        return "Untitled"
    name = str(name)
    name = re.sub(r'[\\/*?:[\\]]', '', name)
    return name[:31]

def parse_skip_rows_string(skip_rows_str: str) -> Set[int]:
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
    def __init__(self, settings: Dict[str, Any], progress_callback: Optional[Callable[[int, str], None]] = None):
        self.source_path = Path(settings["source_file"])
        self.dest_path = Path(settings["dest_file"])
        self.source_header_end_row = settings["source_header_end_row"]
        self.dest_header_end_row = settings["dest_header_end_row"]
        self.dest_write_start_row = settings["dest_write_start_row"]
        self.dest_write_end_row = settings["dest_write_end_row"]
        self.group_by_column = settings.get("group_by_column")
        self.source_sheet_name = settings.get("source_sheet")
        self.master_sheet_name = settings.get("master_sheet")
        self.mappings = settings.get("mappings", {})
        self.source_columns = settings.get("source_columns", {})
        self.dest_columns = settings.get("dest_columns", {})
        self.single_value_mappings = settings.get("single_value_mapping", {})
        self.limit_columns = settings.get("limit_columns", False)
        self.template_max_col = 0 # Will be calculated later
        self.progress_callback = progress_callback

    def _update_progress(self, value: int, message: str):
        if self.progress_callback:
            self.progress_callback(value, message)

    def _get_writable_cell(self, worksheet, row_idx, col_idx):
        cell = worksheet.cell(row=row_idx, column=col_idx)
        if isinstance(cell, MergedCell):
            for merged_range in worksheet.merged_cells.ranges:
                if cell.coordinate in merged_range:
                    return worksheet.cell(row=merged_range.min_row, column=merged_range.min_col)
        return cell

    def run_transfer(self):
        transfer_logger.info("--- Starting new transfer process (Constructive Method) ---")
        if not self.group_by_column:
            raise ValueError("'Group by Column' must be selected for this operation.")
        if not self.master_sheet_name:
            raise ValueError("'Master Sheet' must be selected for this operation.")

        self._update_progress(5, "Reading source data...")
        source_data = self._read_source_data()
        if not source_data:
            raise ValueError("No data found in source file.")

        self._update_progress(15, "Grouping data...")
        grouped_data = self._group_data(source_data)
        
        self._update_progress(20, "Loading master template...")
        wb_template_vals = openpyxl.load_workbook(self.dest_path, data_only=True)
        wb_template_formulas = openpyxl.load_workbook(self.dest_path, data_only=False)

        try:
            if self.master_sheet_name not in wb_template_vals.sheetnames:
                raise ValueError(f"Master sheet '{self.master_sheet_name}' not found.")

            master_sheet_vals = wb_template_vals[self.master_sheet_name]
            master_sheet_formulas = wb_template_formulas[self.master_sheet_name]

            # Calculate the column limit based on user setting
            if self.limit_columns:
                self.template_max_col = self._get_template_max_column(master_sheet_vals)
                transfer_logger.info(f"Column optimization enabled. Detected last used column as {self.template_max_col}")
            else:
                self.template_max_col = master_sheet_vals.max_column
                transfer_logger.info(f"Column optimization disabled. Using sheet max_column: {self.template_max_col}")

            output_wb = Workbook()
            if output_wb.active:
                output_wb.remove(output_wb.active)

            total_groups = len(grouped_data)
            for i, (group_name, data_rows) in enumerate(grouped_data.items()):
                self._update_progress(25 + int((i / total_groups) * 70), f"Processing group {i+1}/{total_groups}: {group_name}")
                
                new_sheet_name = _sanitize_sheet_name(group_name)
                if new_sheet_name in output_wb.sheetnames:
                    new_sheet_name = _sanitize_sheet_name(f"{group_name}_{i+1}")
                
                new_sheet = output_wb.create_sheet(title=new_sheet_name)

                last_header_row = self._copy_range(master_sheet_vals, master_sheet_formulas, new_sheet, 1, self.dest_header_end_row, 1)
                last_data_row = self._write_data_rows(new_sheet, master_sheet_vals, master_sheet_formulas, data_rows, last_header_row + 1)
                
                footer_start_row = self.dest_write_end_row + 1 if self.dest_write_end_row > 0 else 0
                if footer_start_row > 0:
                    self._copy_range(master_sheet_vals, master_sheet_formulas, new_sheet, footer_start_row, master_sheet_vals.max_row, last_data_row + 1)

                self._write_group_identifier(new_sheet, group_name, self.group_by_column)
                if data_rows:
                    self._write_single_values(new_sheet, data_rows[0])

            output_filename = self.dest_path.with_name(f"{self.dest_path.stem}-output{self.dest_path.suffix}")
            output_wb.save(output_filename)
            transfer_logger.info(f"Successfully created new file: {output_filename}")

        finally:
            wb_template_vals.close()
            wb_template_formulas.close()
            self._update_progress(100, "Transfer completed successfully")

    def _copy_range(self, source_sheet_vals: Worksheet, source_sheet_formulas: Worksheet, dest_sheet: Worksheet, min_row: int, max_row: int, dest_start_row: int) -> int:
        transfer_logger.info(f"Copying range from {source_sheet_vals.title}:{min_row}-{max_row} to {dest_sheet.title}:{dest_start_row}")
        
        for col, dim in source_sheet_vals.column_dimensions.items():
            dest_sheet.column_dimensions[col] = copy(dim)

        for r_idx, row in enumerate(source_sheet_vals.iter_rows(min_row=min_row, max_row=max_row, max_col=self.template_max_col)):
            dest_row_idx = dest_start_row + r_idx
            if row[0].row in source_sheet_vals.row_dimensions:
                dest_sheet.row_dimensions[dest_row_idx] = copy(source_sheet_vals.row_dimensions[row[0].row])

            for c_idx, source_cell in enumerate(row):
                dest_cell = dest_sheet.cell(row=dest_row_idx, column=c_idx + 1)
                if source_cell.has_style:
                    dest_cell.font = copy(source_cell.font); dest_cell.border = copy(source_cell.border); dest_cell.fill = copy(source_cell.fill); dest_cell.number_format = source_cell.number_format; dest_cell.protection = copy(source_cell.protection); dest_cell.alignment = copy(source_cell.alignment)
                
                formula = source_sheet_formulas.cell(row=source_cell.row, column=source_cell.column).value
                if isinstance(formula, str) and formula.startswith('='):
                    try:
                        translator = Translator(formula, origin=source_cell.coordinate)
                        dest_cell.value = translator.translate_formula(dest_cell.coordinate)
                    except Exception as e:
                        transfer_logger.warning(f"Could not translate formula '{formula}' from {source_cell.coordinate}. Writing as is. Error: {e}")
                        dest_cell.value = source_cell.value
                else:
                    dest_cell.value = source_cell.value

        for mc_range in source_sheet_vals.merged_cells.ranges:
            if mc_range.min_row >= min_row and mc_range.max_row <= max_row:
                offset = dest_start_row - min_row
                new_mc_coord = f"{get_column_letter(mc_range.min_col)}{mc_range.min_row + offset}:{get_column_letter(mc_range.max_col)}{mc_range.max_row + offset}"
                dest_sheet.merge_cells(new_mc_coord)
        
        return dest_start_row + (max_row - min_row)

    def _write_data_rows(self, dest_sheet: Worksheet, master_sheet_vals: Worksheet, master_sheet_formulas: Worksheet, data_rows: List[Dict[str, Any]], start_row: int) -> int:
        transfer_logger.info(f"Writing {len(data_rows)} data rows, starting at row {start_row}")
        template_row_idx = self.dest_write_start_row
        current_write_row = start_row

        for row_data in data_rows:
            # Step 1: Stamp the entire template row using the robust _copy_range function.
            # This copies styles, static values, and formulas from the template data row.
            self._copy_range(master_sheet_vals, master_sheet_formulas, dest_sheet, 
                             min_row=template_row_idx, max_row=template_row_idx, 
                             dest_start_row=current_write_row)

            # Step 2: Overwrite the stamped row with the actual mapped data.
            for source_col, dest_col in self.mappings.items():
                if dest_col in self.dest_columns:
                    dest_col_num = self.dest_columns[dest_col]
                    # Get the cell in the newly created row to write to.
                    cell_to_write = self._get_writable_cell(dest_sheet, current_write_row, dest_col_num)
                    cell_to_write.value = row_data.get(source_col)
            
            current_write_row += 1
        
        return current_write_row - 1

    def _get_template_max_column(self, worksheet: Worksheet) -> int:
        """Scans a worksheet to find the last column that contains data or has a style."""
        max_col = 0
        # Scan all rows to find the rightmost column with any content
        for row in worksheet.iter_rows():
            for cell in row:
                # A cell is considered "used" if it has a value or a non-default style.
                # openpyxl's has_style is True if font, border, fill, etc. are not default.
                if cell.value is not None or cell.has_style:
                    if cell.column > max_col:
                        max_col = cell.column
        
        # As a fallback, if the sheet is completely empty, use the worksheet's property
        return max_col if max_col > 0 else worksheet.max_column

    def _read_source_data(self) -> List[Dict[str, Any]]:
        workbook = None
        try:
            workbook = openpyxl.load_workbook(self.source_path, data_only=True)
            if self.source_sheet_name and self.source_sheet_name in workbook.sheetnames:
                worksheet = workbook[self.source_sheet_name]
            else:
                worksheet = workbook.active
            
            start_data_row = self.source_header_end_row + 1
            data = []
            for row_index in range(start_data_row, worksheet.max_row + 1):
                row_data, has_data = {}, False
                for header_name, col_index in self.source_columns.items():
                    value = worksheet.cell(row=row_index, column=col_index).value
                    if isinstance(value, str):
                        value = value.strip()
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
        grouped = {}
        for row in source_data:
            key = str(row.get(self.group_by_column, "Uncategorized"))
            if key not in grouped:
                grouped[key] = []
            grouped[key].append(row)
        return grouped

    def _write_group_identifier(self, worksheet, group_name: str, group_by_header: str):
        for row in worksheet.iter_rows():
            for cell in row:
                anchor_cell = self._get_writable_cell(worksheet, cell.row, cell.column)
                cell_value = str(anchor_cell.value).strip() if anchor_cell.value is not None else ""
                if cell_value.lower() == group_by_header.lower():
                    max_col = anchor_cell.column
                    for merged_range in worksheet.merged_cells.ranges:
                        if anchor_cell.coordinate in merged_range:
                            max_col = merged_range.max_col
                            break
                    target_col = max_col + 1
                    target_row = anchor_cell.row
                    try:
                        target_cell_anchor = self._get_writable_cell(worksheet, target_row, target_col)
                        target_cell_anchor.value = group_name
                        return
                    except Exception as e:
                        transfer_logger.error(f"Error writing group identifier for '{group_name}': {e}")
                        return

    def _write_single_values(self, worksheet, group_first_row: Dict[str, Any]):
        if not self.single_value_mappings:
            return
        
        transfer_logger.info(f"Writing single values to sheet '{worksheet.title}'")
        for field, mapping in self.single_value_mappings.items():
            source_col = mapping.get("source_col")
            dest_cell_addr = mapping.get("dest_cell")
            
            if not source_col or not dest_cell_addr:
                continue

            value = group_first_row.get(source_col)
            
            try:
                col_str, row_idx = coordinate_from_string(dest_cell_addr)
                col_idx = column_index_from_string(col_str)

                target_cell = self._get_writable_cell(worksheet, row_idx, col_idx)
                target_cell.value = value
                transfer_logger.debug(f"Wrote single value for '{field}' (Value: {value}) to {dest_cell_addr} (actual: {target_cell.coordinate})")
            except Exception as e:
                transfer_logger.error(f"Failed to write single value for field '{field}'. Error: {e}")
