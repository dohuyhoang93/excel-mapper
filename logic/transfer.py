"""
Data transfer engine for copying data between Excel files while preserving formatting
"""
import openpyxl
from openpyxl.styles import *
from openpyxl.utils import get_column_letter
from typing import List, Dict, Any, Optional, Callable
import shutil
from pathlib import Path
import logging
from datetime import datetime
from typing import Tuple
from copy import copy

class ExcelTransferEngine:
    """Handles data transfer between Excel files with formatting preservation"""
    
    def __init__(self, source_path: str, destination_path: str):
        self.source_path = Path(source_path)
        self.destination_path = Path(destination_path)
        self.backup_path = None
        self.progress_callback = None
        
    def set_progress_callback(self, callback: Callable[[int], None]):
        """Set callback function for progress updates"""
        self.progress_callback = callback
        
    def create_backup(self) -> Path:
        """Create backup of destination file"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"{self.destination_path.stem}_backup_{timestamp}{self.destination_path.suffix}"
        self.backup_path = self.destination_path.parent / backup_name
        
        shutil.copy2(self.destination_path, self.backup_path)
        logging.info(f"Backup created: {self.backup_path}")
        return self.backup_path
        
    def restore_backup(self):
        """Restore from backup if something goes wrong"""
        if self.backup_path and self.backup_path.exists():
            shutil.copy2(self.backup_path, self.destination_path)
            logging.info("Restored from backup")
            
    def cleanup_backup(self):
        """Remove backup file after successful transfer"""
        if self.backup_path and self.backup_path.exists():
            self.backup_path.unlink()
            logging.info("Backup cleaned up")
            
    def read_source_data(self, header_row: int, sort_column: Optional[str] = None) -> Tuple[List[str], List[Dict[str, Any]]]:
        """
        Read data from source Excel file
        
        Args:
            header_row: Row number containing headers
            sort_column: Column name to sort by (optional)
            
        Returns:
            Tuple of (headers, data_rows)
        """
        workbook = openpyxl.load_workbook(self.source_path, data_only=True)
        worksheet = workbook.active
        
        try:
            # Get headers
            headers = []
            for col in range(1, worksheet.max_column + 1):
                cell = worksheet.cell(row=header_row, column=col)
                header_value = str(cell.value) if cell.value else f"Column_{col}"
                headers.append(header_value.strip())
            
            # Read data rows
            data = []
            for row in range(header_row + 1, worksheet.max_row + 1):
                row_data = {}
                has_data = False
                
                for col, header in enumerate(headers, start=1):
                    cell = worksheet.cell(row=row, column=col)
                    value = cell.value
                    
                    if value is not None:
                        has_data = True
                    
                    row_data[header] = value
                
                if has_data:
                    data.append(row_data)
            
            # Sort data if requested
            if sort_column and sort_column in headers:
                data = sorted(data, key=lambda x: str(x.get(sort_column, "")) if x.get(sort_column) is not None else "")
                logging.info(f"Data sorted by column: {sort_column}")
            
            return headers, data
            
        finally:
            workbook.close()
    
    def get_destination_column_map(self, header_row: int) -> Dict[str, int]:
        """
        Get mapping of column names to column numbers in destination file
        
        Args:
            header_row: Row number containing headers
            
        Returns:
            Dictionary mapping column names to column numbers
        """
        workbook = openpyxl.load_workbook(self.destination_path)
        worksheet = workbook.active
        
        try:
            column_map = {}
            merged_ranges = worksheet.merged_cells.ranges
            
            for col in range(1, worksheet.max_column + 1):
                cell = worksheet.cell(row=header_row, column=col)
                cell_value = cell.value
                
                # Handle merged cells
                if cell_value is None:
                    for merged_range in merged_ranges:
                        if cell.coordinate in merged_range:
                            top_left = worksheet.cell(merged_range.min_row, merged_range.min_col)
                            cell_value = top_left.value
                            break
                
                if cell_value:
                    column_map[str(cell_value).strip()] = col
            
            return column_map
            
        finally:
            workbook.close()
    
    def write_data_to_destination(self, data: List[Dict[str, Any]], 
                                mappings: Dict[str, str], 
                                dest_header_row: int,
                                preserve_formulas: bool = True) -> int:
        """
        Write data to destination file preserving formatting
        
        Args:
            data: List of data rows to write
            mappings: Dictionary mapping source columns to destination columns
            dest_header_row: Row number of headers in destination
            preserve_formulas: Whether to skip cells with formulas
            
        Returns:
            Number of rows written
        """
        workbook = openpyxl.load_workbook(self.destination_path)
        worksheet = workbook.active
        
        try:
            # Get destination column mapping
            dest_columns = self.get_destination_column_map_from_worksheet(worksheet, dest_header_row)
            
            # Find data start row
            data_start_row = dest_header_row + 1
            
            # Clear existing data (but preserve formatting)
            self._clear_data_area(worksheet, data_start_row, dest_columns, preserve_formulas)
            
            # Write new data
            rows_written = 0
            total_rows = len(data)
            
            for i, row_data in enumerate(data):
                current_row = data_start_row + i
                
                # Update progress
                if self.progress_callback and i % 10 == 0:
                    progress = int((i / total_rows) * 100)
                    self.progress_callback(progress)
                
                # Write mapped columns
                for source_col, dest_col in mappings.items():
                    if dest_col in dest_columns:
                        dest_col_num = dest_columns[dest_col]
                        source_value = row_data.get(source_col)
                        
                        if source_value is not None:
                            dest_cell = worksheet.cell(row=current_row, column=dest_col_num)
                            
                            # Skip formula cells if requested
                            if preserve_formulas and dest_cell.data_type == 'f':
                                logging.warning(f"Skipping formula cell at {dest_cell.coordinate}")
                                continue
                            
                            # Copy value while preserving cell formatting
                            self._copy_value_preserve_format(dest_cell, source_value)
                
                rows_written += 1
            
            # Save workbook
            workbook.save(self.destination_path)
            logging.info(f"Successfully wrote {rows_written} rows to destination")
            
            return rows_written
            
        finally:
            workbook.close()
    
    def get_destination_column_map_from_worksheet(self, worksheet, header_row: int) -> Dict[str, int]:
        """Get column mapping from worksheet object"""
        column_map = {}
        merged_ranges = worksheet.merged_cells.ranges
        
        for col in range(1, worksheet.max_column + 1):
            cell = worksheet.cell(row=header_row, column=col)
            cell_value = cell.value
            
            # Handle merged cells
            if cell_value is None:
                for merged_range in merged_ranges:
                    if cell.coordinate in merged_range:
                        top_left = worksheet.cell(merged_range.min_row, merged_range.min_col)
                        cell_value = top_left.value
                        break
            
            if cell_value:
                column_map[str(cell_value).strip()] = col
        
        return column_map
    
    def _clear_data_area(self, worksheet, start_row: int, dest_columns: Dict[str, int], preserve_formulas: bool):
        """Clear data area while preserving formatting"""
        max_row = worksheet.max_row
        
        for row in range(start_row, max_row + 1):
            for col_name, col_num in dest_columns.items():
                cell = worksheet.cell(row=row, column=col_num)
                
                # Only clear non-formula cells if preserve_formulas is True
                if not preserve_formulas or cell.data_type != 'f':
                    cell.value = None
    
    def _copy_value_preserve_format(self, dest_cell, source_value):
        """Copy value to destination cell while preserving formatting"""
        # Store original formatting
        original_font = copy(dest_cell.font) if dest_cell.font else None
        original_border = copy(dest_cell.border) if dest_cell.border else None
        original_fill = copy(dest_cell.fill) if dest_cell.fill else None
        original_alignment = copy(dest_cell.alignment) if dest_cell.alignment else None
        original_number_format = dest_cell.number_format
        
        # Set the value
        dest_cell.value = source_value
        
        # Restore formatting
        if original_font:
            dest_cell.font = original_font
        if original_border:
            dest_cell.border = original_border
        if original_fill:
            dest_cell.fill = original_fill
        if original_alignment:
            dest_cell.alignment = original_alignment
        if original_number_format:
            dest_cell.number_format = original_number_format
    
    def transfer_data(self, source_header_row: int, dest_header_row: int, 
                     mappings: Dict[str, str], sort_column: Optional[str] = None) -> Dict[str, Any]:
        """
        Complete data transfer operation
        
        Args:
            source_header_row: Header row in source file
            dest_header_row: Header row in destination file
            mappings: Column mappings
            sort_column: Column to sort by
            
        Returns:
            Dictionary with transfer results
        """
        try:
            # Create backup
            self.create_backup()
            
            # Read source data
            if self.progress_callback:
                self.progress_callback(10)
            
            headers, data = self.read_source_data(source_header_row, sort_column)
            
            if self.progress_callback:
                self.progress_callback(30)
            
            # Validate mappings
            validation_errors = self._validate_mappings(mappings, headers, dest_header_row)
            if validation_errors:
                raise ValueError(f"Mapping validation failed: {'; '.join(validation_errors)}")
            
            if self.progress_callback:
                self.progress_callback(40)
            
            # Write to destination
            rows_written = self.write_data_to_destination(data, mappings, dest_header_row)
            
            if self.progress_callback:
                self.progress_callback(100)
            
            # Cleanup backup on success
            self.cleanup_backup()
            
            return {
                'success': True,
                'rows_written': rows_written,
                'source_rows': len(data),
                'mappings_used': len(mappings),
                'sorted_by': sort_column
            }
            
        except Exception as e:
            # Restore backup on failure
            self.restore_backup()
            logging.error(f"Transfer failed: {str(e)}")
            raise
    
    def _validate_mappings(self, mappings: Dict[str, str], source_headers: List[str], dest_header_row: int) -> List[str]:
        """Validate that mappings are correct"""
        errors = []
        
        # Check that all source columns exist
        for source_col in mappings.keys():
            if source_col not in source_headers:
                errors.append(f"Source column '{source_col}' not found")
        
        # Check that all destination columns exist
        dest_columns = self.get_destination_column_map(dest_header_row)
        for dest_col in mappings.values():
            if dest_col not in dest_columns:
                errors.append(f"Destination column '{dest_col}' not found")
        
        # Check for duplicate destination mappings
        dest_values = list(mappings.values())
        duplicates = [col for col in set(dest_values) if dest_values.count(col) > 1]
        if duplicates:
            errors.append(f"Duplicate destination columns: {', '.join(duplicates)}")
        
        return errors
    
    def preview_transfer(self, source_header_row: int, dest_header_row: int, 
                        mappings: Dict[str, str], preview_rows: int = 5) -> Dict[str, Any]:
        """
        Preview what the transfer would look like
        
        Args:
            source_header_row: Header row in source file
            dest_header_row: Header row in destination file  
            mappings: Column mappings
            preview_rows: Number of rows to preview
            
        Returns:
            Preview data dictionary
        """
        try:
            # Read limited source data
            headers, data = self.read_source_data(source_header_row)
            preview_data = data[:preview_rows]
            
            # Get destination columns
            dest_columns = self.get_destination_column_map(dest_header_row)
            
            # Validate mappings
            validation_errors = self._validate_mappings(mappings, headers, dest_header_row)
            
            return {
                'source_headers': headers,
                'destination_headers': list(dest_columns.keys()),
                'preview_data': preview_data,
                'total_source_rows': len(data),
                'mappings': mappings,
                'validation_errors': validation_errors,
                'is_valid': len(validation_errors) == 0
            }
            
        except Exception as e:
            logging.error(f"Preview failed: {str(e)}")
            return {
                'error': str(e),
                'is_valid': False
            }
            