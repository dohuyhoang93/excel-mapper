# Configuration settings for Excel Data Mapper
import os
from pathlib import Path

class Config:
    """Application configuration settings"""
    
    # Application settings
    APP_NAME = "Excel Data Mapper"
    APP_VERSION = "1.0.0"
    
    # Default theme
    DEFAULT_THEME = "flatly"
    DARK_THEME = "superhero"
    
    # Default header rows
    DEFAULT_SOURCE_HEADER_ROW = 1
    DEFAULT_DEST_HEADER_ROW = 9  # Based on destination file structure
    
    # File settings
    SUPPORTED_EXTENSIONS = [".xlsx", ".xls"]
    CONFIG_DIR = "configs"
    LOG_FILE = "app.log"
    BACKUP_SUFFIX = ".backup"
    
    # GUI settings
    WINDOW_SIZE = "900x700"
    MIN_WINDOW_SIZE = (800, 600)
    
    # Progress settings
    PROGRESS_UPDATE_FREQUENCY = 10  # Update progress every N rows
    
    # Validation settings
    MAX_FILE_PATH_LENGTH = 255
    
    # Logging settings
    LOG_FORMAT = '%(asctime)s - %(levelname)s - %(message)s'
    LOG_LEVEL = "INFO"
    
    # Default column mappings for common cases
    COMMON_MAPPINGS = {
        'content': ['Contents', 'Content', 'Description'],
        'purpose': ['Purpose', 'Reason', 'Use'],
        'amount': ['Amount', 'Value', 'Total'],
        'vat': ['VAT', 'Tax', 'VAT rate'],
        'currency': ['Currency', 'Curr', 'CCY'],
        'date': ['Trading date', 'Date', 'Transaction date'],
        'number': ['No.', 'Number', 'ID'],
        'code': ['Code Reference', 'Code', 'Reference'],
        'subtotal': ['Sub total', 'Subtotal', 'Total']
    }
    
    @classmethod
    def get_config_dir(cls):
        """Get configuration directory path"""
        config_path = Path(cls.CONFIG_DIR)
        config_path.mkdir(exist_ok=True)
        return config_path
    
    @classmethod
    def get_icon_path(cls):
        """Get icon file path for different environments"""
        import sys
        
        if getattr(sys, 'frozen', False):
            # PyInstaller bundle
            base_path = sys._MEIPASS
        else:
            # Development
            base_path = Path(__file__).parent
        
        return Path(base_path) / "icon.ico"